# atlas_qt.py
import sys
import traceback
import importlib
import importlib.util
import importlib.machinery
import types
from pathlib import Path
from typing import List, Tuple, Dict, Any

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QListWidget, QListWidgetItem,
    QStackedWidget, QSplitter, QVBoxLayout, QLabel, QToolBar, QAction,
    QFileDialog, QMessageBox, QStyle
)
from PyQt5.Qt import QDesktopServices

from plugin_api import ToolSpec  # shared identity!
# BaseToolPage is only needed by plugins; the shell doesn’t depend on it.


# ---------- Paths & runtime helpers ----------

def app_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent


def tools_dir() -> Path:
    return (app_root() / "tools").resolve()


def qicon_from(spec_icon: Any) -> QIcon:
    if isinstance(spec_icon, QIcon):
        return spec_icon
    if isinstance(spec_icon, str):
        # Allow relative paths inside tools/ or app root
        p = Path(spec_icon)
        if not p.is_absolute():
            candidate = tools_dir() / p
            if candidate.exists():
                p = candidate
            else:
                candidate = app_root() / p
                if candidate.exists():
                    p = candidate
        if p.exists():
            return QIcon(str(p))
    # fallback to a standard icon
    style = QApplication.instance().style() if QApplication.instance() else None
    return style.standardIcon(QStyle.SP_ComputerIcon) if style else QIcon()


# ---------- Plugin loading ----------

class PluginLoadResult:
    def __init__(self) -> None:
        self.specs: List[ToolSpec] = []
        self.errors: List[str] = []
        # Keep module names for later reload clean-up
        self.module_names: List[str] = []


def _load_module_from_path(mod_name: str, file_path: Path) -> types.ModuleType:
    """
    Load a Python module from source file explicitly (used in frozen builds).
    """
    loader = importlib.machinery.SourceFileLoader(mod_name, str(file_path))
    spec = importlib.util.spec_from_loader(mod_name, loader)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot create spec for {mod_name} at {file_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = module
    spec.loader.exec_module(module)
    return module


def discover_plugins() -> PluginLoadResult:
    """
    Discover plugin modules inside tools/ and normalize their ToolSpec.
    Accepts either a ToolSpec instance (from plugin_api) or a dict
    with keys: id, name, factory, icon?, order?.
    """
    result = PluginLoadResult()
    tdir = tools_dir()
    tdir.mkdir(parents=True, exist_ok=True)

    # Collect .py files (ignore dunders and temp files)
    py_files = [
        p for p in tdir.glob("*.py")
        if p.name not in {"__init__.py"} and not p.name.startswith("_") and "~" not in p.name
    ]
    # Deterministic order by filename; we'll resort by spec.order later
    py_files.sort(key=lambda p: p.name.lower())

    for py in py_files:
        mod_name = ""
        try:
            stem = py.stem

            if getattr(sys, "frozen", False):
                # In frozen builds, import directly from file to avoid
                # PyInstaller module graph issues.
                mod_name = f"_atlas_plugin_{stem}"
                # If reloading, drop prior module
                if mod_name in sys.modules:
                    del sys.modules[mod_name]
                mod = _load_module_from_path(mod_name, py)
            else:
                # Dev: import as package module (tools.<name>)
                pkg_name = "tools"
                mod_name = f"{pkg_name}.{stem}"
                # Ensure tools/ is importable as a package
                pkg_init = tdir / "__init__.py"
                if not pkg_init.exists():
                    pkg_init.write_text("# package marker\n", encoding="utf-8")

                # Remove prior to force fresh import on reload
                if mod_name in sys.modules:
                    del sys.modules[mod_name]
                mod = importlib.import_module(mod_name)

            result.module_names.append(mod_name)

            if not hasattr(mod, "get_tool_spec"):
                result.errors.append(f"{py.name}: missing get_tool_spec()")
                continue

            raw = mod.get_tool_spec()

            # Normalize to ToolSpec
            if isinstance(raw, ToolSpec):
                spec = raw
            elif isinstance(raw, dict):
                try:
                    spec = ToolSpec(
                        id=raw["id"],
                        name=raw["name"],
                        factory=raw["factory"],
                        icon=raw.get("icon"),
                        order=raw.get("order", 100),
                    )
                except Exception:
                    result.errors.append(f"{py.name}: dict missing required keys (id, name, factory)")
                    continue
            else:
                result.errors.append(f"{py.name}: get_tool_spec() must return ToolSpec or dict")
                continue

            # Basic validation
            if not spec.id or not spec.name or not callable(spec.factory):
                result.errors.append(f"{py.name}: invalid ToolSpec (id/name/factory)")
                continue

            result.specs.append(spec)

        except Exception as e:
            tb = "".join(traceback.format_exception_only(type(e), e)).strip()
            result.errors.append(f"{py.name}: {tb}")

    # Sort by order then name
    result.specs.sort(key=lambda s: (s.order, s.name.lower()))
    return result


# ---------- Main Window ----------

class AtlasWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ATLAS")
        self.resize(1100, 720)

        self._loaded_module_names: List[str] = []

        # UI
        self.error_banner = QLabel()
        self.error_banner.setWordWrap(True)
        self.error_banner.setVisible(False)
        self.error_banner.setStyleSheet(
            "background:#ffecec;border:1px solid #f5a9a9;color:#900;padding:6px;border-radius:6px;"
        )

        self.list = QListWidget()
        self.list.setMinimumWidth(220)
        self.stack = QStackedWidget()

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.list)
        splitter.addWidget(self.stack)
        splitter.setStretchFactor(1, 1)

        central = QWidget()
        lay = QVBoxLayout(central)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.error_banner)
        lay.addWidget(splitter)
        self.setCentralWidget(central)

        # Toolbar
        tb = QToolBar("Main")
        tb.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, tb)

        act_reload = QAction("Reload Tools", self)
        act_reload.triggered.connect(self.reload_tools)
        act_open = QAction("Open Tools Folder", self)
        act_open.triggered.connect(self.open_tools_folder)
        act_new = QAction("New Plugin…", self)
        act_new.triggered.connect(self.scaffold_new_plugin)
        act_theme = QAction("Toggle Theme", self)
        act_theme.triggered.connect(self.toggle_theme)

        tb.addAction(act_reload)
        tb.addSeparator()
        tb.addAction(act_open)
        tb.addAction(act_new)
        tb.addSeparator()
        tb.addAction(act_theme)

        # First load
        self.load_tools(first_time=True)

        # Sidebar selection
        self.list.currentRowChanged.connect(self.stack.setCurrentIndex)

    # ----- Theme -----
    _dark = False

    def toggle_theme(self):
        self._dark = not self._dark
        if self._dark:
            self.setStyleSheet("""
                * { font-size: 14px; }
                QMainWindow { background: #1b1f24; color: #e6e6e6; }
                QWidget { background: #1b1f24; color: #e6e6e6; }
                QListWidget { background: #151a1e; }
                QTextEdit, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                    background: #0f1418; color: #e6e6e6; border: 1px solid #333; border-radius: 6px; padding: 4px;
                }
                QPushButton {
                    background: #22303c; border: 1px solid #3b4b58; padding: 6px 10px; border-radius: 8px;
                }
                QToolBar { background: #151a1e; border: none; }
            """)
        else:
            self.setStyleSheet("")

    # ----- Tools -----

    def open_tools_folder(self):
        tdir = tools_dir()
        tdir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(tdir)))

    def scaffold_new_plugin(self):
        # Simple name prompt via file dialog: user picks a filename; we write a scaffold.
        tdir = tools_dir()
        tdir.mkdir(parents=True, exist_ok=True)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Create Plugin", str(tdir / "my_tool.py"), "Python Files (*.py)"
        )
        if not file_path:
            return
        dest = Path(file_path)
        if not dest.name.endswith(".py"):
            dest = dest.with_suffix(".py")

        scaffold = '''from PyQt5.QtWidgets import QLabel
from plugin_api import ToolSpec, BaseToolPage


class MyToolPage(BaseToolPage):
    def __init__(self, parent=None):
        super().__init__("My Tool", "Describe your tool here", parent)
        self.content.addWidget(QLabel("Hello from MyTool!"))


def get_tool_spec():
    return ToolSpec(
        id="my_tool",
        name="My Tool",
        factory=lambda: MyToolPage(),
        icon=None,
        order=50
    )
'''
        dest.write_text(scaffold, encoding="utf-8")
        QMessageBox.information(self, "Plugin created", f"New plugin scaffold written:\n{dest}")

    def clear_loaded_ui(self):
        # Clear sidebar and pages
        self.list.clear()
        while self.stack.count():
            w = self.stack.widget(0)
            self.stack.removeWidget(w)
            w.deleteLater()

    def load_tools(self, first_time: bool = False):
        # Discover
        result = discover_plugins()

        # Keep module names for possible cleanup
        self._loaded_module_names = result.module_names

        # Rebuild UI
        self.clear_loaded_ui()

        # Build pages
        for spec in result.specs:
            try:
                page = spec.factory()
                if not isinstance(page, QWidget):
                    raise TypeError("factory() did not return a QWidget")

                self.stack.addWidget(page)
                item = QListWidgetItem(qicon_from(spec.icon), spec.name)
                self.list.addItem(item)

            except Exception as e:
                err = "".join(traceback.format_exception_only(type(e), e)).strip()
                result.errors.append(f"{spec.name}: {err}")

        # Errors banner
        if result.errors:
            msg = "Some plugins failed to load:\n" + "\n".join(result.errors)
            self.error_banner.setText(msg)
            self.error_banner.setVisible(True)
        else:
            self.error_banner.setVisible(False)

        # Select first tool (if any)
        if self.stack.count() > 0:
            self.list.setCurrentRow(0)
        elif first_time:
            # If truly nothing, show a hint
            self.error_banner.setText(
                "No tools were found. Add .py files to the tools/ folder, then use Reload Tools."
            )
            self.error_banner.setVisible(True)

    def reload_tools(self):
        # Best-effort: drop previously loaded plugin modules to force a clean import.
        for name in list(self._loaded_module_names):
            if name in sys.modules:
                del sys.modules[name]
        self._loaded_module_names.clear()
        self.load_tools(first_time=False)


def main():
    app = QApplication(sys.argv)
    win = AtlasWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
