# atlas_qt.py
# ATLAS — PyQt5 shell with plugin-based tools (auto-load from ./tools)
# Run:  pip install PyQt5 && python atlas_qt.py

import sys
import traceback
from pathlib import Path
from types import ModuleType
from typing import Callable, Optional, List, Tuple, Dict

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPalette, QColor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QListWidget, QListWidgetItem, QStackedWidget,
    QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QToolBar, QAction, QStyleFactory,
    QMessageBox, QFileDialog, QStyle
)
import importlib
import importlib.util

APP_DIR = Path(__file__).parent
TOOLS_DIR = APP_DIR / "tools"


# ---------- Base building blocks the plugins may import ----------

class BaseToolPage(QWidget):
    """
    Optional helper for plugin UIs. It just gives a simple page wrapper with a header.
    Plugins can ignore this and build any QWidget they want.
    """
    def __init__(self, title: str, subtitle: str = "", parent: Optional[QWidget] = None):
        super().__init__(parent)
        from PyQt5.QtWidgets import QVBoxLayout, QLabel, QFrame
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        title_lbl = QLabel(title)
        title_lbl.setStyleSheet("font-size: 20px; font-weight: 600;")
        root.addWidget(title_lbl)

        if subtitle:
            sub_lbl = QLabel(subtitle)
            sub_lbl.setStyleSheet("color:#666;")
            root.addWidget(sub_lbl)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        self.content = QVBoxLayout()
        self.content.setSpacing(8)
        root.addLayout(self.content)
        root.addStretch(1)


class ToolSpec:
    """
    The simple contract a plugin must return from get_tool_spec():
      - id:      stable string id (no spaces). used internally.
      - name:    display name in the sidebar.
      - factory: callable that returns a QWidget (the tool UI).
      - icon:    optional QIcon or str path (png/ico/svg).
      - order:   optional sort key (lower shows earlier).
    """
    def __init__(
        self,
        id: str,
        name: str,
        factory: Callable[[], QWidget],
        icon: Optional[object] = None,
        order: int = 100
    ):
        self.id = id
        self.name = name
        self.factory = factory
        self.icon = icon
        self.order = order


# ---------- Plugin loader ----------

def _as_qicon(icon: Optional[object]) -> QIcon:
    if icon is None:
        return QIcon()
    if isinstance(icon, QIcon):
        return icon
    # treat as filepath
    try:
        p = Path(str(icon))
        if p.is_absolute():
            return QIcon(str(p))
        # relative to tools dir
        return QIcon(str((TOOLS_DIR / p).resolve()))
    except Exception:
        return QIcon()

def discover_plugins() -> Tuple[List[ToolSpec], List[str]]:
    """
    Import every .py file in ./tools that contains `get_tool_spec()`.
    Returns (tools, errors) so the app can still run if a plugin fails.
    """
    tools: List[ToolSpec] = []
    errors: List[str] = []

    if not TOOLS_DIR.exists():
        return tools, errors

    # Ensure we can import "tools.*"
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))

    for py in sorted(TOOLS_DIR.glob("*.py")):
        if py.name.startswith("_"):
            continue  # allow private helpers
        mod_name = f"tools.{py.stem}"
        try:
            # Fresh import (remove cached on reload)
            if mod_name in sys.modules:
                del sys.modules[mod_name]
            mod = importlib.import_module(mod_name)

            if not hasattr(mod, "get_tool_spec"):
                errors.append(f"{py.name}: missing get_tool_spec()")
                continue

            spec = mod.get_tool_spec()
            # minimal validation
            if not isinstance(spec, ToolSpec):
                errors.append(f"{py.name}: get_tool_spec() must return ToolSpec")
                continue

            # resolve icon if str
            spec.icon = _as_qicon(spec.icon)
            tools.append(spec)
        except Exception as e:
            tb = traceback.format_exc()
            errors.append(f"{py.name}: {e}\n{tb}")

    # sort by order then name
    tools.sort(key=lambda t: (t.order, t.name.lower()))
    return tools, errors


# ---------- Error Page for broken plugins ----------

class ErrorPage(BaseToolPage):
    def __init__(self, title: str, error_text: str, parent: Optional[Widget] = None):
        super().__init__(title, "Plugin failed to load", parent)
        from PyQt5.QtWidgets import QTextEdit
        box = QTextEdit()
        box.setReadOnly(True)
        box.setPlainText(error_text)
        box.setMinimumHeight(200)
        self.content.addWidget(box)


# ---------- Main Window ----------

class AtlasWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ATLAS")
        self.setMinimumSize(1120, 680)

        # Toolbar
        tb = QToolBar("Main")
        self.addToolBar(Qt.TopToolBarArea, tb)

        act_reload = QAction("Reload Tools", self)
        act_reload.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        act_reload.triggered.connect(self.reload_tools)

        act_theme = QAction("Light/Dark", self)
        act_theme.setIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        act_theme.triggered.connect(self.toggle_theme)

        act_open_tools = QAction("Open Tools Folder", self)
        act_open_tools.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))
        act_open_tools.triggered.connect(self.open_tools_folder)

        tb.addAction(act_reload)
        tb.addAction(act_theme)
        tb.addSeparator()
        tb.addAction(act_open_tools)

        # Sidebar + content
        self.nav = QListWidget()
        self.nav.setFixedWidth(240)
        self.pages = QStackedWidget()

        central = QWidget()
        lay = QHBoxLayout(central)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.addWidget(self.nav)
        lay.addWidget(self.pages, 1)
        self.setCentralWidget(central)

        self.statusBar().showMessage("Ready")
        self._dark = False

        # Load on start
        self.reload_tools()

        # Nav binding
        self.nav.currentRowChanged.connect(self.pages.setCurrentIndex)

    def open_tools_folder(self):
        if sys.platform.startswith("win"):
            import os
            os.startfile(str(TOOLS_DIR))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            import os
            os.system(f'open "{TOOLS_DIR}"')
        else:
            import os
            os.system(f'xdg-open "{TOOLS_DIR}"')

    def clear_tools(self):
        self.nav.clear()
        while self.pages.count():
            w = self.pages.widget(0)
            self.pages.removeWidget(w)
            w.deleteLater()

    def reload_tools(self):
        TOOLS_DIR.mkdir(parents=True, exist_ok=True)
        self.clear_tools()

        tools, errors = discover_plugins()

        if errors:
            # Add one "Errors" page if anything failed
            err_text = "Some plugins failed to load:\n\n" + "\n\n".join(errors)
            item = QListWidgetItem("⚠ Plugin Errors")
            self.nav.addItem(item)
            page = ErrorPage("Plugin Errors", err_text)
            self.pages.addWidget(page)

        if not tools and not errors:
            # No tools found at all
            empty = ErrorPage("No tools found",
                              f"No plugins found in:\n{TOOLS_DIR}\n\nCreate a .py file with get_tool_spec().")
            item = QListWidgetItem("No tools")
            self.nav.addItem(item)
            self.pages.addWidget(empty)
            self.nav.setCurrentRow(0)
            return

        for spec in tools:
            item = QListWidgetItem(spec.icon if isinstance(spec.icon, QIcon) else QIcon(), spec.name)
            self.nav.addItem(item)
            try:
                page = spec.factory()
            except Exception as e:
                page = ErrorPage(spec.name, f"Factory failed:\n{e}\n\n{traceback.format_exc()}")
            self.pages.addWidget(page)

        # Select first real tool (or errors page if present first)
        self.nav.setCurrentRow(0)

    def toggle_theme(self):
        self._dark = not self._dark
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        pal = self.palette()
        if self._dark:
            pal.setColor(QPalette.Window, QColor(45,45,48))
            pal.setColor(QPalette.WindowText, Qt.white)
            pal.setColor(QPalette.Base, QColor(37,37,38))
            pal.setColor(QPalette.AlternateBase, QColor(45,45,48))
            pal.setColor(QPalette.ToolTipBase, Qt.white)
            pal.setColor(QPalette.ToolTipText, Qt.white)
            pal.setColor(QPalette.Text, Qt.white)
            pal.setColor(QPalette.Button, QColor(45,45,48))
            pal.setColor(QPalette.ButtonText, Qt.white)
            pal.setColor(QPalette.BrightText, Qt.red)
            pal.setColor(QPalette.Highlight, QColor(14, 99, 156))
            pal.setColor(QPalette.HighlightedText, Qt.white)
        else:
            pal = QApplication.style().standardPalette()
        self.setPalette(pal)


def main():
    app = QApplication(sys.argv)
    QApplication.setStyle(QStyleFactory.create("Fusion"))
    w = AtlasWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
