# tools/premium_allocator.py
from datetime import datetime
import os
import shutil
import importlib
import inspect
import json
import traceback
from pathlib import Path

from PyQt5.QtWidgets import (
    QGroupBox, QHBoxLayout, QVBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QTextEdit, QProgressBar,
    QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, QTimer
from PyQt5.Qt import QDesktopServices, QUrl

from plugin_api import ToolSpec, BaseToolPage, runtime_path


# ---------------- Worker (runs in background thread) ----------------
class CalcWorker(QObject):
    log = pyqtSignal(str)
    done = pyqtSignal(str)          # emits output file path
    failed = pyqtSignal(str)        # emits error message

    def __init__(self, input_path, mode, output_root):
        super().__init__()
        self.input_path = Path(input_path)        # working copy the user edited
        self.mode = str(mode)                     # "RAS" or "TIV"
        self.output_root = Path(output_root)

    # ----- config -----
    def _load_backend_config(self) -> dict:
        # backend_config.json should live at project root (next to atlas_qt.py / plugin_api.py)
        cfg_path = runtime_path("backend_config.json")
        if cfg_path.exists():
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
            except Exception as e:
                self.log.emit(f"Warning: failed to read backend_config.json: {e}")
        return {}

    # ----- entrypoint selection (uses config first) -----
    def _pick_entrypoint(self, mod, module_name: str, preferred_names: tuple[str, ...]):
        """
        Return (callable, cfg_for_module, picked_name, picked_via).
        Priority:
          1) backend_config.json -> entrypoint
          2) ATLAS_ENTRYPOINT / ATLAS_RUN
          3) One of preferred_names
          4) Runner classes with run()/calculate()/process()
          5) Heuristic scan (names containing ras/tiv/allocate/weighted)
        """
        cfg_all = self._load_backend_config()
        cfg = cfg_all.get(module_name, {}) if isinstance(cfg_all, dict) else {}
        # 1) config
        ep_from_cfg = cfg.get("entrypoint")
        if ep_from_cfg:
            cand = getattr(mod, ep_from_cfg, None)
            if callable(cand):
                return cand, cfg, ep_from_cfg, "config"

        # 2) module-level variables
        for var in ("ATLAS_ENTRYPOINT", "ATLAS_RUN"):
            ep = getattr(mod, var, None)
            if isinstance(ep, str):
                cand = getattr(mod, ep, None)
                if callable(cand):
                    return cand, cfg, ep, f"module:{var}"
            elif callable(ep):
                name = getattr(ep, "__name__", var)
                return ep, cfg, name, f"module:{var}"

        # 3) preferred names
        for name in preferred_names:
            f = getattr(mod, name, None)
            if callable(f):
                return f, cfg, name, "preferred"

        # 4) class runners (instantiate no-arg and use run/calculate/process)
        runner_names = ["AtlasBackend", "Runner", "Engine", "Allocator"]
        if self.mode.upper() == "RAS":
            runner_names += ["RASRunner", "RASEngine", "RASAllocator"]
        else:
            runner_names += ["TIVRunner", "TIVEngine", "TIVAllocator"]

        for cname in runner_names:
            C = getattr(mod, cname, None)
            if isinstance(C, type):
                try:
                    inst = C()
                    for meth in ("run", "calculate", "process"):
                        if hasattr(inst, meth) and callable(getattr(inst, meth)):
                            fn = getattr(inst, meth)
                            return fn, cfg, f"{cname}.{meth}", "runner"
                except Exception:
                    pass

        # 5) heuristic scan
        for name, obj in inspect.getmembers(mod, inspect.isfunction):
            lname = name.lower()
            if self.mode.upper() == "RAS":
                if any(k in lname for k in ("build_ras", "ras", "allocate", "distribution", "premium")):
                    return obj, cfg, name, "heuristic"
            else:
                if any(k in lname for k in ("build_tiv", "tiv", "weighted", "allocate", "distribution")):
                    return obj, cfg, name, "heuristic"

        return None, cfg, None, "none"

    def _call_entrypoint(self, fn, cfg: dict, input_path: Path, output_path: Path) -> Path:
        """
        Call `fn` using config-driven calling or smart fallbacks.
        Returns a Path to the produced workbook (must exist).
        """
        sig = inspect.signature(fn)
        call_mode = (cfg.get("call") or "auto").lower()

        # Build default kwargs from signature (auto mapping)
        kwargs = {}
        in_aliases = {"input_path", "input_file", "in_path", "infile", "source", "workbook", "xlsx_in", "path_in", "path_str", "path"}
        out_aliases = {"output_path", "output_file", "out_path", "outfile", "dest", "xlsx_out", "path_out"}
        log_aliases = {"log", "logger", "emit", "progress_cb", "cb"}
        for p in sig.parameters.values():
            nm = p.name.lower()
            if nm in in_aliases:
                kwargs[p.name] = str(input_path)
            elif nm in out_aliases:
                kwargs[p.name] = str(output_path)
            elif nm in log_aliases:
                kwargs[p.name] = self.log.emit
            elif nm == "mode":
                kwargs[p.name] = self.mode

        # Apply explicit param mapping from config (overrides autodetect)
        param_map = cfg.get("params") or {}
        for generic, actual in param_map.items():
            if generic == "input_path":
                kwargs[actual] = str(input_path)
            elif generic == "output_path":
                kwargs[actual] = str(output_path)
            elif generic == "log":
                kwargs[actual] = self.log.emit
            elif generic == "mode":
                kwargs[actual] = self.mode

        # Build positional args if requested
        args = []
        if call_mode == "positional":
            order = cfg.get("args") or []
            for key in order:
                if key == "input_path":
                    args.append(str(input_path))
                elif key == "output_path":
                    args.append(str(output_path))
                elif key == "log":
                    args.append(self.log.emit)
                elif key == "mode":
                    args.append(self.mode)

        # Try configured style first
        try_styles = []
        if call_mode == "kwargs":
            try_styles = [("kwargs-only", lambda: fn(**kwargs))]
        elif call_mode == "positional":
            try_styles = [("positional", lambda: fn(*args))]
        else:  # auto
            try_styles = [
                ("kwargs-only", lambda: fn(**kwargs)),
                ("positional-config", lambda: fn(*args)) if cfg.get("args") else None,
                ("(in,out,log)", lambda: fn(str(input_path), str(output_path), self.log.emit)),
                ("(in,out)",     lambda: fn(str(input_path), str(output_path))),
                ("(in)",         lambda: fn(str(input_path))),
            ]
            try_styles = [t for t in try_styles if t is not None]

        last_err = None
        for label, caller in try_styles:
            try:
                self.log.emit(f"Invoking {getattr(fn, '__module__', '?')}.{getattr(fn, '__name__', 'callable')} {label}")
                ret = caller()
                out = Path(ret) if isinstance(ret, (str, Path)) and ret else output_path
                if out.exists():
                    return out
                self.log.emit("Backend did not create the expected file; trying another call style…")
            except TypeError as e:
                last_err = e
                continue
            except Exception as e:
                tb = traceback.format_exc(limit=2)
                raise RuntimeError(f"Backend raised an error: {e}\n{tb}")

        raise RuntimeError(
            "Backend function did not produce an output file at the expected location "
            f"({output_path}). Last signature error: {last_err}"
        )

    def _invoke_backend(self, module_name: str, preferred_names, input_path: Path, output_path: Path) -> Path:
        self.log.emit(f"Loading backend: {module_name}")
        try:
            mod = importlib.import_module(module_name)
        except Exception as e:
            raise RuntimeError(f"Could not import {module_name}: {e}")

        fn, cfg, picked_name, picked_via = self._pick_entrypoint(mod, module_name, preferred_names)
        if fn is None:
            raise RuntimeError(
                f"{module_name} is missing an entrypoint. "
                f"Tried names: {', '.join(preferred_names)} or set backend_config.json for this module."
            )
        self.log.emit(f"Using entrypoint: {module_name}.{picked_name} (via {picked_via})")
        return self._call_entrypoint(fn, cfg, input_path, output_path)

    def run(self):
        try:
            if not self.input_path.exists():
                raise FileNotFoundError(f"Input workbook not found: {self.input_path}")

            self.output_root.mkdir(parents=True, exist_ok=True)

            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            base = self.input_path.stem

            if self.mode.upper() == "RAS":
                out_name = f"RAS-{base}-{ts}.xlsx"
                module_name = "ras_module"
                names = ("ATLAS_RUN", "run", "run_ras", "calculate", "calculate_ras", "main", "process", "build_ras")
            else:
                out_name = f"TIV-{base}-{ts}.xlsx"
                module_name = "tiv_module"
                names = ("ATLAS_RUN", "run", "run_tiv", "calculate", "calculate_tiv", "main", "process", "build_tiv")

            out_file = self.output_root / out_name

            self.log.emit(f"Starting calculation in {self.mode} mode…")
            self.log.emit(f"Input:  {self.input_path}")
            self.log.emit(f"Target: {out_file}")

            produced = self._invoke_backend(module_name, names, self.input_path, out_file)

            self.log.emit("Calculation finished.")
            self.done.emit(str(produced))

        except Exception as e:
            self.failed.emit(str(e))


# ---------------- UI Page ----------------

class PremiumAllocatorPage(BaseToolPage):
    def __init__(self, parent=None):
        super().__init__("Premium Allocator", "RAS Algorithm / TIV distribution", parent)

        # --- UI blocks ---
        wb = QGroupBox("Workbook")
        wl = QHBoxLayout(wb)
        self.lbl_selected = QLabel("Selected: —")
        self.btn_open_template = QPushButton("Open Template")
        self.btn_upload = QPushButton("Upload")
        wl.addWidget(self.lbl_selected, 1)
        wl.addWidget(self.btn_open_template)
        wl.addWidget(self.btn_upload)

        mode = QGroupBox("Mode")
        ml = QHBoxLayout(mode)
        self.rb_ras = QRadioButton("RAS Algorithm Distribution")
        self.rb_tiv = QRadioButton("TIV Weighted Distribution")
        self.rb_ras.setChecked(True)
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.rb_ras)
        self.mode_group.addButton(self.rb_tiv)
        ml.addWidget(self.rb_ras)
        ml.addWidget(self.rb_tiv)
        ml.addStretch(1)

        actions = QHBoxLayout()
        self.btn_calculate = QPushButton("Calculate")
        self.btn_open_out = QPushButton("Open Output Folder")
        self.btn_open_out.setEnabled(False)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # indeterminate when visible
        self.progress.setVisible(False)
        actions.addWidget(self.btn_calculate)
        actions.addWidget(self.btn_open_out)
        actions.addStretch(1)
        actions.addWidget(self.progress)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(160)
        self.log.setPlaceholderText("Status will appear here…")

        self.content.addWidget(wb)
        self.content.addWidget(mode)
        self.content.addLayout(actions)
        self.content.addWidget(self.log)

        # --- runtime state & folders ---
        self.selected_path = None  # Path or None
        self.base_dir = Path.home() / "Documents" / "ATLAS"
        self.work_dir = self.base_dir / "PA_Working"
        self.output_root = self.base_dir / "PA_Output"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.output_root.mkdir(parents=True, exist_ok=True)

        # modeless instructions box (kept as attribute so it stays open)
        self._instr_box = None  # QMessageBox or None

        # thread/worker holders
        self._thread = None  # QThread or None
        self._worker = None  # CalcWorker or None

        # --- wire signals ---
        self.btn_open_template.clicked.connect(self.on_open_template)
        self.btn_upload.clicked.connect(self.on_upload)
        self.btn_calculate.clicked.connect(self.on_calculate)
        self.btn_open_out.clicked.connect(self.on_open_output_folder)

        self._log_boot()

    # ---------- helpers ----------
    def _log(self, msg):
        self.log.append(msg)

    def _log_boot(self):
        tpl = runtime_path("Template.xlsx")
        where = "found" if tpl.exists() else "NOT found"
        self._log(f"Template.xlsx {where} at: {tpl}")
        self._log(f"Working folder: {self.work_dir}")
        self._log(f"Output folder:  {self.output_root}")

    def _show_instructions_modeless(self):
        html = (
            "<b>How to use Premium Allocator</b><br><br>"
            "1) <b>Input data</b> into the opened working copy.<br>"
            "2) Press <b>CTRL + S</b> to save the workbook.<br>"
            "3) Click <b>Upload</b> (the same file is already selected).<br>"
            "4) <b>Choose Mode</b> (RAS or TIV).<br>"
            "5) Click <b>Calculate</b> to produce the output."
        )
        if self._instr_box is None:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Information)
            box.setWindowTitle("Instructions")
            box.setTextFormat(Qt.RichText)
            box.setText(html)
            box.setStandardButtons(QMessageBox.Close)
            box.setWindowModality(Qt.NonModal)   # modeless
            box.setAttribute(Qt.WA_DeleteOnClose, False)
            self._instr_box = box
        else:
            self._instr_box.setText(html)

        self._instr_box.show()
        self._instr_box.raise_()
        self._instr_box.activateWindow()

    # ---------- actions ----------
    def on_open_template(self):
        """
        Show a modeless instruction popup AND open a fresh working COPY of the template.
        The master Template.xlsx remains pristine; each click creates a new timestamped copy.
        Also auto-selects that copy for Upload/Calculate.
        """
        self._show_instructions_modeless()

        master = runtime_path("Template.xlsx")
        if not master.exists():
            QMessageBox.warning(
                self,
                "Template not found",
                "Template.xlsx was not found next to the app.\n"
                "If you are running a frozen build, ensure it was bundled."
            )
            self._log("Template.xlsx missing.")
            return

        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        working_file = self.work_dir / f"Template-Working-{ts}.xlsx"

        try:
            shutil.copy2(str(master), str(working_file))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not prepare working copy:\n{e}")
            self._log(f"ERROR copying master template: {e}")
            return

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(working_file)))
        self.selected_path = working_file
        self.lbl_selected.setText(f"Selected: {working_file.name}")
        self._log(f"Opened fresh working copy: {working_file}")

    def on_upload(self):
        """
        Confirm that the selected workbook exists.
        If a working copy was just opened, it's already selected; otherwise allow browsing.
        """
        if self.selected_path and Path(self.selected_path).exists():
            self._log(f"Re-confirmed workbook: {self.selected_path}")
            QMessageBox.information(self, "Upload", f"Workbook selected:\n{self.selected_path}")
            return

        fn, _ = QFileDialog.getOpenFileName(
            self, "Select workbook", str(self.work_dir if self.work_dir.exists() else Path.home()),
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)"
        )
        if not fn:
            return
        self.selected_path = Path(fn)
        self.lbl_selected.setText(f"Selected: {self.selected_path.name}")
        self._log(f"Selected workbook: {self.selected_path}")

    def on_calculate(self):
        """
        Background calculation that uses the UPLOADED WORKING COPY as input.
        Streams logs; writes output to Documents/ATLAS/PA_Output; auto-opens result.
        """
        if not self.selected_path or not Path(self.selected_path).exists():
            QMessageBox.warning(self, "No workbook selected",
                                "Please click 'Open Template' (and save) or 'Upload' to choose a workbook first.")
            return

        mode = "RAS" if self.rb_ras.isChecked() else "TIV"

        # UI: busy
        self.progress.setVisible(True)
        self.btn_calculate.setEnabled(False)
        self.btn_upload.setEnabled(False)
        self.btn_open_template.setEnabled(False)

        # Thread + worker
        self._thread = QThread(self)
        self._worker = CalcWorker(self.selected_path, mode, self.output_root)
        self._worker.moveToThread(self._thread)

        # Wire signals
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._log)
        self._worker.done.connect(self._on_calc_done)
        self._worker.failed.connect(self._on_calc_failed)

        # Cleanup on finish
        self._worker.done.connect(lambda _: self._teardown_worker())
        self._worker.failed.connect(lambda _: self._teardown_worker())

        self._thread.start()

    # ----- worker callbacks -----
    def _on_calc_done(self, out_path: str):
        self._log(f"Output written: {out_path}")
        self.btn_open_out.setEnabled(True)

        def _open_output():
            opened = False
            try:
                os.startfile(out_path)  # Windows-native; brings Excel to front
                opened = True
            except Exception as e:
                self._log(f"startfile failed: {e}")
            if not opened:
                try:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(out_path)))
                    opened = True
                except Exception as e:
                    self._log(f"QDesktopServices failed: {e}")
            if opened:
                self._log(f"Opened output file: {out_path}")
            else:
                self._log("Could not auto-open output; use 'Open Output Folder' instead.")

        QTimer.singleShot(150, _open_output)
        QMessageBox.information(self, "Calculation complete", f"Output saved to:\n{out_path}")

    def _on_calc_failed(self, message: str):
        self._log(f"ERROR: {message}")
        QMessageBox.critical(self, "Calculation failed", message)

    def _teardown_worker(self):
        # Reset UI
        self.progress.setVisible(False)
        self.btn_calculate.setEnabled(True)
        self.btn_upload.setEnabled(True)
        self.btn_open_template.setEnabled(True)

        # Properly stop thread/worker
        try:
            if self._thread:
                self._thread.quit()
                self._thread.wait(2000)
        finally:
            self._thread = None
            self._worker = None

    def on_open_output_folder(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_root)))


def get_tool_spec():
    return ToolSpec(
        id="premium_allocator",
        name="Premium Allocator",
        factory=lambda: PremiumAllocatorPage(),
        icon=None,
        order=10
    )
