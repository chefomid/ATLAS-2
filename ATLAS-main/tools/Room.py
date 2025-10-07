# tools/Acquisition_Data_Room.py
from pathlib import Path
import shutil, tempfile
from datetime import datetime

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QWidget, QGroupBox, QHBoxLayout, QVBoxLayout, QSplitter, QPushButton,
    QLabel, QListWidget, QTextEdit, QFileDialog, QAbstractItemView,
    QSizePolicy, QTabWidget, QMessageBox, QLayout, QFrame
)

from plugin_api import ToolSpec, BaseToolPage

# ---------- Helpers ----------
def _human_join(seq, sep=", ", last=" and "):
    seq = list(seq)
    if not seq: return ""
    if len(seq) == 1: return seq[0]
    return sep.join(seq[:-1]) + last + seq[-1]

def _safe_tempdir(prefix: str = "atlas_adr_") -> Path:
    p = Path(tempfile.mkdtemp(prefix=prefix))
    (p / "outputs").mkdir(exist_ok=True)
    return p

def _set_layout_zero(layout: QLayout):
    if not layout:
        return
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)

class FileDropListWidget(QListWidget):
    """Drag-and-drop aware list for local files."""
    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
        else: super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
        else: super().dragMoveEvent(e)

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            for url in e.mimeData().urls():
                lf = url.toLocalFile()
                if lf and Path(lf).exists():
                    self.addItem(lf)
            e.acceptProposedAction()
        else:
            super().dropEvent(e)

# ---------- Main Page ----------
class AcquisitionDataRoomPage(BaseToolPage):
    """
    Acquisition Data Room
    Panels: Inputs • Outputs • Narrative
    """
    def __init__(self, parent=None):
        super().__init__("Acquisition Data Room", "M&A consolidation and underwriting prep", parent)
        self.session_dir = _safe_tempdir()
        self.outputs_dir = self.session_dir / "outputs"

        # Compact styling for minimal chrome and tight groups
        self.setStyleSheet("""
        * { font-size: 12px; }
        QWidget { background: palette(base); }
        QSplitter::handle { margin: 0; }
        QTabWidget::pane { border: 0; margin: 0; padding: 0; }
        QTabBar::tab { padding: 3px 8px; margin: 0; }

        QGroupBox {
            margin-top: 4px; padding-top: 2px;
            border: 1px solid rgba(0,0,0,40); border-radius: 6px;
        }
        QGroupBox::title {
            subcontrol-origin: margin; subcontrol-position: top left;
            padding: 0px 6px; margin-top: 0px; font-weight: 600;
        }
        QListWidget, QTextEdit { padding: 2px; }
        """)

        # Build UI directly into BaseToolPage.content (no extra wrappers)
        main = self._build_main_ui()

        # Ensure our content area has zero margins/spacing
        try:
            self.content.setContentsMargins(0, 0, 0, 0)
            self.content.setSpacing(0)
        except Exception:
            pass
        self.setContentsMargins(0, 0, 0, 0)
        if self.layout():
            _set_layout_zero(self.layout())

        self.content.addWidget(main, 1)

        # Extra-hard clamp: remove unexpected gaps around title/header
        QTimer.singleShot(0, self._tighten_chrome)

        # After show, bias the vertical split so Narrative gets more height
        QTimer.singleShot(0, self._tune_heights)

        # Wire actions
        self.btn_add_files.clicked.connect(self._add_files)
        self.btn_remove_files.clicked.connect(self._remove_selected_uploads)
        self.btn_clear_inputs.clicked.connect(self._clear_inputs)
        self.btn_process.clicked.connect(self._process_inputs)
        self.btn_save_all.clicked.connect(self._save_all_outputs)

        self._refresh_input_stats()
        self._refresh_output_list()
        self.log_view.append(f"[{datetime.now().strftime('%H:%M:%S')}] Session started")

    # ---------- De-chrome helpers ----------
    def _tighten_chrome(self):
        """Aggressively remove top padding/whitespace above panels."""
        # Zero out root layout and any immediate child layouts
        for w in (self, getattr(self, "content", None)):
            if w:
                _set_layout_zero(w.layout())

        # Try to find any header-like container commonly used by BaseToolPage.
        self._squash_header_gaps()
        # Keep tuning after a resize completes
        self._tune_heights()

    def _squash_header_gaps(self):
        """
        Heuristically locate header/title areas and clamp margins/height.
        Works even if plugin_api's BaseToolPage changes its internal names.
        """
        # Clamp any top-level QFrames or widgets named/header-like
        candidates = []
        for w in self.findChildren(QWidget):
            name = (w.objectName() or "").lower()
            if any(k in name for k in ("header", "title", "subtitle")):
                candidates.append(w)
        # Fall back: first label(s) at the top of the page
        if not candidates:
            labels = [w for w in self.findChildren(QLabel)]
            # Prefer labels with larger font or bold (title-ish)
            candidates = labels[:2] if labels else []

        for w in candidates:
            try:
                w.setContentsMargins(0, 0, 0, 0)
            except Exception:
                pass
            lay = w.layout()
            if lay:
                _set_layout_zero(lay)
            # Avoid extra vertical stretch
            sp = w.sizePolicy()
            sp.setVerticalPolicy(QSizePolicy.Minimum)
            w.setSizePolicy(sp)

        # Also ensure the page root layout has no top padding
        if self.layout():
            _set_layout_zero(self.layout())

    # ---------- UI ----------
    def _build_main_ui(self) -> QWidget:
        # ----- Horizontal split: Inputs | Outputs -----
        mid = QSplitter(Qt.Horizontal)
        mid.setChildrenCollapsible(False)
        mid.setHandleWidth(6)

        # Inputs
        left = QWidget(); ll = QVBoxLayout(left)
        ll.setContentsMargins(4, 2, 4, 2); ll.setSpacing(4)

        grp_in = QGroupBox("Inputs")
        gi = QVBoxLayout(grp_in); gi.setContentsMargins(6, 4, 6, 6); gi.setSpacing(4)

        self.lbl_input_stats = QLabel("Queued: 0")

        self.list_inputs = FileDropListWidget()
        self.list_inputs.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_inputs.setAlternatingRowColors(True)
        self.list_inputs.setUniformItemSizes(True)

        row_controls = QHBoxLayout(); row_controls.setSpacing(6)
        self.btn_add_files = QPushButton("Add")
        self.btn_remove_files = QPushButton("Remove")
        self.btn_clear_inputs = QPushButton("Clear")
        self.btn_process = QPushButton("Process")
        row_controls.addWidget(self.btn_add_files)
        row_controls.addWidget(self.btn_remove_files)
        row_controls.addWidget(self.btn_clear_inputs)
        row_controls.addStretch(1)
        row_controls.addWidget(self.btn_process)

        gi.addWidget(self.lbl_input_stats)
        gi.addWidget(self.list_inputs, 1)
        gi.addLayout(row_controls)
        grp_in.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        ll.addWidget(grp_in, 1)

        # Outputs
        right = QWidget(); rl = QVBoxLayout(right)
        rl.setContentsMargins(4, 2, 4, 2); rl.setSpacing(4)

        grp_out = QGroupBox("Outputs")
        go = QVBoxLayout(grp_out); go.setContentsMargins(6, 4, 6, 6); go.setSpacing(4)

        self.lbl_output_stats = QLabel("Generated: 0")

        self.list_outputs = QListWidget()
        self.list_outputs.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_outputs.setAlternatingRowColors(True)
        self.list_outputs.setUniformItemSizes(True)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        self.btn_save_all = QPushButton("Save")
        r2.addStretch(1); r2.addWidget(self.btn_save_all)

        go.addWidget(self.lbl_output_stats)
        go.addWidget(self.list_outputs, 1)
        go.addLayout(r2)
        rl.addWidget(grp_out, 1)

        mid.addWidget(left)
        mid.addWidget(right)
        mid.setStretchFactor(0, 1)
        mid.setStretchFactor(1, 1)
        mid.setSizes([600, 600])

        # ----- Bottom: Narrative (tabs) -----
        bottom = QWidget()
        bl = QVBoxLayout(bottom)
        # zero margins so no gray ring around the Narrative group
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(0)

        grp_nav = QGroupBox("Narrative")
        gnl = QVBoxLayout(grp_nav)
        # tight inner margins so the content fills the group
        gnl.setContentsMargins(6, 4, 6, 6)
        gnl.setSpacing(4)

        self.tabs = QTabWidget()
        tbf = self.tabs.tabBar().font(); tbf.setBold(True); self.tabs.tabBar().setFont(tbf)

        self.narrative_view = QTextEdit()
        self.narrative_view.setPlaceholderText("Org summary, exposures, coverage recommendations, next steps…")

        self.log_view = QTextEdit(); self.log_view.setReadOnly(True)

        self.tabs.addTab(self.narrative_view, "Draft")
        self.tabs.addTab(self.log_view, "Log")

        gnl.addWidget(self.tabs, 1)
        bl.addWidget(grp_nav, 1)

        # ----- Vertical split: (mid) | (bottom) -----
        self._main_split = QSplitter(Qt.Vertical)  # keep handle to reshape heights
        self._main_split.setChildrenCollapsible(False)
        self._main_split.setHandleWidth(6)
        self._main_split.addWidget(mid)
        self._main_split.addWidget(bottom)

        # Bias bottom (Narrative) taller by default
        self._main_split.setStretchFactor(0, 2)
        self._main_split.setStretchFactor(1, 3)

        # Root container
        container = QWidget()
        root = QVBoxLayout(container)
        root.setContentsMargins(0, 0, 0, 0)  # no outer margins
        root.setSpacing(0)
        root.addWidget(self._main_split, 1)
        container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        return container

    # --- Keep the Narrative tall & remove dead space ---
    def _tune_heights(self):
        """Give ~58% height to the Narrative pane to minimize gray space."""
        if not hasattr(self, "_main_split"):
            return
        total = max(self._main_split.height(), 1)
        top_h = int(total * 0.42)
        bot_h = max(total - top_h, 1)
        self._main_split.setSizes([top_h, bot_h])

    def resizeEvent(self, e):
        super().resizeEvent(e)
        # Maintain the ratio on window resize so Narrative stays extended
        self._tune_heights()

    # ---------- Inputs ----------
    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Add Files")
        if not files: return
        for f in files: self.list_inputs.addItem(f)
        self._refresh_input_stats()
        self.log_view.append(f"Added {len(files)} file(s).")

    def _remove_selected_uploads(self):
        for it in self.list_inputs.selectedItems():
            self.list_inputs.takeItem(self.list_inputs.row(it))
        self._refresh_input_stats()

    def _clear_inputs(self):
        self.list_inputs.clear(); self._refresh_input_stats()

    def _refresh_input_stats(self):
        self.lbl_input_stats.setText(f"Queued: {self.list_inputs.count()}")

    # ---------- Outputs ----------
    def _refresh_output_list(self):
        self.list_outputs.clear()
        items = [p.name for p in sorted(self.outputs_dir.glob("*")) if p.is_file()]
        for n in items: self.list_outputs.addItem(n)
        self.lbl_output_stats.setText(f"Generated: {len(items)}")

    def _save_all_outputs(self):
        files = [p for p in sorted(self.outputs_dir.glob("*")) if p.is_file()]
        if not files:
            QMessageBox.information(self, "Save Outputs", "No generated files to save yet.")
            return
        target_dir = QFileDialog.getExistingDirectory(self, "Choose folder to save all outputs")
        if not target_dir: return
        copied = []
        for src in files:
            dst = Path(target_dir) / src.name
            try:
                shutil.copy2(str(src), str(dst))
                copied.append(src.name)
            except Exception as e:
                self.log_view.append(f"Save failed for {src.name}: {e}")
        if copied:
            self.log_view.append(f"Saved {len(copied)} file(s): {_human_join(copied)}.")

    # ---------- Templates & Processing ----------
    def _ensure_templates(self):
        created = []
        sov = self.outputs_dir / "Consolidated_SOV.csv"
        if not sov.exists():
            sov.write_text(
                "Loc #,Entity Name,Address,City,State,Zip,Building Value,BI Limit,Contents\n",
                encoding="utf-8"
            ); created.append(sov.name)
        bor = self.outputs_dir / "Broker_of_Record_Letter.txt"
        if not bor.exists():
            bor.write_text(
                "BROKER OF RECORD LETTER\n\nDate: ___\nTo: ___\n\n"
                "We hereby appoint USI Insurance Services as our exclusive Broker of Record effective ___.\n\n"
                "Sincerely,\n\n__________________________\nName, Title\n",
                encoding="utf-8"
            ); created.append(bor.name)
        uw = self.outputs_dir / "Underwriting_Packet.md"
        if not uw.exists():
            uw.write_text(
                "# Underwriting Packet (Placeholder)\n\n- Executive Summary\n- Consolidated SOV reference\n"
                "- Loss runs (attach when available)\n- Policies & endorsements index\n- Recommendations\n",
                encoding="utf-8"
            ); created.append(uw.name)
        if created:
            self.log_view.append(f"Created templates: {_human_join(created)}.")

    def _process_inputs(self):
        count = self.list_inputs.count()
        if count == 0:
            self.log_view.append("No input files to process."); return

        self.outputs_dir.mkdir(exist_ok=True)
        self._ensure_templates()

        mapping = self.outputs_dir / "Consolidation_Mapping_Report.csv"
        lines = ["Source,Action,Notes"]

        appended_rows = 0
        for i in range(count):
            src = Path(self.list_inputs.item(i).text())
            action, notes = "Indexed", ""
            try:
                ext = src.suffix.lower()
                if ext in {".xlsx", ".xls", ".csv"}:
                    try:
                        import pandas as pd
                        df = pd.read_excel(str(src)) if ext in {".xlsx",".xls"} else pd.read_csv(str(src))
                        cols = [str(c).strip().lower() for c in df.columns]
                        if any("loc" in c for c in cols) and any("address" in c for c in cols):
                            sov_path = self.outputs_dir / "Consolidated_SOV.csv"
                            rows = []
                            # find helper to locate first matching column by keywords
                            def _pick(col_keywords):
                                for kw in col_keywords:
                                    for c in df.columns:
                                        cl = str(c).lower()
                                        if kw in cl:
                                            return c
                                return None
                            for _, row in df.iterrows():
                                rec = [
                                    row.get(_pick(["loc"]), ""),
                                    row.get(_pick(["entity","name"]), ""),
                                    row.get(_pick(["address"]), ""),
                                    row.get(_pick(["city"]), ""),
                                    row.get(_pick(["state"]), ""),
                                    row.get(_pick(["zip","postal"]), ""),
                                    row.get(_pick(["building","tiv"]), ""),
                                    row.get(_pick(["bi"]), ""),
                                    row.get(_pick(["contents"]), ""),
                                ]
                                rows.append(rec)
                            if rows:
                                with open(sov_path, "a", encoding="utf-8") as f:
                                    for rec in rows:
                                        safe = ["" if v is None else str(v).replace(",", " ") for v in rec]
                                        f.write(",".join(safe) + "\n")
                                appended_rows += len(rows)
                                action, notes = "Normalized+Appended", f"{len(rows)} SOV rows"
                        else:
                            action, notes = "Indexed", "Spreadsheet (non-SOV)"
                    except Exception as e:
                        action, notes = "Indexed", f"read error: {e}"
                else:
                    action, notes = "Indexed", "Document/Image"
            except Exception as e:
                action, notes = "Skipped", str(e)

            lines.append(f"{src.name},{action},{notes}")

        mapping.write_text("\n".join(lines) + "\n", encoding="utf-8")
        self._refresh_output_list()
        self.log_view.append(
            f"Processed {count} file(s). Appended ~{appended_rows} SOV row(s). "
            f"See Consolidation_Mapping_Report.csv."
        )

        # Auto-generate a narrative draft (no buttons in UI)
        self._generate_narrative_draft()

    # ---------- Narrative (auto only) ----------
    def _generate_narrative_draft(self):
        files = [self.list_inputs.item(i).text() for i in range(self.list_inputs.count())]
        draft = (
            "**Integration Narrative (Draft)**\n\n"
            f"Reviewed files: {', '.join(Path(f).name for f in files) if files else 'none'}\n\n"
            "Overview:\nThis acquisition includes mixed property and liability exposures. "
            "Initial documents suggest standard ISO-based programs; verify endorsement compliance.\n\n"
            "Key Points:\n"
            "- Normalize SOV (COPE completeness, address standardization)\n"
            "- Confirm GL endorsements (Blanket AI, Waiver of Subrogation, Primary & Non-Contributory)\n"
            "- Review Umbrella/Excess adequacy vs. exposure profile\n"
            "- Obtain and analyze 5-year loss runs (frequency/severity)\n\n"
            "Next Steps:\n"
            "- Finalize Consolidated_SOV and reconcile with schedules\n"
            "- Prepare Broker of Record letter\n"
            "- Assemble Underwriting Packet for market submission\n"
        )
        self.narrative_view.setPlainText(draft)
        # Also write narrative to outputs so "Save" includes it
        (self.outputs_dir / "Narrative.md").write_text(draft, encoding="utf-8")
        self._refresh_output_list()
        self.log_view.append("Narrative draft updated.")

# ---------- ToolSpec ----------
def get_tool_spec() -> ToolSpec:
    def factory(): return AcquisitionDataRoomPage()
    try:
        return ToolSpec("acquisition_data_room", "Acquisition Data Room", factory)
    except TypeError:
        try:
            return ToolSpec("acquisition_data_room", "Acquisition Data Room", None, factory)
        except TypeError:
            return ToolSpec(id="acquisition_data_room", name="Acquisition Data Room", factory=factory)
