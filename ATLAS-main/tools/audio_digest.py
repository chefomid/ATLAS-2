# tools/audio_digest.py
from PyQt5.QtCore import Qt, QTimer, QTime
from PyQt5.QtWidgets import (
    QWidget, QGroupBox, QHBoxLayout, QVBoxLayout, QLabel, QPushButton,
    QLineEdit, QCheckBox, QComboBox, QSplitter, QTextEdit,
    QFileDialog, QProgressBar, QMessageBox, QTabWidget
)

from plugin_api import ToolSpec, BaseToolPage


class AudioDigestPage(BaseToolPage):
    """
    UI-only: Real-time audio capture -> transcript -> summarized notes.
    No audio backend wired yet; all actions are stubbed.
    """
    def __init__(self, parent=None):
        super().__init__("Scribe Assistant", "Real-time capture • Transcript • Summary", parent)

        c = self.content  # Use BaseToolPage's content layout

        # ---------- Session ----------
        box_ws = QGroupBox("Session")
        ws = QHBoxLayout(box_ws)

        self.txt_title = QLineEdit()
        self.txt_title.setPlaceholderText("Session title (e.g., Client Kickoff, Renewal Call)…")

        self.cmb_client = QComboBox()
        self.cmb_client.setEditable(True)
        self.cmb_client.setPlaceholderText("Select or type a client…")

        self.edt_folder = QLineEdit()
        self.edt_folder.setPlaceholderText("Save folder for audio/transcripts…")

        btn_browse = QPushButton("Browse…")
        btn_browse.clicked.connect(self._choose_folder)

        ws.addWidget(QLabel("Title:"))
        ws.addWidget(self.txt_title)
        ws.addSpacing(8)
        ws.addWidget(QLabel("Client:"))
        ws.addWidget(self.cmb_client, 1)
        ws.addSpacing(8)
        ws.addWidget(QLabel("Save to:"))
        ws.addWidget(self.edt_folder, 2)
        ws.addWidget(btn_browse)

        # ---------- Split: Controls (left) • Output (right) ----------
        split = QSplitter(Qt.Horizontal)

        # Left controls
        left = QWidget()
        ll = QVBoxLayout(left)

        grp_rec = QGroupBox("Recorder")
        rl = QVBoxLayout(grp_rec)

        line1 = QHBoxLayout()
        self.cmb_device = QComboBox(); self.cmb_device.setEditable(True)
        self.cmb_device.setPlaceholderText("Input device (placeholder)…")
        self.cmb_samplerate = QComboBox(); self.cmb_samplerate.addItems(["16 kHz", "22.05 kHz", "44.1 kHz", "48 kHz"])
        self.cmb_samplerate.setCurrentText("16 kHz")
        self.cmb_format = QComboBox(); self.cmb_format.addItems(["WAV (PCM)", "MP3", "FLAC"])
        line1.addWidget(QLabel("Device:")); line1.addWidget(self.cmb_device)
        line1.addSpacing(6); line1.addWidget(QLabel("Rate:")); line1.addWidget(self.cmb_samplerate)
        line1.addSpacing(6); line1.addWidget(QLabel("Format:")); line1.addWidget(self.cmb_format)

        line2 = QHBoxLayout()
        self.btn_record = QPushButton("● Record")
        self.btn_pause  = QPushButton("⏸ Pause"); self.btn_pause.setEnabled(False)
        self.btn_stop   = QPushButton("■ Stop");  self.btn_stop.setEnabled(False)
        self.btn_record.clicked.connect(self._start_recording_stub)
        self.btn_pause.clicked.connect(self._pause_recording_stub)
        self.btn_stop.clicked.connect(self._stop_recording_stub)
        line2.addWidget(self.btn_record); line2.addWidget(self.btn_pause); line2.addWidget(self.btn_stop)

        line3 = QHBoxLayout()
        self.lbl_timer = QLabel("00:00"); self.lbl_timer.setStyleSheet("font-weight:600;")
        self.meter = QProgressBar(); self.meter.setRange(0, 100); self.meter.setValue(0)
        self.ch_auto_save = QCheckBox("Auto-save on stop"); self.ch_auto_save.setChecked(True)
        line3.addWidget(QLabel("Time:")); line3.addWidget(self.lbl_timer)
        line3.addSpacing(12); line3.addWidget(QLabel("Level:")); line3.addWidget(self.meter, 1)
        line3.addSpacing(12); line3.addWidget(self.ch_auto_save)

        rl.addLayout(line1); rl.addLayout(line2); rl.addLayout(line3)

        grp_actions = QGroupBox("Actions")
        al = QVBoxLayout(grp_actions)
        rowA = QHBoxLayout()
        self.btn_transcribe = QPushButton("Generate Transcript"); self.btn_transcribe.setEnabled(False)
        self.btn_transcribe.clicked.connect(self._transcribe_stub)
        self.btn_summarize = QPushButton("Summarize to Notes"); self.btn_summarize.setEnabled(False)
        self.btn_summarize.clicked.connect(self._summarize_stub)
        rowA.addWidget(self.btn_transcribe); rowA.addWidget(self.btn_summarize)

        rowB = QHBoxLayout()
        self.btn_export_docx = QPushButton("Export Summary (.docx)"); self.btn_export_docx.setEnabled(False)
        self.btn_export_txt  = QPushButton("Export Transcript (.txt)"); self.btn_export_txt.setEnabled(False)
        self.btn_export_docx.clicked.connect(self._export_docx_stub)
        self.btn_export_txt.clicked.connect(self._export_txt_stub)
        rowB.addWidget(self.btn_export_docx); rowB.addWidget(self.btn_export_txt)

        al.addLayout(rowA); al.addLayout(rowB)
        ll.addWidget(grp_rec); ll.addWidget(grp_actions); ll.addStretch()

        # Right output tabs
        right = QWidget()
        rl_out = QVBoxLayout(right)
        self.tabs = QTabWidget()
        self.txt_transcript = QTextEdit(); self.txt_transcript.setPlaceholderText("Transcript will appear here…")
        self.txt_summary = QTextEdit();    self.txt_summary.setPlaceholderText("Summarized notes will appear here…")
        self.tabs.addTab(self.txt_transcript, "Transcript")
        self.tabs.addTab(self.txt_summary, "Summary")
        rl_out.addWidget(self.tabs)

        split.addWidget(left); split.addWidget(right)
        split.setStretchFactor(0, 0); split.setStretchFactor(1, 1)

        # Footer
        foot = QHBoxLayout()
        self.lbl_status = QLabel("Ready."); foot.addWidget(self.lbl_status, 1)

        # Assemble into BaseToolPage.content
        c.addWidget(box_ws)
        c.addWidget(split)
        c.addLayout(foot)

        # Timer stubs
        self._rec_timer = QTimer(self); self._rec_timer.timeout.connect(self._tick_clock)
        self._elapsed = QTime(0, 0, 0); self._is_recording = False; self._is_paused = False

    # ---------- Stubs ----------
    def _choose_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Choose save folder")
        if path: self.edt_folder.setText(path)

    def _start_recording_stub(self):
        if self._is_recording and not self._is_paused: return
        self._is_recording = True; self._is_paused = False
        self._rec_timer.start(1000)
        self.btn_record.setEnabled(False); self.btn_pause.setEnabled(True); self.btn_stop.setEnabled(True)
        self.lbl_status.setText("Recording (stub)…")
        self.btn_transcribe.setEnabled(True)

    def _pause_recording_stub(self):
        if not self._is_recording: return
        self._is_paused = not self._is_paused
        if self._is_paused:
            self._rec_timer.stop(); self.lbl_status.setText("Paused."); self.btn_pause.setText("▶ Resume")
        else:
            self._rec_timer.start(1000); self.lbl_status.setText("Recording (stub)…"); self.btn_pause.setText("⏸ Pause")

    def _stop_recording_stub(self):
        if not self._is_recording: return
        self._is_recording = False; self._is_paused = False
        self._rec_timer.stop()
        self.btn_record.setEnabled(True); self.btn_pause.setEnabled(False); self.btn_pause.setText("⏸ Pause")
        self.btn_stop.setEnabled(False)
        self.lbl_status.setText("Stopped. (Audio file would be saved here.)")

    def _tick_clock(self):
        self._elapsed = self._elapsed.addSecs(1)
        self.lbl_timer.setText(self._elapsed.toString("mm:ss"))
        self.meter.setValue((self._elapsed.second() * 13) % 100)

    def _transcribe_stub(self):
        self.txt_transcript.setPlainText(
            "[Demo transcript]\n"
            "Speaker 1: Thanks for joining. Today we’ll confirm renewal tasks, COI handling, and the new flood endorsement.\n"
            "Speaker 2: Let’s also capture action items and responsible parties.\n"
        )
        self.lbl_status.setText("Transcript generated (stub).")
        self.btn_summarize.setEnabled(True); self.btn_export_txt.setEnabled(True)

    def _summarize_stub(self):
        self.txt_summary.setPlainText(
            "• Purpose: Renewal prep call; confirm COI workflow and new flood endorsement.\n"
            "• Key decisions: Maintain existing AI limits; add Flood endorsement review.\n"
            "• Action items:\n"
            "  - AM to draft COI template variants (due Fri)\n"
            "  - Client to send updated locations list (due Mon)\n"
            "• Next meeting: Tues @ 10am PT\n"
        )
        self.lbl_status.setText("Summary created (stub).")
        self.btn_export_docx.setEnabled(True)

    def _export_docx_stub(self):
        QMessageBox.information(self, "Export", "Would export Summary to .docx (stub).")

    def _export_txt_stub(self):
        QMessageBox.information(self, "Export", "Would export Transcript to .txt (stub).")


def get_tool_spec() -> ToolSpec:
    """
    Register with ToolSpec (id, name, factory). Factory is zero-arg per plugin_api.
    """
    def factory() -> QWidget:
        return AudioDigestPage()

    return ToolSpec(
        id="audio_digest",
        name="Scribe Assistant",
        factory=factory,
        order=60,
    )
