# tools/Insured_Intelligence_Desk.py
from pathlib import Path
from PyQt5.QtCore import Qt, QSettings, QTimer
from PyQt5.QtWidgets import (
    QWidget, QGroupBox, QHBoxLayout, QVBoxLayout, QSplitter, QComboBox, QPushButton,
    QLabel, QListWidget, QTabWidget, QTextEdit, QLineEdit, QFileDialog, QAbstractItemView,
    QSizePolicy
)
from plugin_api import ToolSpec, BaseToolPage


class ClientBriefingRoomPage(BaseToolPage):
    """
    Insured Intelligence Desk
    Left: Upload Files • Center: Chat & Log • Right: Extracted Assets
    """
    def __init__(self, parent=None):
        super().__init__("Insured Intelligence Desk", "Corpus-aware account assistant", parent)

        self.settings = QSettings("ATLAS", "InsuredIntelligenceDesk")
        self._last_client_root = self.settings.value("last_client_root", "", type=str)

        # Build UI into a single container
        self._container = self._build_container()

        # Try to attach the container using common BaseToolPage patterns.
        attached = False
        for method_name in (
            "set_body", "set_body_widget", "setContent", "set_content",
            "setCentralWidget", "setWidget", "set_page_widget"
        ):
            m = getattr(self, method_name, None)
            if callable(m):
                try:
                    m(self._container)
                    attached = True
                    break
                except Exception:
                    pass

        if not attached:
            body_layout = getattr(self, "body_layout", None)
            if body_layout and hasattr(body_layout, "addWidget"):
                body_layout.addWidget(self._container)
                attached = True

        if not attached:
            lay = QVBoxLayout(self)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.addWidget(self._container)

        if not attached:
            QTimer.singleShot(0, self._late_attach)

        # Wiring
        self.btn_open_folder.clicked.connect(self._choose_client_folder)
        self.btn_reindex.clicked.connect(self._simulate_reindex)
        self.btn_send.clicked.connect(self._simulate_answer)
        self.btn_open_asset.clicked.connect(self._simulate_open_asset)

        self.btn_add_files.clicked.connect(self._add_files)
        self.btn_remove_files.clicked.connect(self._remove_selected_uploads)
        self.btn_import_files.clicked.connect(self._simulate_import_to_corpus)

        # Restore last session
        if self._last_client_root:
            self._load_client_summary(Path(self._last_client_root))

    # ---------- Compatibility hooks ----------
    def build(self) -> QWidget:
        return self._container
    def build_body(self) -> QWidget:
        return self._container
    def content_widget(self) -> QWidget:
        return self._container
    def body(self) -> QWidget:
        return self._container

    # ---------- Internal: create the UI container ----------
    def _build_container(self) -> QWidget:
        # ---- Top bar ----
        top = QGroupBox("Workspace")
        tl = QHBoxLayout(top)
        tl.setContentsMargins(8, 6, 8, 6)  # compact header
        tl.setSpacing(8)
        # keep the top bar from stealing height
        top.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        self.cmb_client = QComboBox()
        self.cmb_client.setEditable(True)
        self.cmb_client.setPlaceholderText("Select or type a client name…")
        self.btn_open_folder = QPushButton("Open Client Folder")
        self.btn_reindex = QPushButton("Analyze / Re-index")
        self.cmb_model = QComboBox()
        self.cmb_model.addItems(["OpenAI — Evidence Mode", "Local — Fast (no web)", "Hybrid"])
        self.scope_line = QLineEdit()
        self.scope_line.setPlaceholderText("Context scope: e.g. Policies;2024;GL;Endorsements")

        tl.addWidget(QLabel("Client:"))
        tl.addWidget(self.cmb_client, 2)
        tl.addWidget(self.btn_open_folder)
        tl.addWidget(self.btn_reindex)
        tl.addWidget(QLabel("Model:"))
        tl.addWidget(self.cmb_model)
        tl.addWidget(self.scope_line, 2)

        # ---- Three panes ----
        splitter = QSplitter(Qt.Horizontal)
        splitter.setObjectName("MainSplitter")

        # Left: Upload Files
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)

        grp_upload = QGroupBox("Upload Files")
        gu = QVBoxLayout(grp_upload)
        self.lbl_upload_stats = QLabel("Queued: 0")
        self.list_uploads = QListWidget()
        self.list_uploads.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_uploads.setAlternatingRowColors(True)

        self.btn_add_files = QPushButton("Add Files")
        self.btn_remove_files = QPushButton("Remove Selected")
        self.btn_import_files = QPushButton("Import to Corpus")

        gu.addWidget(self.lbl_upload_stats)
        gu.addWidget(self.list_uploads, 1)
        row_up = QHBoxLayout()
        row_up.addWidget(self.btn_add_files)
        row_up.addWidget(self.btn_remove_files)
        row_up.addWidget(self.btn_import_files)
        gu.addLayout(row_up)

        # make the group expand vertically
        grp_upload.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.list_uploads.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        ll.addWidget(grp_upload, 1)

        # Center: Chat & Log
        center = QWidget()
        cl = QVBoxLayout(center)
        cl.setContentsMargins(0, 0, 0, 0)

        grp_chat = QGroupBox("Account Chat")
        gc = QVBoxLayout(grp_chat)
        self.chat_view = QTextEdit()
        self.chat_view.setReadOnly(True)
        self.prompt_box = QLineEdit()
        self.prompt_box.setPlaceholderText("Ask: e.g., How many locations this term? Do they have blanket AI?")
        self.btn_send = QPushButton("Ask")
        row = QHBoxLayout()
        row.addWidget(self.prompt_box, 1)
        row.addWidget(self.btn_send)
        gc.addWidget(self.chat_view, 1)
        gc.addLayout(row)

        self.tabs_center = QTabWidget()
        self.tabs_center.addTab(grp_chat, "Chat")
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.tabs_center.addTab(self.log_view, "Output Log")

        # expand center stack vertically
        self.tabs_center.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        grp_chat.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.chat_view.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        cl.addWidget(self.tabs_center, 1)

        # Right: Extracted Assets
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)

        grp_assets = QGroupBox("Extracted Assets")
        ga = QVBoxLayout(grp_assets)
        self.list_assets = QListWidget()
        self.list_assets.setSelectionMode(self.list_assets.SingleSelection)
        self.btn_open_asset = QPushButton("Open / Reveal")
        ga.addWidget(self.list_assets, 1)
        ga.addWidget(self.btn_open_asset)

        # expand vertically
        grp_assets.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.list_assets.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        rl.addWidget(grp_assets, 1)

        splitter.addWidget(left)
        splitter.addWidget(center)
        splitter.addWidget(right)

        # distribute width; ensure the splitter claims vertical space
        splitter.setStretchFactor(0, 1)  # left
        splitter.setStretchFactor(1, 2)  # center
        splitter.setStretchFactor(2, 1)  # right

        # sensible minimum heights so they grow to consume gray area
        min_h = 550
        grp_upload.setMinimumHeight(min_h)
        grp_assets.setMinimumHeight(min_h)
        self.tabs_center.setMinimumHeight(min_h)

        # Wrap into a body container
        container = QWidget()
        root = QVBoxLayout(container)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)
        root.addWidget(top)
        root.addWidget(splitter, 1)  # <- splitter gets the vertical stretch
        # also set the container to prefer expanding vertically
        container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        return container

    def _late_attach(self):
        for method_name in ("set_body", "set_body_widget", "setContent", "set_content",
                            "setCentralWidget", "setWidget", "set_page_widget"):
            m = getattr(self, method_name, None)
            if callable(m):
                try:
                    m(self._container)
                    return
                except Exception:
                    pass
        body_layout = getattr(self, "body_layout", None)
        if body_layout and hasattr(body_layout, "addWidget"):
            body_layout.addWidget(self._container)
            return
        lay = self.layout()
        if lay is None:
            lay = QVBoxLayout(self)
        lay.addWidget(self._container)

    # --------- Minimal behaviors (stubs to hook up later) ----------
    def _choose_client_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select Client Folder")
        if not path:
            return
        self._last_client_root = path
        self.settings.setValue("last_client_root", path)
        self._load_client_summary(Path(path))

    def _load_client_summary(self, root: Path):
        self.cmb_client.setEditText(root.name)
        self.list_assets.clear()
        for a in ["vehicle_schedule.xlsx", "location_schedule.xlsx", "forms_detected.json", "policy_register.csv"]:
            self.list_assets.addItem(a)
        self.chat_view.append("<i>Loaded client:</i> " + root.name)
        self._refresh_upload_stats()

    def _simulate_reindex(self):
        self.chat_view.append("<i>Re-indexing requested… (stub)</i>")

    def _simulate_answer(self):
        q = self.prompt_box.text().strip()
        if not q:
            return
        self.prompt_box.clear()
        self.chat_view.append(f"<b>You:</b> {q}")
        self.chat_view.append("<b>AI:</b> Based on current index, there are <b>9 locations</b>. "
                              "See: policy_register.csv (rows 12–20).")
        self.log_view.append(f"[{self.cmb_client.currentText()}] Q: {q} | A: stubbed\n")

    def _simulate_open_asset(self):
        item = self.list_assets.currentItem()
        if not item:
            return
        self.chat_view.append(f"<i>Revealing asset:</i> {item.text()}")

    # --------- Upload panel handlers ---------
    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add Files",
            str(Path(self._last_client_root)) if self._last_client_root else "",
            "All Files (*);;Documents (*.pdf *.docx *.xlsx *.csv *.txt *.eml *.msg);;Images (*.png *.jpg *.jpeg *.tif)"
        )
        if not files:
            return
        for f in files:
            self.list_uploads.addItem(f)
        self._refresh_upload_stats()
        self.log_view.append(f"[{self.cmb_client.currentText()}] Added {len(files)} file(s) to upload queue.\n")

    def _remove_selected_uploads(self):
        for item in self.list_uploads.selectedItems():
            self.list_uploads.takeItem(self.list_uploads.row(item))
        self._refresh_upload_stats()

    def _simulate_import_to_corpus(self):
        count = self.list_uploads.count()
        if count == 0:
            self.chat_view.append("<i>No files queued for import.</i>")
            return
        self.chat_view.append(f"<i>Importing {count} file(s) to corpus… (stub)</i>")
        self.list_uploads.clear()
        self._refresh_upload_stats()
        self.log_view.append(f"[{self.cmb_client.currentText()}] Imported queued files into corpus (stub).\n")

    def _refresh_upload_stats(self):
        self.lbl_upload_stats.setText(f"Queued: {self.list_uploads.count()}")


def get_tool_spec() -> ToolSpec:
    """
    Plugin loader expects (id, name, factory). Variants handled.
    """
    def factory():
        return ClientBriefingRoomPage()

    try:
        return ToolSpec("insured_intelligence_desk", "Insured Intelligence Desk", factory)
    except TypeError:
        try:
            return ToolSpec("insured_intelligence_desk", "Insured Intelligence Desk", None, factory)
        except TypeError:
            return ToolSpec(id="insured_intelligence_desk", name="Insured Intelligence Desk", factory=factory)
