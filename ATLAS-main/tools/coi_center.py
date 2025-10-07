# tools/coi_center.py
from pathlib import Path
from dataclasses import dataclass
from typing import List, Optional

import json
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

from PyQt5.QtCore import Qt, QTimer, QSettings, QUrl
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QLabel, QPushButton, QLineEdit,
    QComboBox, QCheckBox, QTextEdit, QTabWidget, QFileDialog, QGroupBox, QMessageBox,
    QInputDialog, QDialog, QDialogButtonBox, QTreeWidget, QTreeWidgetItem,
    QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem, QHeaderView
)
from PyQt5.QtGui import QDesktopServices

from plugin_api import ToolSpec, BaseToolPage

# Backend API
from services.coi_backend import (
    # legacy account helpers retained for UI actions if needed
    list_accounts, ensure_account_dirs, add_account, delete_account,
    add_email_to_account, delete_email_from_account, get_accounts_index,
    account_root_path, add_files_to_account, add_folder_to_account,
    import_and_auto_analyze,
    # Option A master-config helpers
    resolve_account_by_sender_master, list_accounts_from_fs_and_csv,
    MASTER_CONFIG_DIR
)


@dataclass
class COIItem:
    thread_id: str
    subject: str
    from_addr: str
    age_minutes: int
    status: str  # reserved


class COICenterPage(BaseToolPage):
    """
    COI Center — Inbox triage • Attachments • Requirements • DoO • Notes
    Uses Accounts/<Account> structure:
      - Emails/Threads/<thr>/meta.json, analysis.json, doo.txt
      - Attachments/<thr>/*
      - Corpus/*
    Account resolution uses master CSV at Accounts/_config/senders.csv (Option A).
    """
    def __init__(self, parent=None):
        super().__init__("COI Center", "Inbox triage • Attachments • Requirements • DoO", parent)

        self.settings = QSettings("ATLAS", "COI Center")
        self._seen_threads = set(self.settings.value("seen_threads", [], type=list))
        self._inbox_items: List[COIItem] = []
        self._current_thread: Optional[COIItem] = None
        self._attach_pages_total = 0
        self._attach_page_cur = 0

        # ===== Top workspace bar =====
        top = QGroupBox()
        top_l = QHBoxLayout(top)

        lbl_account = QLabel("Account"); f = lbl_account.font(); f.setBold(True); lbl_account.setFont(f)
        self.cmb_account = QComboBox(); self.cmb_account.setMinimumWidth(260)
        self._refresh_account_combo()

        self.btn_add_account = QPushButton("Add Account")
        self.btn_view_accounts = QPushButton("View Accounts…")
        self.btn_open_config = QPushButton("Open Config")  # NEW

        self.btn_refresh = QPushButton("Refresh")
        self.chk_watch = QCheckBox("Watch Inbox")
        lbl_model = QLabel("Model"); f2 = lbl_model.font(); f2.setBold(True); lbl_model.setFont(f2)
        self.cmb_model = QComboBox(); self.cmb_model.addItems(["OpenAI (prod)", "OpenAI (dev)", "Local"])

        top_l.addWidget(lbl_account); top_l.addWidget(self.cmb_account)
        top_l.addWidget(self.btn_add_account); top_l.addWidget(self.btn_view_accounts); top_l.addWidget(self.btn_open_config)
        top_l.addStretch(1)
        top_l.addWidget(self.btn_refresh); top_l.addWidget(self.chk_watch)
        top_l.addWidget(lbl_model); top_l.addWidget(self.cmb_model)
        self.content.addWidget(top)

        # Small status path
        path_label = QLabel(f"Config: {MASTER_CONFIG_DIR}")
        pf = path_label.font(); pf.setPointSize(max(8, pf.pointSize()-1)); path_label.setFont(pf)
        path_label.setStyleSheet("color:#666;")
        self.content.addWidget(path_label)

        # ===== Main vertical splitter =====
        main_vsplit = QSplitter(Qt.Vertical); main_vsplit.setChildrenCollapsible(False)
        self.content.addWidget(main_vsplit, 1)

        # --- COI Requests table ---
        top_box = QGroupBox()
        top_box_l = QVBoxLayout(top_box)

        title = QLabel("COI Requests")
        tf = title.font(); tf.setPointSize(tf.pointSize() + 1); tf.setBold(True); title.setFont(tf)
        top_box_l.addWidget(title)

        search_row = QHBoxLayout()
        lbl_search = QLabel("Search"); sf = lbl_search.font(); sf.setBold(True); lbl_search.setFont(sf)
        self.txt_search = QLineEdit(); self.txt_search.setPlaceholderText("Filter by subject or sender…")
        search_row.addWidget(lbl_search); search_row.addWidget(self.txt_search, 1)
        top_box_l.addLayout(search_row)

        self.tbl_inbox = QTableWidget()
        self.tbl_inbox.setColumnCount(4)
        self.tbl_inbox.setHorizontalHeaderLabels(["Subject", "Sender", "Account", "Age"])
        self.tbl_inbox.verticalHeader().setVisible(False)
        self.tbl_inbox.setSelectionBehavior(self.tbl_inbox.SelectRows)
        self.tbl_inbox.setSelectionMode(self.tbl_inbox.SingleSelection)
        self.tbl_inbox.setEditTriggers(self.tbl_inbox.NoEditTriggers)

        # Header sizing: Subject stretches; others auto-fit (Age stays compact)
        hdr = self.tbl_inbox.horizontalHeader()
        hf = hdr.font(); hf.setBold(True); hdr.setFont(hf)
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)            # Subject
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)   # Sender
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)   # Account
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)   # Age

        top_box_l.addWidget(self.tbl_inbox, 1)
        main_vsplit.addWidget(top_box)

        # --- Account Corpus bar ---
        corpus_bar = QGroupBox()
        corpus_l = QHBoxLayout(corpus_bar)
        lbl_corpus = QLabel("Account Corpus"); cf = lbl_corpus.font(); cf.setBold(True); lbl_corpus.setFont(cf)
        self.btn_add_files = QPushButton("Add Files")
        self.btn_add_folder = QPushButton("Add Folder")
        self.btn_import_drop = QPushButton("Import from Drop")
        corpus_l.addWidget(lbl_corpus); corpus_l.addStretch(1)
        corpus_l.addWidget(self.btn_add_files); corpus_l.addWidget(self.btn_add_folder); corpus_l.addWidget(self.btn_import_drop)
        main_vsplit.addWidget(corpus_bar)

        # --- Bottom tabs ---
        self.tabs = QTabWidget()
        tbf = self.tabs.tabBar().font(); tbf.setBold(True); self.tabs.tabBar().setFont(tbf)

        # Email Preview tab (with Open Email button)
        preview_box = QGroupBox()
        pv_l = QVBoxLayout(preview_box)
        pv_btns = QHBoxLayout()
        self.btn_open_email = QPushButton("Open Email")
        pv_btns.addWidget(self.btn_open_email); pv_btns.addStretch(1)
        self.txt_email = QTextEdit(); self.txt_email.setReadOnly(True)
        pv_l.addLayout(pv_btns); pv_l.addWidget(self.txt_email, 1)
        self.tabs.addTab(preview_box, "Email Preview")

        # Email Attachments tab
        attach_wrap = QSplitter(Qt.Horizontal); attach_wrap.setChildrenCollapsible(False)
        self.lst_attachments = QListWidget()
        right_attach = QGroupBox("Attachment Preview")
        ra_l = QVBoxLayout(right_attach)
        self.lbl_attach_name = QLabel("(no attachment selected)")
        self.lbl_attach_pages = QLabel("Page 0 / 0")
        nav = QHBoxLayout()
        self.btn_prev_page = QPushButton("◀ Prev"); self.btn_next_page = QPushButton("Next ▶")
        nav.addWidget(self.btn_prev_page); nav.addWidget(self.btn_next_page); nav.addStretch(1)
        self.txt_attach_preview = QTextEdit(); self.txt_attach_preview.setReadOnly(True)
        ra_l.addWidget(self.lbl_attach_name); ra_l.addWidget(self.lbl_attach_pages); ra_l.addLayout(nav); ra_l.addWidget(self.txt_attach_preview, 1)
        attach_wrap.addWidget(self.lst_attachments); attach_wrap.addWidget(right_attach)
        attach_wrap.setSizes([260, 640])
        self.tabs.addTab(attach_wrap, "Email Attachments")

        # Extracted Requirements
        self.txt_requirements = QTextEdit()
        self.txt_requirements.setPlaceholderText("Structured requirements (limits, holder, AI/WOS/PNC, notice)…")
        self.tabs.addTab(self.txt_requirements, "Extracted Requirements")

        # DoO Draft
        self.txt_doo = QTextEdit()
        self.txt_doo.setPlaceholderText("Draft Description of Operations (exclude non-compliant wording).")
        self.tabs.addTab(self.txt_doo, "DoO Draft")

        # Notes
        self.txt_notes = QTextEdit()
        self.txt_notes.setPlaceholderText("Analyst/internal notes for this thread.")
        self.tabs.addTab(self.txt_notes, "Notes")

        main_vsplit.addWidget(self.tabs)
        main_vsplit.setSizes([380, 64, 520])

        # ===== Signals =====
        self.btn_refresh.clicked.connect(self._refresh_clicked)
        self.txt_search.textChanged.connect(self._apply_filter)
        self.tbl_inbox.itemSelectionChanged.connect(self._on_select_thread)

        self.btn_add_files.clicked.connect(self._add_files)
        self.btn_add_folder.clicked.connect(self._add_folder)
        self.btn_import_drop.clicked.connect(self._import_from_drop)

        self.btn_add_account.clicked.connect(self._add_account_dialog)
        self.btn_view_accounts.clicked.connect(self._manage_accounts_dialog)
        self.btn_open_config.clicked.connect(self._open_config)

        self.cmb_account.currentIndexChanged.connect(self._on_account_changed)

        self.lst_attachments.itemSelectionChanged.connect(self._on_select_attachment)
        self.btn_prev_page.clicked.connect(self._prev_attach_page)
        self.btn_next_page.clicked.connect(self._next_attach_page)

        self.btn_open_email.clicked.connect(self._open_email_file)

        # Initial load from storage (no demo)
        self._csv_mtime = 0.0
        self._rebuild_inbox_from_storage()

        # Watch Inbox polling
        self.timer = QTimer(self); self.timer.setInterval(15_000)
        self.timer.timeout.connect(self._maybe_poll)
        self.chk_watch.stateChanged.connect(self._toggle_watch)

        # Mapping watcher (poll CSV)
        self._mapping_timer = QTimer(self)
        self._mapping_timer.setInterval(10_000)
        self._mapping_timer.timeout.connect(self._maybe_reload_mapping)
        self._mapping_timer.start()

    # ===== Helpers: datetime & storage scan =====
    def _parse_email_date(self, s: str) -> Optional[datetime]:
        if not s:
            return None
        try:
            dt = parsedate_to_datetime(s)
            if dt and dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt
        except Exception:
            return None

    def _minutes_ago(self, dt: Optional[datetime], fallback_path: Optional[Path] = None) -> int:
        if dt is None and fallback_path and fallback_path.exists():
            newest = max((p.stat().st_mtime for p in fallback_path.glob("*") if p.is_file()), default=None)
            if newest is not None:
                dt = datetime.fromtimestamp(newest, tz=timezone.utc)
        if dt is None:
            return 0
        now = datetime.now(tz=timezone.utc)
        delta = now - dt
        return max(int(delta.total_seconds() // 60), 0)

    def _format_age(self, minutes: int) -> str:
        if minutes < 60:
            return f"{minutes} min"
        hours, rem = divmod(minutes, 60)
        if hours < 24:
            return f"{hours} h {rem} m" if rem else f"{hours} h"
        days, rem_h = divmod(hours, 24)
        if rem_h == 0 and rem == 0:
            return f"{days} d"
        if rem == 0:
            return f"{days} d {rem_h} h"
        return f"{days} d {rem_h} h {rem} m"

    # ===== Inbox build (from storage) =====
    def _rebuild_inbox_from_storage(self):
        """Scan Emails/Threads/*/meta.json and build the inbox."""
        acc = self._current_account_name()
        items: List[COIItem] = []
        if acc:
            thr_root = account_root_path(acc) / "Emails" / "Threads"
            if thr_root.exists():
                for thr_dir in sorted(thr_root.glob("*")):
                    if not thr_dir.is_dir():
                        continue
                    meta_path = thr_dir / "meta.json"
                    if not meta_path.exists():
                        continue
                    try:
                        meta = json.loads(meta_path.read_text(encoding="utf-8"))
                    except Exception:
                        meta = {}
                    subject = (meta.get("Subject") or "").strip() or "(no subject)"
                    sender = (meta.get("From") or "").strip()
                    dt = self._parse_email_date(meta.get("Date") or "")
                    # Age fallback: newest file in thread dir
                    age_minutes = self._minutes_ago(dt, fallback_path=thr_dir)
                    items.append(COIItem(
                        thread_id=thr_dir.name,
                        subject=subject,
                        from_addr=sender,
                        age_minutes=age_minutes,
                        status="",
                    ))
        self._populate_inbox(items)

    # ===== UI table utils =====
    def _populate_inbox(self, items: List[COIItem]):
        self._inbox_items = items[:]
        self._render_inbox_table(items)

    def _render_inbox_table(self, items: List[COIItem]):
        self.tbl_inbox.setSortingEnabled(False)
        self.tbl_inbox.setRowCount(len(items))

        for r, it in enumerate(items):
            subj = QTableWidgetItem(it.subject)
            sender = QTableWidgetItem(it.from_addr)
            # Account via master CSV mapping
            account = QTableWidgetItem(resolve_account_by_sender_master(it.from_addr))
            age = QTableWidgetItem(self._format_age(it.age_minutes))

            sf = subj.font(); sf.setBold(True); subj.setFont(sf)
            subj.setData(Qt.UserRole, it)
            for cell in (subj, sender, account, age):
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)

            self.tbl_inbox.setItem(r, 0, subj)
            self.tbl_inbox.setItem(r, 1, sender)
            self.tbl_inbox.setItem(r, 2, account)
            self.tbl_inbox.setItem(r, 3, age)

        self.tbl_inbox.setSortingEnabled(True)

    def _apply_filter(self, text: str):
        q = text.lower().strip()
        if not q:
            self._render_inbox_table(self._inbox_items)
            return
        filtered = [it for it in self._inbox_items if q in it.subject.lower() or q in it.from_addr.lower()]
        self._render_inbox_table(filtered)

    # ===== Row selection =====
    def _on_select_thread(self):
        row = self.tbl_inbox.currentRow()
        if row < 0:
            return
        cell = self.tbl_inbox.item(row, 0)
        if not cell:
            return
        it: COIItem = cell.data(Qt.UserRole)
        if not it:
            return

        self._current_thread = it
        self._seen_threads.add(it.thread_id)

        acc = self._current_account_name()
        body_text = ""
        # Load meta.json for preview
        if acc:
            meta_path = account_root_path(acc) / "Emails" / "Threads" / it.thread_id / "meta.json"
            if meta_path.exists():
                try:
                    meta = json.loads(meta_path.read_text(encoding="utf-8"))
                    body_text = meta.get("BodyText", "") or ""
                except Exception:
                    body_text = ""

        self.txt_email.setPlainText(
            f"Subject: {it.subject}\nFrom: {it.from_addr}\nAccount: {resolve_account_by_sender_master(it.from_addr)}\n"
            f"Age: {self._format_age(it.age_minutes)}\n\n"
            f"{body_text}"
        )

        # Load analysis artifacts
        if not acc:
            self.txt_requirements.clear(); self.txt_doo.clear(); self.txt_notes.clear()
            return

        base = account_root_path(acc) / "Emails" / "Threads" / it.thread_id
        req = base / "analysis.json"
        doo = base / "doo.txt"
        notes = base / "notes.txt"

        self.txt_requirements.setPlainText(req.read_text(encoding="utf-8", errors="ignore") if req.exists() else "")
        self.txt_doo.setPlainText(doo.read_text(encoding="utf-8", errors="ignore") if doo.exists() else "")
        self.txt_notes.setPlainText(notes.read_text(encoding="utf-8", errors="ignore") if notes.exists() else "")

        # Populate attachments from Attachments/<thread-id>
        self._load_thread_attachments()

    # ===== Refresh / Watch =====
    def _refresh_clicked(self):
        acc = self._current_account_name()
        if acc:
            import_and_auto_analyze(acc)
        self._rebuild_inbox_from_storage()

    def _toggle_watch(self, state: int):
        if state == Qt.Checked:
            if not self.timer.isActive():
                self.timer.start()
        else:
            self.timer.stop()

    def _maybe_poll(self):
        acc = self._current_account_name()
        if acc:
            import_and_auto_analyze(acc)
        self._rebuild_inbox_from_storage()

    # ===== Mapping CSV watcher =====
    def _maybe_reload_mapping(self):
        try:
            mtime = (MASTER_CONFIG_DIR / "senders.csv").stat().st_mtime
        except Exception:
            mtime = 0.0
        if mtime != self._csv_mtime:
            self._csv_mtime = mtime
            # Refresh account dropdown + re-render table so Account column updates
            self._refresh_account_combo()
            self._render_inbox_table(self._inbox_items)

    # ===== Account changes =====
    def _refresh_account_combo(self):
        names = list_accounts_from_fs_and_csv()
        cur = self.cmb_account.currentText()
        self.cmb_account.blockSignals(True)
        self.cmb_account.clear()
        self.cmb_account.addItems(["Select account…"] + names)
        # keep selection if still present
        if cur and cur in names:
            self.cmb_account.setCurrentText(cur)
        self.cmb_account.blockSignals(False)

    def _current_account_name(self) -> Optional[str]:
        name = self.cmb_account.currentText()
        return None if (not name or name == "Select account…") else name

    def _on_account_changed(self):
        name = self._current_account_name()
        if name:
            ensure_account_dirs(name)
            import_and_auto_analyze(name)
        self._rebuild_inbox_from_storage()

    # ===== Accounts manage =====
    def _add_account_dialog(self):
        name, ok = QInputDialog.getText(self, "Add Account", "Account name:")
        if not ok or not name.strip():
            return
        emails_text, ok2 = QInputDialog.getMultiLineText(
            self, "Account Emails (optional legacy index)",
            "Enter one email per line (optional):",
            ""
        )
        if not ok2:
            return
        emails = [e.strip() for e in emails_text.splitlines() if e.strip()]
        add_account(name.strip(), emails)
        ensure_account_dirs(name.strip())
        root_path = account_root_path(name.strip())
        self._refresh_account_combo()
        QMessageBox.information(self, "Account Created", f"Account saved.\nFolder:\n{root_path}")

    def _manage_accounts_dialog(self):
        dlg = QDialog(self); dlg.setWindowTitle("Accounts (legacy index view)")
        v = QVBoxLayout(dlg)
        tree = QTreeWidget(); tree.setHeaderLabels(["Account / Email", "Type"])

        def render_tree():
            tree.clear()
            idx = get_accounts_index()
            for acct in idx.get("accounts", []):
                top = QTreeWidgetItem([acct.get("name", ""), "account"])
                tree.addTopLevelItem(top)
                for e in acct.get("emails", []):
                    top.addChild(QTreeWidgetItem([e, "email"]))
                top.setExpanded(True)

        render_tree()
        v.addWidget(tree, 1)

        row = QHBoxLayout()
        btn_add_acct = QPushButton("Add Account")
        btn_del_acct = QPushButton("Delete Account")
        btn_add_email = QPushButton("Add Email")
        btn_del_email = QPushButton("Delete Email")
        row.addWidget(btn_add_acct); row.addWidget(btn_del_acct); row.addWidget(btn_add_email); row.addWidget(btn_del_email)
        row.addStretch(1)
        v.addLayout(row)

        btns = QDialogButtonBox(QDialogButtonBox.Close)
        v.addWidget(btns)

        def add_acct():
            name, ok = QInputDialog.getText(dlg, "Add Account", "Account name:")
            if not ok or not name.strip():
                return
            add_account(name.strip(), [])
            ensure_account_dirs(name.strip())
            render_tree(); self._refresh_account_combo()

        def del_acct():
            items = tree.selectedItems()
            if not items:
                return
            it = items[0]
            if it.text(1) != "account":
                QMessageBox.information(dlg, "Accounts", "Select an account to delete.")
                return
            name = it.text(0)
            if QMessageBox.question(dlg, "Delete Account", f"Delete account '{name}'?") != QMessageBox.Yes:
                return
            delete_account(name)
            render_tree(); self._refresh_account_combo()

        def add_email():
            items = tree.selectedItems()
            if not items:
                return
            it = items[0]
            target = it.text(0) if it.text(1) == "account" else it.parent().text(0)
            email, ok = QInputDialog.getText(dlg, "Add Email", f"Add email to '{target}':")
            if not ok or not email.strip():
                return
            add_email_to_account(target, email.strip())
            render_tree()

        def del_email():
            items = tree.selectedItems()
            if not items:
                return
            it = items[0]
            if it.text(1) != "email":
                QMessageBox.information(dlg, "Accounts", "Select an email to delete.")
                return
            email = it.text(0); acct = it.parent().text(0)
            delete_email_from_account(acct, email)
            render_tree()

        btn_add_acct.clicked.connect(add_acct)
        btn_del_acct.clicked.connect(del_acct)
        btn_add_email.clicked.connect(add_email)
        btn_del_email.clicked.connect(del_email)
        btns.rejected.connect(dlg.reject)
        dlg.exec_()

    # ===== Open master config =====
    def _open_config(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(MASTER_CONFIG_DIR)))

    # ===== Corpus actions =====
    def _add_files(self):
        acc = self._current_account_name()
        if not acc:
            QMessageBox.warning(self, "Select account", "Choose an account first.")
            return
        files, _ = QFileDialog.getOpenFileNames(self, "Add files to account Corpus", "", "All Files (*.*)")
        if not files:
            return
        added, errors = add_files_to_account(acc, [Path(f) for f in files])
        msg = f"Added {added} file(s) to Corpus."
        if errors: msg += f"\nErrors: {len(errors)}"
        QMessageBox.information(self, "Account Corpus", msg)

    def _add_folder(self):
        acc = self._current_account_name()
        if not acc:
            QMessageBox.warning(self, "Select account", "Choose an account first.")
            return
        folder = QFileDialog.getExistingDirectory(self, "Select folder to add to Corpus", "")
        if not folder:
            return
        added, errors = add_folder_to_account(acc, Path(folder))
        msg = f"Added {added} file(s) from folder."
        if errors: msg += f"\nErrors: {len(errors)}"
        QMessageBox.information(self, "Account Corpus", msg)

    def _import_from_drop(self):
        acc = self._current_account_name()
        if not acc:
            QMessageBox.warning(self, "Select account", "Choose an account first.")
            return
        stats = import_and_auto_analyze(acc)
        QMessageBox.information(
            self, "Drop Import",
            f"Imported {stats['emails']} email(s), {stats['attachments']} attachment(s) "
            f"across {stats['threads']} thread(s).\nAuto-analyzed: {stats.get('analyzed', 0)}"
        )
        self._rebuild_inbox_from_storage()

    # ===== Attachments =====
    def _load_thread_attachments(self):
        self.lst_attachments.clear()
        it = self._current_thread
        acc = self._current_account_name()
        if not (it and acc):
            return
        att = account_root_path(acc) / "Attachments" / it.thread_id
        if not att.exists():
            self.lbl_attach_name.setText("(no attachment selected)")
            self.lbl_attach_pages.setText("Page 0 / 0")
            self.txt_attach_preview.clear()
            self._attach_pages_total = 0
            self._attach_page_cur = 0
            return
        for p in sorted(att.glob("*")):
            if not p.is_file():
                continue
            self.lst_attachments.addItem(QListWidgetItem(p.name))
        if self.lst_attachments.count() > 0:
            self.lst_attachments.setCurrentRow(0)
        else:
            self.lbl_attach_name.setText("(no attachment selected)")
            self.lbl_attach_pages.setText("Page 0 / 0")
            self.txt_attach_preview.clear()
            self._attach_pages_total = 0
            self._attach_page_cur = 0

    def _on_select_attachment(self):
        it = self.lst_attachments.currentItem()
        if not it:
            self.lbl_attach_name.setText("(no attachment selected)")
            self.lbl_attach_pages.setText("Page 0 / 0")
            self.txt_attach_preview.clear()
            self._attach_pages_total = 0
            self._attach_page_cur = 0
            return
        name = it.text()
        total = 3 if name.lower().endswith(".pdf") else 1  # placeholder paging
        self._attach_pages_total = total
        self._attach_page_cur = 1 if total > 0 else 0
        self.lbl_attach_name.setText(name)
        self.lbl_attach_pages.setText(f"Page {self._attach_page_cur} / {self._attach_pages_total}")
        self.txt_attach_preview.setPlainText(f"[Preview stub for '{name}']\nUse Prev/Next to flip pages.")

    def _prev_attach_page(self):
        if self._attach_pages_total <= 1: return
        if self._attach_page_cur > 1:
            self._attach_page_cur -= 1
            self.lbl_attach_pages.setText(f"Page {self._attach_page_cur} / {self._attach_pages_total}")

    def _next_attach_page(self):
        if self._attach_pages_total <= 1: return
        if self._attach_page_cur < self._attach_pages_total:
            self._attach_page_cur += 1
            self.lbl_attach_pages.setText(f"Page {self._attach_page_cur} / {self._attach_pages_total}")

    # ===== Open Email (uses OS default handler) =====
    def _open_email_file(self):
        it = self._current_thread
        acc = self._current_account_name()
        if not (it and acc):
            QMessageBox.information(self, "Open Email", "Select a request and account first.")
            return
        thr_dir = account_root_path(acc) / "Emails" / "Threads" / it.thread_id
        if not thr_dir.exists():
            QMessageBox.information(self, "Open Email", "No saved email found for this request.")
            return
        candidates = sorted([p for p in thr_dir.glob("*") if p.suffix.lower() in (".eml", ".msg")],
                            key=lambda p: p.stat().st_mtime, reverse=True)
        if not candidates:
            QMessageBox.information(self, "Open Email", "No .eml or .msg file found.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(candidates[0])))


# ---- Tool registration ----
def get_tool_spec() -> ToolSpec:
    def factory() -> QWidget:
        return COICenterPage()
    return ToolSpec(
        id="coi_center",
        name="COI Center",
        factory=factory,
        icon=None,
        order=200,
    )
