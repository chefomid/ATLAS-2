# services/coi_backend.py
from pathlib import Path
from typing import List, Dict, Any, Tuple
import json
import shutil
import hashlib
import time
import os
import re

from email import policy
from email.parser import BytesParser

# Optional .msg parser (Windows/Outlook). If unavailable, .msg parsing degrades gracefully.
try:
    import extract_msg  # pip install extract-msg
except Exception:
    extract_msg = None


# =========================
# Storage roots / constants
# =========================
# Point storage to your Accounts folder (can be overridden by COI_DATA_ROOT env var)
DEFAULT_ROOT = r"C:\Users\Orcc_\OneDrive\Desktop\USI Automations\RAS Prem Alloc_Alg\RAS Alg - fork\Accounts"
DATA_ROOT = Path(os.environ.get("COI_DATA_ROOT", DEFAULT_ROOT))

# --- Master config (Option A) ---
MASTER_CONFIG_DIR = DATA_ROOT / "_config"
MASTER_SENDERS = MASTER_CONFIG_DIR / "senders.csv"
MASTER_COUNTERS = MASTER_CONFIG_DIR / "counters.csv"  # optional, not required for current flow

INDEX_PATH = DATA_ROOT / "accounts.json"  # retained for compatibility (not critical with Option A)

# Windows-safe folder name
_INVALID_CHARS = re.compile(r'[<>:"/\\|?*]+')


def safe_account_folder(name: str) -> str:
    n = name.strip()
    n = _INVALID_CHARS.sub("_", n)
    n = n.rstrip(". ") or "Account"
    return n


# =================
# JSON I/O helpers
# =================
def _read_json(path: Path, default: Dict[str, Any]) -> Dict[str, Any]:
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _write_json(path: Path, obj: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2), encoding="utf-8")


# ==========================
# Master config helpers (CSV)
# ==========================
def _ensure_master_config_dirs() -> None:
    MASTER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not MASTER_SENDERS.exists():
        MASTER_SENDERS.write_text("# sender_email,account_name\n", encoding="utf-8")
    if not MASTER_COUNTERS.exists():
        MASTER_COUNTERS.write_text("# account_name,last_request_number\n", encoding="utf-8")


def load_sender_map_from_master() -> Dict[str, str]:
    """
    Reads _config/senders.csv where each non-comment line is:
      sender_email,account_name
    Returns dict: {lower(email): account_name}
    """
    _ensure_master_config_dirs()
    mp: Dict[str, str] = {}
    if not MASTER_SENDERS.exists():
        return mp
    for line in MASTER_SENDERS.read_text(encoding="utf-8", errors="ignore").splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        parts = s.split(",", 1)
        if len(parts) < 2:
            continue
        email = parts[0].strip().lower()
        acct = parts[1].strip()
        if email and acct:
            mp[email] = acct
    return mp


def resolve_account_by_sender_master(email: str) -> str:
    """Lookup account via _config/senders.csv. Returns '' if unknown."""
    e = (email or "").strip().lower()
    if not e:
        return ""
    mp = load_sender_map_from_master()
    return mp.get(e, "")


def list_accounts_from_fs_and_csv() -> List[str]:
    """
    Build account list from:
      - existing folders under DATA_ROOT (Accounts/<Account>/…)
      - account names referenced in _config/senders.csv
    """
    _ensure_master_config_dirs()
    names = set()
    if DATA_ROOT.exists():
        for p in DATA_ROOT.iterdir():
            if p.is_dir() and p.name not in ("_config",):
                names.add(p.name)
    for acct in load_sender_map_from_master().values():
        if acct:
            names.add(acct)
    return sorted(names)


# ==========================
# Account index (legacy JSON)
# ==========================
def get_accounts_index() -> Dict[str, Any]:
    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    return _read_json(INDEX_PATH, {"accounts": []})


def save_accounts_index(idx: Dict[str, Any]) -> None:
    _write_json(INDEX_PATH, idx)


def list_accounts() -> List[Dict[str, Any]]:
    """List accounts from legacy index (kept for compatibility)."""
    return get_accounts_index().get("accounts", [])


def add_account(name: str, emails: List[str]) -> None:
    """Create/update an account in legacy index and ensure folders exist."""
    idx = get_accounts_index()
    nm = name.strip()
    for acct in idx["accounts"]:
        if acct.get("name", "").strip().lower() == nm.lower():
            current = {e.strip().lower() for e in acct.get("emails", [])}
            for e in emails:
                current.add(e.strip().lower())
            acct["emails"] = sorted(current)
            save_accounts_index(idx)
            ensure_account_dirs(nm)
            return
    idx["accounts"].append({"name": nm, "emails": sorted({e.strip().lower() for e in emails})})
    save_accounts_index(idx)
    ensure_account_dirs(nm)


def delete_account(name: str) -> None:
    idx = get_accounts_index()
    idx["accounts"] = [a for a in idx.get("accounts", []) if a.get("name") != name]
    save_accounts_index(idx)
    root = account_root_path(name)
    if root.exists():
        shutil.rmtree(root, ignore_errors=True)


def add_email_to_account(name: str, email: str) -> None:
    idx = get_accounts_index()
    for acct in idx.get("accounts", []):
        if acct.get("name") == name:
            current = {e.strip().lower() for e in acct.get("emails", [])}
            current.add(email.strip().lower())
            acct["emails"] = sorted(current)
            break
    save_accounts_index(idx)


def delete_email_from_account(name: str, email: str) -> None:
    idx = get_accounts_index()
    for acct in idx.get("accounts", []):
        if acct.get("name") == name:
            acct["emails"] = [e for e in acct.get("emails", []) if e.lower() != email.strip().lower()]
            break
    save_accounts_index(idx)


# ==========================
# Account filesystem helpers
# ==========================
def account_root_path(name: str) -> Path:
    return DATA_ROOT / safe_account_folder(name)


def ensure_account_dirs(name: str) -> None:
    """Create standardized folders for an account."""
    root = account_root_path(name)
    (root / "Corpus").mkdir(parents=True, exist_ok=True)
    (root / "Emails" / "Drop").mkdir(parents=True, exist_ok=True)
    (root / "Emails" / "Threads").mkdir(parents=True, exist_ok=True)
    (root / "Attachments").mkdir(parents=True, exist_ok=True)


# ============
# Corpus ops
# ============
def add_files_to_account(name: str, files: List[Path]) -> Tuple[int, List[str]]:
    """Copy files into account Corpus/."""
    ensure_account_dirs(name)
    dst = account_root_path(name) / "Corpus"
    dst.mkdir(parents=True, exist_ok=True)
    added = 0
    errors: List[str] = []
    for p in files:
        try:
            (dst / p.name).write_bytes(p.read_bytes())
            added += 1
        except Exception as ex:
            errors.append(f"{p.name}: {ex}")
    return added, errors


def add_folder_to_account(name: str, folder: Path) -> Tuple[int, List[str]]:
    """Recursively copy a folder into Corpus/ (preserve subfolders)."""
    ensure_account_dirs(name)
    dst_root = account_root_path(name) / "Corpus"
    dst_root.mkdir(parents=True, exist_ok=True)
    added, errors = 0, []
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        rel = p.relative_to(folder)
        dst = dst_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            dst.write_bytes(p.read_bytes())
            added += 1
        except Exception as ex:
            errors.append(f"{rel.as_posix()}: {ex}")
    return added, errors


# ============================
# Outlook drop folder ingestion
# ============================
def account_drop_path(name: str) -> Path:
    """Folder where Outlook rule saves .eml/.msg for this account."""
    return account_root_path(name) / "Emails" / "Drop"


def _safe_thread_id(msg_dict: Dict[str, Any]) -> str:
    """Prefer Message-Id/Thread-Index; fallback to hash of subject|from|date."""
    for k in ("Message-Id", "Thread-Index", "Thread-Id", "In-Reply-To"):
        v = msg_dict.get(k) or msg_dict.get(k.lower())
        if v:
            h = hashlib.sha1(v.encode("utf-8", errors="ignore")).hexdigest()[:16]
            return f"thr_{h}"
    subject = (msg_dict.get("Subject") or msg_dict.get("subject") or "").strip()
    sender = (msg_dict.get("From") or msg_dict.get("from") or "").strip()
    date = (msg_dict.get("Date") or msg_dict.get("date") or "").strip()
    raw = f"{subject}|{sender}|{date}"
    h = hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()[:16]
    return f"thr_{h}"


def _parse_eml(eml_path: Path) -> Dict[str, Any]:
    data = eml_path.read_bytes()
    msg = BytesParser(policy=policy.default).parsebytes(data)
    info: Dict[str, Any] = {
        "Subject": msg.get("Subject", ""),
        "From": msg.get("From", ""),
        "To": msg.get("To", ""),
        "Cc": msg.get("Cc", ""),
        "Date": msg.get("Date", ""),
        "Message-Id": msg.get("Message-Id", ""),
        "In-Reply-To": msg.get("In-Reply-To", ""),
        "Thread-Index": msg.get("Thread-Index", ""),
        "BodyText": msg.get_body(preferencelist=("plain", "html")).get_content() if msg.get_body() else "",
        "Attachments": []
    }
    for part in msg.iter_attachments():
        fname = part.get_filename() or "attachment.bin"
        try:
            payload = part.get_payload(decode=True) or b""
        except Exception:
            payload = b""
        info["Attachments"].append((fname, payload))
    return info


def _parse_msg(msg_path: Path) -> Dict[str, Any]:
    if extract_msg is None:
        return {
            "Subject": msg_path.stem,
            "From": "",
            "To": "",
            "Cc": "",
            "Date": time.ctime(msg_path.stat().st_mtime),
            "Message-Id": "",
            "In-Reply-To": "",
            "Thread-Index": "",
            "BodyText": "",
            "Attachments": []
        }
    m = extract_msg.Message(str(msg_path))
    info = {
        "Subject": m.subject or "",
        "From": m.sender or "",
        "To": ", ".join(m.to) if m.to else "",
        "Cc": ", ".join(m.cc) if m.cc else "",
        "Date": m.date or "",
        "Message-Id": "",
        "In-Reply-To": "",
        "Thread-Index": "",
        "BodyText": m.body or "",
        "Attachments": []
    }
    for att in m.attachments:
        info["Attachments"].append((att.longFilename or att.shortFilename or "attachment.bin", att.data or b""))
    return info


def _thread_dir(account: str, thread_id: str) -> Path:
    return account_root_path(account) / "Emails" / "Threads" / thread_id


def _attachments_dir(account: str, thread_id: str) -> Path:
    return account_root_path(account) / "Attachments" / thread_id


def import_drop_folder(name: str) -> Dict[str, int]:
    """
    Import .eml/.msg from Emails/Drop into Emails/Threads/<thr>/.
    Save sidecar meta.json + original message. Route attachments to Attachments/<thr>/.
    """
    drop = account_drop_path(name)
    drop.mkdir(parents=True, exist_ok=True)

    emails = attachments = 0
    threads_touched = set()

    for p in sorted(drop.glob("**/*")):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext not in (".eml", ".msg"):
            continue

        try:
            info = _parse_eml(p) if ext == ".eml" else _parse_msg(p)
            thr_id = _safe_thread_id(info)
            threads_touched.add(thr_id)

            thr_root = _thread_dir(name, thr_id)
            thr_root.mkdir(parents=True, exist_ok=True)

            # Save original message (dedupe by name)
            dst_msg = thr_root / p.name
            if not dst_msg.exists():
                dst_msg.write_bytes(p.read_bytes())
            emails += 1

            # Sidecar meta.json for quick UI
            meta = {k: info.get(k) for k in (
                "Subject", "From", "To", "Cc", "Date", "Message-Id", "In-Reply-To", "Thread-Index", "BodyText"
            )}
            (thr_root / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

            # Attachments → Attachments/<thr-id>/
            att_root = _attachments_dir(name, thr_id)
            att_root.mkdir(parents=True, exist_ok=True)
            for fname, blob in info.get("Attachments", []):
                if not blob:
                    continue
                dst = att_root / fname
                if dst.exists():
                    base = dst.stem
                    ext2 = dst.suffix
                    i = 1
                    while True:
                        cand = dst.parent / f"{base}_{i}{ext2}"
                        if not cand.exists():
                            dst = cand
                            break
                        i += 1
                dst.write_bytes(blob)
                attachments += 1

        except Exception:
            # Log in production; keep importing
            continue

    return {"emails": emails, "attachments": attachments, "threads": len(threads_touched)}


# =========================
# Analysis helpers (auto)
# =========================
def has_artifacts(account: str, thread_id: str) -> bool:
    base = _thread_dir(account, thread_id)
    return (base / "analysis.json").exists() or (base / "doo.txt").exists()


def save_request_artifacts(account: str, thread_id: str,
                           analysis_json: str = "",
                           doo_text: str = "",
                           notes_text: str = "") -> None:
    base = _thread_dir(account, thread_id)
    base.mkdir(parents=True, exist_ok=True)
    if analysis_json and analysis_json.strip():
        (base / "analysis.json").write_text(analysis_json, encoding="utf-8")
    if doo_text and doo_text.strip():
        (base / "doo.txt").write_text(doo_text, encoding="utf-8")
    if notes_text and notes_text.strip():
        (base / "notes.txt").write_text(notes_text, encoding="utf-8")


def _load_latest_meta(account: str, thread_id: str) -> Dict[str, Any]:
    meta_path = _thread_dir(account, thread_id) / "meta.json"
    if not meta_path.exists():
        return {}
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _list_attachment_names(account: str, thread_id: str) -> List[str]:
    att = _attachments_dir(account, thread_id)
    if not att.exists():
        return []
    return [p.name for p in sorted(att.glob("*")) if p.is_file()]


def _cheap_extract_requirements(text: str) -> Dict[str, Any]:
    """Lightweight heuristic (placeholder for AI)."""
    t = (text or "").lower()
    req = {"holder": "", "terms": [], "limits": {}}

    if "additional insured" in t: req["terms"].append("Additional Insured")
    if "waiver of subrogation" in t: req["terms"].append("Waiver of Subrogation")
    if "primary & noncontributory" in t or "primary and noncontributory" in t:
        req["terms"].append("Primary & Noncontributory")
    if "30 days notice" in t or "thirty (30) days notice" in t or "30-day notice" in t:
        req["terms"].append("30 days notice of cancellation")

    import re as _re
    m = _re.search(r"(each occurrence|occurrence limit)[^\d]{0,20}(\$?\d[\d,]{2,})", t)
    if m:
        req["limits"]["GL Each Occurrence"] = m.group(2).replace("$", "")
    return req


def auto_analyze_request(account: str, thread_id: str) -> None:
    meta = _load_latest_meta(account, thread_id)
    body = meta.get("BodyText", "") or ""
    subj = meta.get("Subject", "") or ""
    holder_guess = ""

    req = _cheap_extract_requirements(subj + "\n" + body)
    if holder_guess:
        req["holder"] = holder_guess

    doo_lines = [
        "Description of Operations:",
        "Certificate provided in connection with the referenced project/contract.",
        "Coverage is afforded only as required by written contract and no broader than the policy terms and endorsements.",
    ]
    if "Additional Insured" in req["terms"]:
        doo_lines.append("Additional Insured status is provided where required by written contract.")
    if "Waiver of Subrogation" in req["terms"]:
        doo_lines.append("Waiver of Subrogation applies where required by written contract.")
    if any("Primary" in s for s in req["terms"]):
        doo_lines.append("Coverage applies on a primary and noncontributory basis where required by written contract.")

    analysis_json = json.dumps({
        "subject": subj,
        "from": meta.get("From", ""),
        "date": meta.get("Date", ""),
        "requested": req,
        "attachments": _list_attachment_names(account, thread_id),
        "notes": "Auto-generated via heuristic stub. Replace with AI + policy form matching.",
    }, indent=2)

    save_request_artifacts(
        account, thread_id,
        analysis_json=analysis_json,
        doo_text="\n".join(doo_lines),
        notes_text=""
    )


def import_and_auto_analyze(account: str) -> Dict[str, int]:
    """
    1) Import .eml/.msg from Emails/Drop into Emails/Threads/<thr> and Attachments/<thr>.
    2) Auto-analyze any thread missing artifacts.
    """
    _ensure_master_config_dirs()
    stats = import_drop_folder(account)
    analyzed = 0
    thr_root = account_root_path(account) / "Emails" / "Threads"
    if thr_root.exists():
        for thr_dir in sorted([p for p in thr_root.iterdir() if p.is_dir()]):
            thr = thr_dir.name
            if not has_artifacts(account, thr):
                try:
                    auto_analyze_request(account, thr)
                    analyzed += 1
                except Exception:
                    pass
    stats["analyzed"] = analyzed
    return stats
