# -*- coding: utf-8 -*-
"""
PulseDesk — Freshdesk × Outlook Classic 首次回覆 SLA 監控
=========================================================
流程：
  1. 客戶寄信到 support@ / servicedesk@  → 郵件一到即開始 15 分鐘計時
  2. Freshdesk 建立工單（has created a new Freshdesk ticket / Ticket Received）→ 自動合併到同一張工單
  3. 同事回覆客戶（Freshdesk「Re: Tkt#123456」或直接用 Outlook 回覆）→ 計時停止
  4. 在 PulseDesk 分派給同事跟進 → 跟進中 / 等待客戶 / 已解決
另外會列出你在 Outlook 手動設定了分類（Categories）的郵件，方便跟進。

本程式只「讀取」Outlook 郵件，不會刪除、移動、修改或自動寄出任何郵件。
「通知同事」只會開啟一封草稿，由你自己按傳送。

執行：python pulsedesk.py          （正式模式，需要 Windows + Outlook Classic + pywin32）
      python pulsedesk.py --demo   （示範模式，用虛擬資料預覽界面）
"""

import argparse
import hashlib
import json
import os
import re
import sqlite3
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs
from urllib.request import urlopen

try:
    import pythoncom
    import win32com.client
    HAS_COM = True
except ImportError:
    HAS_COM = False

APP = "PulseDesk"
VERSION = "2.7"
DATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "PulseDesk")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

# 同事及使用 support@ / servicedesk@ 回覆時的簽名（多行簽名用換行分隔）
DEFAULT_STAFF = [
    {"name": "Alan Wong", "email": "", "signature": "Thanks\nRegards"},
    {"name": "Alan Wong 2", "email": "", "signature": "Thanks"},
    {"name": "Steven Li", "email": "", "signature": "Million Thanks"},
    {"name": "Stevan Cheung", "email": "", "signature": "Thanks & Best Regards"},
    {"name": "Chester Choi", "email": "", "signature": "Yours Sincerely"},
    {"name": "Kelvin Wong", "email": "", "signature": "Many Thanks"},
    {"name": "Benny Chan", "email": "", "signature": "Kind Regards"},
    {"name": "Sai Ho", "email": "", "signature": "Thank you"},
    {"name": "Simon Wong", "email": "", "signature": "Yours Truly"},
    {"name": "Matthew Cheng", "email": "", "signature": "Thanks & Regards"},
    {"name": "Ricky Chiu", "email": "", "signature": "Respectfully yours"},
]

DEFAULT_CONFIG = {
    "port": 8765,
    "internal_domains": ["premier-technology.com"],
    "support_addresses": ["support@premier-technology.com", "servicedesk@premier-technology.com"],
    "freshdesk_url": "https://premiertechnology.freshdesk.com",
    "sla_minutes": 15,
    "warning_minutes": 10,
    "critical_minutes": 13,
    "assign_minutes": 30,
    "unverified_hours": 8,
    "scan_interval_seconds": 60,    # 後備掃描；新郵件主要靠 Outlook 即時事件
    "days_back": 3,
    "scan_folders": [],             # 空白 = 自己信箱的收件匣（及寄件備份）
    "include_subfolders": False,
    "scan_sent_items": True,
    "deep_scan_minutes": 120,       # 每隔多久完整覆查一次（補回離線同步的舊郵件）
    "extra_folders": [],
    "ignore_senders": ["noreply", "no-reply", "donotreply", "mailer-daemon", "postmaster"],
    "ignore_subjects": ["please ignore", "automatic reply", "auto reply", "autoreply", "out of office",
                        "自動回覆", "自动回复", "undeliverable", "delivery status notification"],
    "colleagues": [dict(c) for c in DEFAULT_STAFF],
    "sound": True,
    "desktop_popup": True,
    "category_refresh_minutes": 15, # 分類完整更新間隔；平時用 Outlook 事件即時更新
    "category_folders": [],         # 空白 = 收件匣及其子資料夾（不包括超過 5000 封的大資料夾）及工作
    "hidden_categories": [],
    "assign_queue_hours": 24,       # 回覆後超過這個時間仍未分派，移出「待分派」
    "config_version": 25,
}

TS_FMT = "%Y-%m-%dT%H:%M:%S"
MERGE_WINDOW = timedelta(hours=12)      # 同一客戶同一主旨，多久內當作同一個請求
LOOKBACK = timedelta(days=3)


def now():
    return datetime.now().replace(microsecond=0)


def iso(dt):
    return dt.strftime(TS_FMT) if dt else None


def parse_iso(s):
    return datetime.strptime(s, TS_FMT) if s else None


def to_naive(dt):
    return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


# =====================================================================
#  設定
# =====================================================================
def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    saved = {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
    except FileNotFoundError:
        pass
    except Exception:
        traceback.print_exc()
    cfg.update(saved)
    if saved and int(saved.get("config_version", 0) or 0) < 24:
        # v2.4：減輕 Outlook 負擔
        cfg["scan_interval_seconds"] = max(60, int(cfg.get("scan_interval_seconds", 60)))
        cfg["deep_scan_minutes"] = max(120, int(cfg.get("deep_scan_minutes", 120)))
        cfg["config_version"] = 24
    if saved and int(cfg.get("config_version", 0) or 0) < 25:
        # v2.5：加入同事簽名（保留原有同事及電郵）
        cols = [dict(c) for c in cfg.get("colleagues") or []]
        by_name = {c["name"].strip().lower(): c for c in cols}
        for st in DEFAULT_STAFF:
            c = by_name.get(st["name"].lower())
            if c is None:
                cols.append(dict(st))
            elif not c.get("signature"):
                c["signature"] = st["signature"]
        cfg["colleagues"] = cols
        cfg["config_version"] = 25
    for k in ("category_scan_seconds", "category_all_stores", "all_accounts"):
        cfg.pop(k, None)
    return cfg


def save_config(cfg):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_PATH)


def sanitize_config(new, old):
    cfg = dict(old)
    ints = {"sla_minutes": (1, 240), "warning_minutes": (1, 240), "critical_minutes": (1, 240),
            "assign_minutes": (1, 1440), "unverified_hours": (1, 168),
            "scan_interval_seconds": (15, 600), "days_back": (1, 30), "category_refresh_minutes": (5, 240),
            "deep_scan_minutes": (30, 1440), "assign_queue_hours": (1, 720)}
    for k, (lo, hi) in ints.items():
        if k in new:
            try:
                cfg[k] = max(lo, min(hi, int(new[k])))
            except (TypeError, ValueError):
                pass
    for k in ("include_subfolders", "scan_sent_items", "sound", "desktop_popup"):
        if k in new:
            cfg[k] = bool(new[k])
    for k in ("internal_domains", "support_addresses", "extra_folders", "scan_folders", "category_folders", "ignore_senders",
              "ignore_subjects", "hidden_categories"):
        if k in new and isinstance(new[k], list):
            vals = [str(v).strip() for v in new[k] if str(v).strip()]
            if k == "internal_domains":
                vals = [v.lower().lstrip("@") for v in vals]
            elif k in ("support_addresses", "ignore_senders", "ignore_subjects"):
                vals = [v.lower() for v in vals]
            cfg[k] = list(dict.fromkeys(vals))
    if "freshdesk_url" in new:
        cfg["freshdesk_url"] = str(new["freshdesk_url"]).strip().rstrip("/")
    if "colleagues" in new and isinstance(new["colleagues"], list):
        cols = []
        for c in new["colleagues"]:
            name = str(c.get("name", "")).strip()
            if name:
                sig = "\n".join(l.strip() for l in str(c.get("signature", "")).splitlines() if l.strip())[:200]
                cols.append({"name": name, "email": str(c.get("email", "")).strip(), "signature": sig})
        cfg["colleagues"] = cols
    if cfg["warning_minutes"] >= cfg["sla_minutes"]:
        cfg["warning_minutes"] = max(1, cfg["sla_minutes"] - 5)
    if not (cfg["warning_minutes"] <= cfg["critical_minutes"] < cfg["sla_minutes"]):
        cfg["critical_minutes"] = max(cfg["warning_minutes"], cfg["sla_minutes"] - 2)
    return cfg


# =====================================================================
#  郵件解析
# =====================================================================
TKT_RE = re.compile(r"Tkt\s*#\s*(\d{3,})", re.I)
REF_RE = re.compile(r"\[#(\d{4,})\]")
LINK_RE = re.compile(r"/helpdesk/tickets/(\d{3,})", re.I)
NEW_RE = re.compile(
    r"^\s*(?:(?P<company>.*?)\s+-\s+)?(?P<requester>[^-]+?)\s+has created a new Freshdesk ticket\s*:?\s*(?P<subject>.*)$",
    re.I)
ACK_RE = re.compile(r"^\s*Ticket Received\s*[-:–]\s*(?P<subject>.*)$", re.I)
PREFIX_RE = re.compile(r"^\s*(?:(?:re|fw|fwd|aw|sv|回覆|回复|答覆|轉寄|转发)\s*[:：]\s*|\[(?:external|ext|外部)\]\s*)+", re.I)
COMMENT_RE = re.compile(r"^\s*new comment\s*-\s*", re.I)
DEAR_RE = re.compile(r"Dear\s+(.{1,60}?)\s*[,，\r\n]")


def clean_subject(s):
    s = s or ""
    for _ in range(5):
        s = PREFIX_RE.sub("", s)
        s = COMMENT_RE.sub("", s)
        s = TKT_RE.sub("", s, count=1).strip()
        s = REF_RE.sub("", s, count=1).strip()
    return re.sub(r"\s+", " ", s).strip()


def norm_subject(s):
    return clean_subject(s).lower()


def ticket_ref(subject):
    m = TKT_RE.search(subject or "") or REF_RE.search(subject or "")
    return m.group(1) if m else None


def domain_of(addr):
    return addr.rsplit("@", 1)[1].lower() if addr and "@" in addr else ""


def is_internal_addr(addr, cfg):
    if not addr:
        return False
    if addr.startswith("/o="):
        return True
    return domain_of(addr) in cfg["internal_domains"]


def is_external_addr(addr, cfg):
    return bool(addr) and "@" in addr and not is_internal_addr(addr, cfg)


def ignore_rule(sender, subject, cfg):
    """回傳符合的忽略規則（沒有則回傳 None）。"""
    s = (sender or "").lower()
    for x in cfg.get("ignore_senders", []):
        if x and x in s:
            return f"忽略寄件人「{x}」"
    subj = (subject or "").lower()
    for x in cfg.get("ignore_subjects", []):
        if x and x in subj:
            return f"忽略主旨「{x}」"
    return None


def is_ignored(sender, subject, cfg):
    return ignore_rule(sender, subject, cfg) is not None


# ---------------------------------------------------------------------
#  辨認回覆的同事
#  - 用 support@ / servicedesk@ 寄出：按簽名辨認（只看新回覆部分，不看下面引用的舊郵件）
#  - 用自己的信箱寄出：按寄件人辨認（因為不同同事的簽名可能相同）
# ---------------------------------------------------------------------
QUOTE_HEAD_RE = re.compile(r"^\s*(?:(?:from|寄件者|發件人|发件人|差出人)\s*[:：]|-{2,}\s*original message|-{2,}\s*原始郵件|_{8,}|>)", re.I)
WROTE_RE = re.compile(r"(?:\bwrote|寫道|写道)\s*[:：]?\s*$", re.I)
SIG_STRIP = " ,.!;:，。！；：-–—~*_\t"


def norm_sig_line(s):
    s = (s or "").replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "").strip().lower()
    s = re.sub(r"\s+", " ", s).replace(" and ", " & ")
    return s.strip(SIG_STRIP)


def reply_part(body):
    """郵件內文中「新回覆」的部分：遇到引用的舊郵件（From: / On ... wrote: / -----Original Message----- 等）就停止。"""
    lines = (body or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out = []
    for i, line in enumerate(lines):
        if QUOTE_HEAD_RE.match(line):
            break
        if WROTE_RE.search(line):
            if out and norm_sig_line(out[-1]).startswith("on "):
                out.pop()
            break
        out.append(line)
    return out


def staff_signatures(cfg):
    res = []
    for c in cfg.get("colleagues") or []:
        sig = [norm_sig_line(l) for l in str(c.get("signature") or "").splitlines()]
        sig = [l for l in sig if l]
        if sig:
            res.append((c["name"], sig))
    return res


def match_signature(body, cfg):
    """回傳 (同事名稱列表, 符合的簽名, 新回覆的最後一行)。
    簽名必須獨立成行並完全相同（例如「Thanks」不會誤認「Thanks & Regards」）；
    多行簽名（例如 Thanks / Regards）優先於較短的簽名；取最接近回覆結尾的簽名。"""
    lines = [norm_sig_line(l) for l in reply_part(body)]
    raw = [l.strip() for l in reply_part(body)]
    idx = [i for i, l in enumerate(lines) if l]
    L = [lines[i] for i in idx]
    best, names, sig_text = None, [], ""
    for name, sig in staff_signatures(cfg):
        k = len(sig)
        for i in range(len(L) - k + 1):
            if L[i:i + k] == sig:
                key = (i + k - 1, k)
                if best is None or key > best:
                    best, names, sig_text = key, [name], " / ".join(raw[idx[j]] for j in range(i, i + k))
                elif key == best and name not in names:
                    names.append(name)
    tail = next((raw[i] for i in reversed(idx)), "")[:60]
    return names, sig_text, tail


def clean_sender_name(name):
    n = (name or "").split("/")[0].strip().strip('"').strip()
    return n


def colleague_by_sender(sender, sender_name, cfg):
    s = (sender or "").lower()
    n = clean_sender_name(sender_name).lower()
    for c in cfg.get("colleagues") or []:
        if c.get("email") and c["email"].strip().lower() == s:
            return c["name"]
    for c in cfg.get("colleagues") or []:
        if n and c["name"].strip().lower() == n:
            return c["name"]
    return None


def identify_replier(m, cfg):
    """回傳 (同事名稱, 判斷依據)。"""
    sender = (m.sender or "").lower()
    shared = sender in set(cfg.get("support_addresses") or [])
    if not shared:
        name = colleague_by_sender(sender, m.sender_name, cfg)
        if name:
            return name, f"寄件人 {m.sender_name or sender}"
        clean = clean_sender_name(m.sender_name) or sender
        return clean, "寄件人（不在同事名單）"
    name = colleague_by_sender("", m.sender_name, cfg)   # 共用信箱但寄件人名稱是同事名稱
    if name:
        return name, f"寄件人名稱 {m.sender_name}"
    try:
        names, sig, tail = match_signature(m.body(), cfg)
    except Exception:
        return "", "未能讀取內文"
    if len(names) == 1:
        return names[0], f"簽名「{sig}」"
    if names:
        return " / ".join(names), f"簽名「{sig}」有多位同事相同，未能確定"
    return "", "未能辨認簽名" + (f"（最後一行：{tail}）" if tail else "")


def classify_explain(m, cfg):
    """把一封郵件分類為工單事件，並說明原因。回傳 (事件 或 None, 原因)。
    m 需提供：subject, sender, sender_name, sender_internal, time, entry_id, store_id, recipients(), body()"""
    subj = m.subject or ""

    def ev(kind, tid, **kw):
        d = {"kind": kind, "ticket_id": tid, "time": m.time, "entry_id": m.entry_id,
             "store_id": m.store_id, "subject": subj, "sender": m.sender, "sender_name": m.sender_name,
             "company": "", "requester": "", "requester_email": "", "externals": []}
        d.update(kw)
        return d

    def first(rx, text):
        mm = rx.search(text or "")
        return mm.group(1) if mm else None

    nm = NEW_RE.match(subj)
    if nm:
        link = first(LINK_RE, m.body())
        tid = link or ticket_ref(subj)
        if not tid:
            return None, "Freshdesk 建立工單通知，但找不到工單號碼"
        return (ev("new_ticket", tid, subject=clean_subject(nm.group("subject")),
                   company=(nm.group("company") or "").strip(), requester=nm.group("requester").strip()),
                "主旨是「has created a new Freshdesk ticket」" + ("，工單號碼來自郵件連結" if link else "，工單號碼來自主旨"))

    am = ACK_RE.match(subj)
    if am and m.sender_internal:
        body = m.body()
        tid = first(LINK_RE, body) or ticket_ref(subj)
        if not tid:
            return None, "Ticket Received 郵件，但找不到工單號碼"
        ext = [a for t, a in m.recipients() if t == 1 and is_external_addr(a, cfg)]
        dear = DEAR_RE.search(body or "")
        return (ev("ack", tid, subject=clean_subject(am.group("subject")),
                   requester=dear.group(1).strip() if dear else "",
                   requester_email=ext[0] if ext else "", externals=ext),
                "Freshdesk 自動回覆「Ticket Received」（不算回覆客戶）")

    tid = ticket_ref(subj)
    if m.sender_internal:
        ext = [a for t, a in m.recipients() if t == 1 and is_external_addr(a, cfg)]
        if not ext:
            return None, "內部寄出，但收件人（To）沒有外部客戶"
        who, how = identify_replier(m, cfg)
        whotxt = f"；回覆同事：{who or '未知'}（{how}）"
        if tid:
            return (ev("agent_reply", tid, subject=clean_subject(subj), requester_email=ext[0], externals=ext,
                       replier=who, replier_how=how),
                    f"內部寄給客戶 {ext[0]}，主旨有工單號碼 #{tid}" + whotxt)
        return (ev("direct_reply", None, subject=clean_subject(subj), requester_email=ext[0], externals=ext,
                   replier=who, replier_how=how),
                f"內部寄給客戶 {ext[0]}，沒有工單號碼，按主旨配對" + whotxt)

    rule = ignore_rule(m.sender, subj, cfg)
    if rule:
        return None, rule
    if tid:
        return (ev("customer_reply", tid, subject=clean_subject(subj), requester_email=m.sender,
                   requester=m.sender_name), f"外部來信，主旨有工單號碼 #{tid}（客戶再回覆）")
    support = set(cfg["support_addresses"])
    hit = [a for _, a in m.recipients() if a in support]
    if hit:
        return (ev("customer_new", None, subject=clean_subject(subj), requester_email=m.sender,
                   requester=m.sender_name), f"外部寄件人，收件人包括 {hit[0]}")
    return None, "外部郵件，但收件人沒有 support / servicedesk"


def classify(m, cfg):
    return classify_explain(m, cfg)[0]


# =====================================================================
#  Outlook 郵件讀取
# =====================================================================
PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
EX_CACHE = {}   # Exchange 內部地址 → SMTP，避免重複查詢通訊錄（查詢通訊錄會令 Outlook 變慢）


class OutlookMail:
    def __init__(self, item, store_id, cfg, sent=False):
        self.item, self.store_id, self.cfg = item, store_id, cfg
        self.entry_id = item.EntryID
        self.subject = item.Subject or ""
        self.time = to_naive(item.SentOn if sent else item.ReceivedTime)
        self.sender_name = item.SenderName or ""
        self._sender = None
        self._ex = False
        self._recips = None
        self._body = None

    @property
    def sender(self):
        if self._sender is None:
            addr = ""
            try:
                if self.item.SenderEmailType == "EX":
                    self._ex = True
                    dn = ""
                    try:
                        dn = (self.item.SenderEmailAddress or "").lower()
                    except Exception:
                        pass
                    addr = EX_CACHE.get(dn, "") if dn else ""
                    if not addr:
                        try:
                            addr = self.item.PropertyAccessor.GetProperty(PR_SENDER_SMTP) or ""
                        except Exception:
                            pass
                    if not addr:
                        try:
                            ex = self.item.Sender.GetExchangeUser()
                            addr = ex.PrimarySmtpAddress if ex else ""
                        except Exception:
                            pass
                    if dn and addr:
                        EX_CACHE[dn] = addr
                else:
                    addr = self.item.SenderEmailAddress or ""
            except Exception:
                pass
            self._sender = (addr or "").lower()
        return self._sender

    @property
    def sender_internal(self):
        addr = self.sender
        return self._ex or is_internal_addr(addr, self.cfg)

    def recipients(self):
        if self._recips is None:
            out = []
            try:
                for r in self.item.Recipients:
                    addr = r.Address or ""
                    if "@" in addr and not addr.startswith("/"):
                        out.append((r.Type, addr.lower()))   # 一般 SMTP 地址，不用再查
                        continue
                    dn = addr.lower()
                    addr = EX_CACHE.get(dn, "") if dn else ""
                    if not addr:
                        try:
                            addr = r.PropertyAccessor.GetProperty(PR_SMTP) or ""
                        except Exception:
                            pass
                    if not addr:
                        try:
                            ex = r.AddressEntry.GetExchangeUser()
                            addr = ex.PrimarySmtpAddress if ex else ""
                        except Exception:
                            pass
                    if dn and addr:
                        EX_CACHE[dn] = addr
                    if not addr:
                        addr = dn
                    out.append((r.Type, addr.lower()))
            except Exception:
                pass
            self._recips = out
        return self._recips

    def body(self):
        if self._body is None:
            try:
                self._body = (self.item.Body or "")[:8000]
            except Exception:
                self._body = ""
        return self._body


# =====================================================================
#  資料庫
# =====================================================================
SCHEMA = """
CREATE TABLE IF NOT EXISTS tickets(
  id TEXT PRIMARY KEY, subject TEXT DEFAULT '', company TEXT DEFAULT '', requester TEXT DEFAULT '',
  requester_email TEXT DEFAULT '', sla_start TEXT, first_response_at TEXT, response_source TEXT,
  last_agent_reply_at TEXT, last_customer_reply_at TEXT,
  assignee TEXT, assigned_at TEXT, status TEXT DEFAULT 'open', status_at TEXT, note TEXT DEFAULT '',
  notif_entry TEXT, ack_entry TEXT, reply_entry TEXT, customer_entry TEXT, store_id TEXT,
  ignored INTEGER DEFAULT 0, seen_at TEXT, updated_at TEXT, origin TEXT DEFAULT 'freshdesk',
  first_responder TEXT, first_responder_how TEXT, last_responder TEXT);
CREATE TABLE IF NOT EXISTS mails(
  entry_id TEXT PRIMARY KEY, kind TEXT, ticket_id TEXT, time TEXT, subject TEXT, sender TEXT,
  sender_name TEXT, store_id TEXT, matched TEXT, dismissed INTEGER DEFAULT 0, replier TEXT, replier_how TEXT);
CREATE INDEX IF NOT EXISTS mails_ticket ON mails(ticket_id);
CREATE INDEX IF NOT EXISTS mails_kind ON mails(kind, time);
CREATE TABLE IF NOT EXISTS activity(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ticket_id TEXT, time TEXT, text TEXT);
CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS processed(entry_id TEXT PRIMARY KEY, time TEXT);
CREATE TABLE IF NOT EXISTS scanlog(
  entry_id TEXT PRIMARY KEY, time TEXT, folder TEXT, sender TEXT, sender_name TEXT, internal INTEGER,
  recipients TEXT, subject TEXT, kind TEXT, reason TEXT, result TEXT, logged_at TEXT, replier TEXT, replier_how TEXT);
CREATE INDEX IF NOT EXISTS scanlog_time ON scanlog(time);
CREATE TABLE IF NOT EXISTS corrections(
  id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, ticket_id TEXT, action TEXT, detail TEXT,
  subject TEXT, sender TEXT);
"""


def placeholder_id(entry_id):
    return "E" + hashlib.sha1(entry_id.encode("utf-8", "ignore")).hexdigest()[:10].upper()


class DB:
    def __init__(self, path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.lock = threading.RLock()
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        with self.lock:
            self.conn.executescript(SCHEMA)
            for table, col, typ in (("tickets", "origin", "TEXT DEFAULT 'freshdesk'"), ("tickets", "first_responder", "TEXT"),
                                    ("tickets", "first_responder_how", "TEXT"), ("tickets", "last_responder", "TEXT"),
                                    ("mails", "replier", "TEXT"), ("mails", "replier_how", "TEXT"),
                                    ("scanlog", "replier", "TEXT"), ("scanlog", "replier_how", "TEXT")):
                cols = {r[1] for r in self.conn.execute(f"PRAGMA table_info({table})")}
                if col not in cols:
                    self.conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
            self.conn.commit()
        if not self.meta("baseline_at"):
            # 由這個版本開始監控的時間；之前的舊郵件不會突然全部變成「已超時」
            self.set_meta("baseline_at", iso(now()))
        if not self.meta("mig24"):
            # 之前「全部標記為已回覆」的舊郵件：回覆時間不明，不計入達標率
            base = self.meta("baseline_at")
            self.x("UPDATE tickets SET response_source='manual_unknown' WHERE response_source='manual' AND ("
                   "(julianday(first_response_at)-julianday(sla_start))*24 > 8 OR "
                   "(origin='email' AND julianday(sla_start) < julianday(?) - 30.0/1440))", (base,))
            self.set_meta("mig24", iso(now()))

    # ---------- 已讀取郵件紀錄（重新開啟程式時不用再完整掃描） ----------
    def load_processed(self, days):
        lo = iso(now() - timedelta(days=days + 2))
        self.x("DELETE FROM processed WHERE time<?", (lo,))
        ids = {r["entry_id"] for r in self.q("SELECT entry_id FROM processed")}
        ids |= {r["entry_id"] for r in self.q("SELECT entry_id FROM mails WHERE time>=? AND "
                                             "(ticket_id IS NOT NULL OR kind NOT IN ('customer_new'))", (lo,))}
        return ids

    def mark_processed(self, rows):
        if not rows:
            return
        with self.lock:
            self.conn.executemany("INSERT OR IGNORE INTO processed(entry_id,time) VALUES(?,?)", rows)
            self.conn.commit()

    def log_scan(self, row):
        with self.lock:
            self.conn.execute("INSERT OR REPLACE INTO scanlog(entry_id,time,folder,sender,sender_name,internal,recipients,"
                              "subject,kind,reason,result,logged_at,replier,replier_how) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                              (row["entry_id"], row["time"], row["folder"], row["sender"], row["sender_name"],
                               row["internal"], row["recipients"], row["subject"], row["kind"], row["reason"],
                               row["result"], iso(now()), row.get("replier"), row.get("replier_how")))
            self.conn.commit()

    def prune_scanlog(self, days=14):
        self.x("DELETE FROM scanlog WHERE time<?", (iso(now() - timedelta(days=days)),))

    def add_correction(self, tk, action, detail=""):
        self.x("INSERT INTO corrections(time,ticket_id,action,detail,subject,sender) VALUES(?,?,?,?,?,?)",
               (iso(now()), tk["id"], action, detail, tk.get("subject") or "", tk.get("requester_email") or ""))

    def clear_scan_state(self):
        self.x("DELETE FROM processed")
        self.x("DELETE FROM meta WHERE k LIKE 'wm:%'")

    def meta(self, k):
        r = self.q("SELECT v FROM meta WHERE k=?", (k,))
        return r[0]["v"] if r else None

    def set_meta(self, k, v):
        self.x("INSERT OR REPLACE INTO meta(k,v) VALUES(?,?)", (k, v))

    def q(self, sql, args=()):
        with self.lock:
            return [dict(r) for r in self.conn.execute(sql, args).fetchall()]

    def x(self, sql, args=()):
        with self.lock:
            cur = self.conn.execute(sql, args)
            self.conn.commit()
            return cur.rowcount

    def get(self, tid):
        rows = self.q("SELECT * FROM tickets WHERE id=?", (tid,))
        return rows[0] if rows else None

    def log(self, tid, text, t=None):
        self.x("INSERT INTO activity(ticket_id,time,text) VALUES(?,?,?)", (tid, iso(t or now()), text))

    def update(self, tid, **fields):
        fields["updated_at"] = iso(now())
        cols = ", ".join(f"{k}=?" for k in fields)
        self.x(f"UPDATE tickets SET {cols} WHERE id=?", (*fields.values(), tid))

    def record_mail(self, ev, tid):
        self.x("INSERT OR IGNORE INTO mails(entry_id,kind,ticket_id,time,subject,sender,sender_name,store_id,replier,replier_how)"
               " VALUES(?,?,?,?,?,?,?,?,?,?)",
               (ev["entry_id"], ev["kind"], tid, iso(ev["time"]), ev["subject"], ev["sender"],
                ev["sender_name"], ev["store_id"], ev.get("replier"), ev.get("replier_how")))
        self.x("UPDATE mails SET ticket_id=? WHERE entry_id=? AND ticket_id IS NULL", (tid, ev["entry_id"]))

    def mail_done(self, entry_id):
        r = self.q("SELECT ticket_id FROM mails WHERE entry_id=?", (entry_id,))
        return bool(r and r[0]["ticket_id"])

    # ---------- 尋找相關工單 ----------
    def find_related(self, ns, emails, t, placeholders_only=False):
        """同一主旨（及同一客戶）而且時間接近的工單。"""
        if not ns:
            return None
        lo, hi = iso(t - LOOKBACK), iso(t + LOOKBACK)
        rows = self.q("SELECT * FROM tickets WHERE ignored=0 AND (sla_start IS NULL OR (sla_start>=? AND sla_start<=?))"
                      + (" AND id LIKE 'E%'" if placeholders_only else ""), (lo, hi))
        emails = {e for e in emails if e}
        best, best_gap = None, None
        for r in rows:
            if norm_subject(r["subject"]) != ns:
                continue
            if emails and r["requester_email"] and r["requester_email"] not in emails:
                continue
            st = parse_iso(r["sla_start"])
            gap = abs((st - t).total_seconds()) if st else 10 ** 9
            if best is None or gap < best_gap:
                best, best_gap = r, gap
        return best

    def merge_into(self, src_id, dst_id, label=None):
        """把臨時郵件工單（E…）合併到 Freshdesk 工單，或把兩張工單合併。"""
        if src_id == dst_id:
            return
        src, dst = self.get(src_id), self.get(dst_id)
        if not src:
            return
        with self.lock:
            if not dst:
                self.x("UPDATE tickets SET id=?, origin='freshdesk', updated_at=? WHERE id=?", (dst_id, iso(now()), src_id))
            else:
                upd = {}
                s1, s2 = parse_iso(src["sla_start"]), parse_iso(dst["sla_start"])
                if s1 and (not s2 or s1 < s2):
                    upd["sla_start"] = src["sla_start"]
                for k in ("customer_entry", "requester_email", "requester", "subject", "store_id"):
                    if src.get(k) and not dst.get(k):
                        upd[k] = src[k]
                f1, f2 = parse_iso(src["first_response_at"]), parse_iso(dst["first_response_at"])
                if f1 and (not f2 or f1 < f2):
                    upd.update(first_response_at=src["first_response_at"], response_source=src["response_source"],
                               reply_entry=src["reply_entry"], first_responder=src.get("first_responder"),
                               first_responder_how=src.get("first_responder_how"))
                for k in ("last_customer_reply_at", "last_agent_reply_at"):
                    a, b = parse_iso(src[k]), parse_iso(dst[k])
                    if a and (not b or a > b):
                        upd[k] = src[k]
                if src["assignee"] and not dst["assignee"]:
                    upd.update(assignee=src["assignee"], assigned_at=src["assigned_at"], status=src["status"],
                               status_at=src["status_at"])
                if src["note"] and not dst["note"]:
                    upd["note"] = src["note"]
                if upd:
                    self.update(dst_id, **upd)
                self.x("DELETE FROM tickets WHERE id=?", (src_id,))
            self.x("UPDATE mails SET ticket_id=? WHERE ticket_id=?", (dst_id, src_id))
            self.x("UPDATE activity SET ticket_id=? WHERE ticket_id=?", (dst_id, src_id))
        self.log(dst_id, label or f"客戶郵件已對應 Freshdesk 工單 Tkt#{dst_id}")

    def absorb(self, tid, ns, emails, t):
        ph = self.find_related(ns, emails, t, placeholders_only=True)
        if ph:
            self.merge_into(ph["id"], tid)

    def ensure(self, tid, ev):
        if not self.get(tid):
            self.x("INSERT INTO tickets(id, subject, seen_at, updated_at, store_id, origin) VALUES(?,?,?,?,?,?)",
                   (tid, ev.get("subject", ""), iso(now()), iso(now()), ev.get("store_id"), "freshdesk"))
            return True
        return False

    def set_first_response(self, row, t, entry, source="email", label="已首次回覆客戶", replier=None, how=None):
        start = parse_iso(row["sla_start"])
        fr = parse_iso(row["first_response_at"])
        upd = {}
        if (start is None or t >= start - timedelta(minutes=2)) and (fr is None or t < fr):
            upd.update(first_response_at=iso(t), response_source=source, reply_entry=entry,
                       first_responder=replier or None, first_responder_how=how)
            if fr is None:
                self.log(row["id"], label + (f"：{replier}" if replier else ""), t)
        last = parse_iso(row["last_agent_reply_at"])
        if last is None or t > last:
            upd["last_agent_reply_at"] = iso(t)
            if replier:
                upd["last_responder"] = replier
        if upd:
            self.update(row["id"], **upd)

    # ---------- 套用郵件事件 ----------
    def apply(self, ev):
        """套用事件，回傳處理結果的說明（用於診斷資料）。"""
        with self.lock:
            if self.mail_done(ev["entry_id"]):
                r = self.q("SELECT ticket_id FROM mails WHERE entry_id=?", (ev["entry_id"],))
                return "之前已處理（工單 %s）" % (r[0]["ticket_id"] if r else "")
            kind, tid, t = ev["kind"], ev["ticket_id"], ev["time"]
            ns = norm_subject(ev["subject"])
            emails = ev.get("externals") or ([ev["requester_email"]] if ev.get("requester_email") else [])

            if kind in ("new_ticket", "ack"):
                self.absorb(tid, ns, emails, t)
                created = self.ensure(tid, ev)
                row = self.get(tid)
                upd = {"origin": "freshdesk"}
                for k in ("subject", "company", "requester", "requester_email"):
                    if ev.get(k) and (not row[k] or (k == "subject" and kind == "new_ticket")):
                        upd[k] = ev[k]
                start = parse_iso(row["sla_start"])
                if start is None or t < start:
                    upd["sla_start"] = iso(t)
                upd["notif_entry" if kind == "new_ticket" else "ack_entry"] = ev["entry_id"]
                upd["store_id"] = row["store_id"] or ev["store_id"]
                self.update(tid, **upd)
                if created:
                    self.log(tid, "Freshdesk 建立工單", t)
                self.record_mail(ev, tid)
                return f"建立 / 更新工單 #{tid}"

            elif kind == "agent_reply":
                self.absorb(tid, ns, emails, t)
                self.ensure(tid, ev)
                row = self.get(tid)
                if not row["requester_email"] and ev.get("requester_email"):
                    self.update(tid, requester_email=ev["requester_email"])
                    row = self.get(tid)
                had = row["first_response_at"]
                self.set_first_response(row, t, ev["entry_id"], "email", "已在 Freshdesk 回覆客戶（Re: Tkt# 郵件）",
                                        ev.get("replier"), ev.get("replier_how"))
                self.record_mail(ev, tid)
                return f"#{tid} 已有首次回覆" if had and parse_iso(had) <= t else f"#{tid} 停止計時（首次回覆）"

            elif kind == "direct_reply":
                row = self.find_related(ns, emails, t)
                if row and (not row["sla_start"] or parse_iso(row["sla_start"]) <= t + timedelta(minutes=2)):
                    had = row["first_response_at"]
                    self.set_first_response(row, t, ev["entry_id"], "email", "同事直接用 Outlook 回覆客戶",
                                            ev.get("replier"), ev.get("replier_how"))
                    self.record_mail(ev, row["id"])
                    return f"配對到 {row['id']}" + ("（已有首次回覆）" if had else "，停止計時")
                return "找不到相同主旨及客戶的請求，不影響計時"

            elif kind == "customer_reply":
                row = self.get(tid)
                if row:
                    last = parse_iso(row["last_customer_reply_at"])
                    if last is None or t > last:
                        self.update(tid, last_customer_reply_at=iso(t))
                self.record_mail(ev, tid)
                return f"#{tid} 標示客戶新回覆" if row else f"#{tid} 不在 PulseDesk 紀錄，列入「客戶回覆了舊工單」"

            elif kind == "customer_new":
                row = self.find_related(ns, [ev["requester_email"]], t)
                if row:
                    st, fr = parse_iso(row["sla_start"]), parse_iso(row["first_response_at"])
                    is_origin = (st is None or t < st) and (st is None or st - t <= MERGE_WINDOW) \
                        and (fr is None or t <= fr)
                    if is_origin:
                        upd = {"sla_start": iso(t), "customer_entry": ev["entry_id"]}
                        if st and row["origin"] == "email":
                            # 之前當作開始的郵件其實是客戶再次來信
                            last = parse_iso(row["last_customer_reply_at"])
                            if last is None or st > last:
                                upd["last_customer_reply_at"] = iso(st)
                        if not row["requester_email"]:
                            upd["requester_email"] = ev["requester_email"]
                        if not row["requester"]:
                            upd["requester"] = ev["requester"]
                        self.update(row["id"], **upd)
                        self.log(row["id"], "收到客戶郵件，開始計時", t)
                        note = f"合併到 {row['id']}，計時由這封郵件開始"
                    elif st and t > st and (t - st) <= MERGE_WINDOW + LOOKBACK:
                        last = parse_iso(row["last_customer_reply_at"])
                        if last is None or t > last:
                            self.update(row["id"], last_customer_reply_at=iso(t))
                        note = f"同一請求再來信（{row['id']}），不重新計時"
                    else:
                        note = f"相同主旨的請求 {row['id']}"
                    self.record_mail(ev, row["id"])
                    return note
                else:
                    pid = placeholder_id(ev["entry_id"])
                    if not self.get(pid):
                        self.x("INSERT INTO tickets(id,subject,requester,requester_email,sla_start,customer_entry,"
                               "store_id,seen_at,updated_at,origin) VALUES(?,?,?,?,?,?,?,?,?,?)",
                               (pid, ev["subject"], ev["requester"], ev["requester_email"], iso(t), ev["entry_id"],
                                ev["store_id"], iso(now()), iso(now()), "email"))
                        self.log(pid, "收到客戶郵件，開始計時", t)
                    self.record_mail(ev, pid)
                    return f"新請求 {pid}，開始計時"

    # ---------- 讀取 ----------
    def tickets(self):
        cutoff = iso(now() - timedelta(days=45))
        return self.q("SELECT * FROM tickets WHERE updated_at>=? OR sla_start>=? OR "
                      "(status!='resolved' AND ignored=0) ORDER BY COALESCE(sla_start, seen_at) DESC LIMIT 2000",
                      (cutoff, cutoff))

    def recent_mails(self, n=40):
        return self.q("SELECT entry_id, kind, ticket_id, time, subject, sender, sender_name, store_id, replier, replier_how "
                      "FROM mails ORDER BY time DESC LIMIT ?", (n,))

    def followups(self):
        """客戶回覆了 PulseDesk 未有紀錄的舊工單（例如「New comment - [#254299]」）。"""
        lo = iso(now() - timedelta(days=3))
        return self.q("SELECT m.* FROM mails m LEFT JOIN tickets t ON t.id=m.ticket_id WHERE m.kind='customer_reply' "
                      "AND m.dismissed=0 AND m.time>=? AND t.id IS NULL ORDER BY m.time DESC LIMIT 60", (lo,))

    def timeline(self, tid):
        mails = self.q("SELECT entry_id, kind, time, subject, sender, sender_name, store_id, replier, replier_how FROM mails "
                       "WHERE ticket_id=? OR matched=? ORDER BY time", (tid, tid))
        acts = self.q("SELECT time, text FROM activity WHERE ticket_id=? ORDER BY time, id", (tid,))
        return {"mails": mails, "activity": acts}


# =====================================================================
#  Outlook 掃描器（即時事件 + 定時掃描）
# =====================================================================
def get_folder_by_path(ns, path):
    parts = [p for p in path.strip().strip("\\").split("\\") if p]
    folder = ns.Folders.Item(parts[0])
    for p in parts[1:]:
        folder = folder.Folders.Item(p)
    return folder


def iter_folders(folder, include_sub):
    yield folder
    if include_sub:
        try:
            for sub in folder.Folders:
                yield from iter_folders(sub, include_sub)
        except Exception:
            pass


class ItemsEvents:
    """Outlook 新郵件事件：郵件一到資料夾就立即處理。"""
    scanner = None
    sent = False
    path = ""

    def OnItemAdd(self, item):
        sc = ItemsEvents.scanner
        if sc is not None:
            sc.pending.append((item, self.sent, self.path))


class CatEvents:
    """分類資料夾的變更事件：你在 Outlook 設定 / 取消分類時即時更新。"""
    scanner = None

    def OnItemAdd(self, item):
        sc = CatEvents.scanner
        if sc is not None:
            sc.cat_pending.append(item)

    def OnItemChange(self, item):
        sc = CatEvents.scanner
        if sc is not None:
            sc.cat_pending.append(item)

    def OnItemRemove(self):
        sc = CatEvents.scanner
        if sc is not None and not sc.cat_dirty_at:
            sc.cat_dirty_at = time.time() + 120   # 郵件被移走或刪除：兩分鐘後重新整理一次


OUTLOOK_COLORS = {
    0: "#9ca3af", 1: "#e5484d", 2: "#f5871f", 3: "#f2b98c", 4: "#f2c200", 5: "#3cb371", 6: "#20b2aa",
    7: "#9aa84b", 8: "#3b82f6", 9: "#8b5cf6", 10: "#a0413c", 11: "#7b9bbd", 12: "#4a6484", 13: "#a3a3a3",
    14: "#6b6b6b", 15: "#333333", 16: "#a4262c", 17: "#c75b12", 18: "#b76e41", 19: "#b89500",
    20: "#1e7d32", 21: "#0e7c6b", 22: "#5d6b2b", 23: "#1e40af", 24: "#5b21b6", 25: "#6b1d1d",
}

PR_CONTENT_COUNT = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
PR_FLAG_STATUS = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
PID_TASK_DUE = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040"
PID_TASK_COMPLETE = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/811C000B"
CAT_FILTER = '@SQL=NOT("urn:schemas-microsoft-com:office:office#Keywords" IS NULL)'
BIG_FOLDER = 5000          # 預設分類資料夾不包括超過這個數量的大資料夾（例如 Alerts）


def folder_count(folder):
    try:
        return int(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
    except Exception:
        try:
            return int(folder.Items.Count)
        except Exception:
            return -1


def _is_dt(v):
    return hasattr(v, "year") and hasattr(v, "hour")


def normalize_rows(arr, ncols):
    """Table.GetArray 可能回傳「行優先」或「列優先」的二維陣列，統一轉成每行一個 tuple。
    第一欄必須是 EntryID（文字），第二欄必須是日期。"""
    if not arr:
        return []
    first = arr[0]
    if len(first) == ncols and (len(arr) != ncols or ncols < 2 or _is_dt(first[1])):
        return [tuple(r) for r in arr]
    return [tuple(arr[c][i] for c in range(ncols)) for i in range(len(first))]


def read_table(table, ncols, batch=50, on_batch=None):
    """分批讀取 Outlook Table（不需要逐封打開郵件）。"""
    use_array = True
    while not table.EndOfTable:
        rows = []
        if use_array:
            try:
                rows = normalize_rows(table.GetArray(batch), ncols)
            except Exception:
                use_array = False
        if not use_array:
            for _ in range(batch):
                if table.EndOfTable:
                    break
                rows.append(tuple(table.GetNextRow().GetValues()))
        if not rows:
            break
        for r in rows:
            yield r
        if on_batch:
            on_batch()


def open_table(folder, columns, sort_col=None, filt=""):
    """建立只包含指定欄位的 Table；回傳 (table, 成功加入的欄位名稱)。"""
    table = folder.GetTable(filt, 0)
    cols = table.Columns
    cols.RemoveAll()
    names = []
    for i, c in enumerate(columns):
        try:
            cols.Add(c)
            names.append(c)
        except Exception:
            if i < 2:
                raise          # EntryID 及日期欄一定要有
    if sort_col:
        table.Sort("[%s]" % sort_col, True)
    return table, names


def is_mail_class(mclass):
    mc = (mclass or "").upper()
    return not mc or mc.startswith("IPM.NOTE")


# ---------------------------------------------------------------------
#  Outlook 錯誤碼 → 看得懂的說明
# ---------------------------------------------------------------------
COM_FULL = {
    0x80010001: "Outlook 正忙（可能有對話框開著），稍後會自動重試",
    0x8001010A: "Outlook 正忙（可能有對話框開著），稍後會自動重試",
    0x800706BA: "Outlook 已關閉或正在重新啟動，稍後會自動重新連接",
    0x800706BE: "Outlook 已關閉或正在重新啟動，稍後會自動重新連接",
    0x80010108: "與 Outlook 的連線中斷，稍後會自動重新連接",
    0x800401E3: "Outlook 未開啟",
    0x80080005: "無法連接 Outlook（Outlook 與 PulseDesk 其中一個以系統管理員身分執行？請用相同方式打開兩者）",
}
COM_LOW = {
    0x0115: "Outlook 未能連接 Exchange 伺服器（網絡問題）。通常是共用信箱或沒有在本機快取的資料夾，稍後會自動重試",
    0x010F: "找不到郵件或資料夾（可能已被移動或刪除）",
    0x0305: "Outlook 未能讀取這個資料夾（可能沒有權限）",
}
SESSION_ERRORS = {0x800706BA, 0x800706BE, 0x80010108, 0x800401E3, 0x80080005}


def com_codes(e):
    codes = []
    args = getattr(e, "args", ()) or ()
    if args and isinstance(args[0], int):
        codes.append(args[0] & 0xFFFFFFFF)
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 5 and isinstance(args[2][5], int):
        codes.append(args[2][5] & 0xFFFFFFFF)
    return codes


def friendly_error(e):
    codes = com_codes(e)
    desc = ""
    args = getattr(e, "args", ()) or ()
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 2 and args[2][2]:
        desc = str(args[2][2]).strip()
    for c in reversed(codes):          # 內層錯誤碼較準確
        hint = COM_FULL.get(c) or ((c & 0x80000000) and COM_LOW.get(c & 0xFFFF))
        if hint:
            return f"{hint}（錯誤碼 0x{c:08X}）"
    if codes:
        return (desc or "Outlook 錯誤") + f"（錯誤碼 0x{codes[-1]:08X}）"
    return f"{type(e).__name__}: {e}"


def is_session_error(e):
    return any(c in SESSION_ERRORS for c in com_codes(e))


class Scanner(threading.Thread):
    """在獨立執行緒讀取 Outlook。
    - 新郵件：Outlook ItemAdd 事件，即時處理
    - 後備掃描：用 Outlook Table 只列出「上次之後」的郵件，已處理的郵件不會再打開
    - Outlook 分類：啟動時及每隔一段時間用 Table 讀取；你改分類時用 ItemChange 事件即時更新"""

    def __init__(self, app):
        super().__init__(daemon=True)
        self.app = app
        self.wake = threading.Event()
        self.cat_wake = threading.Event()
        self.full_requested = False
        self.manual_sync = False
        self.reset_folders = False
        self.reset_cat_folders = False
        self.seen = app.db.load_processed(app.cfg["days_back"])
        self.to_mark = []
        self.pending = []
        self.cat_pending = []
        self.cat_dirty_at = None
        self.hooks = []
        self.hooked_ids = None
        self.cat_hooks = []
        self._ns = None
        self._scan_folders = None
        self._cat_folders = None
        self.last_deep = 0
        self.last_popup_check = 0
        self.cat_items = {}
        self.cat_master = []
        self.item_fail = {}
        self.scan_opened = 0
        self.health = {"mode": "live", "running": False, "last_ok": None, "last_error": None,
                       "last_error_at": None, "duration": None, "scanned": 0, "new_events": 0,
                       "rows_read": 0, "totals": {}, "folders": [], "scans": 0, "outlook_user": "",
                       "instant": False, "last_instant": None, "progress": 0, "progress_at": None,
                       "scan_kind": "", "remembered": len(self.seen), "alive_at": iso(now()),
                       "cat_folders": [], "cat_live": False, "cat_refresh_at": None, "cat_duration": None,
                       "busy": "", "folder_errors": {}}
        self.categories = {"master": [], "items": [], "scanned_at": None, "error": None, "folders": 0}
        self.popup_alerted = {}
        self.popup_open = False

    # ---------- 外部要求 ----------
    def rescan(self, full=False, cats=False, folders=False, cat_folders=False, manual=False):
        if manual:
            self.manual_sync = True
        if full:
            self.full_requested = True
        if folders or full:
            self.reset_folders = True
        if cat_folders:
            self.reset_cat_folders = True
        if cats or cat_folders:
            self.cat_wake.set()
        self.wake.set()

    # ---------- Outlook 連線（重用，不會每次重新建立） ----------
    def session(self):
        if self._ns is None:
            outlook = win32com.client.Dispatch("Outlook.Application")
            self._ns = outlook.GetNamespace("MAPI")
            try:
                self.health["outlook_user"] = self._ns.CurrentUser.Name
            except Exception:
                pass
        return self._ns

    def drop_session(self):
        self._ns = None
        self._scan_folders = None
        self._cat_folders = None
        self.hooks, self.hooked_ids, self.cat_hooks = [], None, []
        self.health["instant"] = False
        self.health["cat_live"] = False

    def heartbeat(self, progress=None):
        h = self.health
        h["alive_at"] = iso(now())
        if progress is not None:
            h["progress"] = progress
            h["progress_at"] = h["alive_at"]

    def pump(self):
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        if self.pending:
            self.process_pending()
        if self.cat_pending:
            self.process_cat_pending()

    # ---------- 主迴圈 ----------
    def run(self):
        pythoncom.CoInitialize()
        ItemsEvents.scanner = self
        CatEvents.scanner = self
        next_scan = 0
        next_cat = time.time() + 20      # 先完成郵件掃描，才讀取分類
        while not self.app.stopping:
            cfg = self.app.cfg
            self.heartbeat()
            if self.full_requested:
                self.full_requested = False
                self.app.db.clear_scan_state()
                self.seen = set()
                self.last_deep = 0
            if self.reset_folders:
                self.reset_folders = False
                self._scan_folders = None
            if self.reset_cat_folders:
                self.reset_cat_folders = False
                self._cat_folders = None
            t = time.time()
            if self.wake.is_set() or t >= next_scan:
                self.wake.clear()
                manual, self.manual_sync = self.manual_sync, False
                # 「立即同步」只讀取新郵件；完整覆查只按時間表進行
                deep = (not manual) and t - self.last_deep >= cfg["deep_scan_minutes"] * 60
                self.scan_once(deep)
                next_scan = time.time() + cfg["scan_interval_seconds"]
            t = time.time()
            if self.cat_wake.is_set() or t >= next_cat or (self.cat_dirty_at and t >= self.cat_dirty_at):
                self.cat_wake.clear()
                self.cat_dirty_at = None
                self.refresh_categories()
                next_cat = time.time() + cfg["category_refresh_minutes"] * 60
            if time.time() - self.last_popup_check >= 5:
                self.last_popup_check = time.time()
                self.check_popups()
            self.pump()
            time.sleep(0.25)

    # ---------- 即時事件 ----------
    def process_pending(self):
        items, self.pending = self.pending, []
        for item, sent, path in items:
            try:
                item = win32com.client.Dispatch(item)
                if item.Class != 43:
                    continue
                store_id = item.Parent.StoreID
                self.handle(item, store_id, sent, None, path)
                self.health["last_instant"] = iso(now())
            except Exception:
                traceback.print_exc()
        self.flush_marks()
        self.check_popups(force=True)

    def process_cat_pending(self):
        items, self.cat_pending = self.cat_pending, []
        hidden = set(self.app.cfg.get("hidden_categories", []))
        changed = False
        for item in items:
            try:
                item = win32com.client.Dispatch(item)
                eid = item.EntryID
                rec = self.cat_record(item, item.Parent)
                if rec:
                    self.cat_items[eid] = rec
                    changed = True
                elif eid in self.cat_items:
                    del self.cat_items[eid]
                    changed = True
            except Exception:
                traceback.print_exc()
        if changed:
            self.publish_categories(hidden)

    # ---------- 處理一封郵件 ----------
    def handle(self, item, store_id, sent, when=None, folder_path=""):
        eid = item.EntryID
        if eid in self.seen:
            return False
        mail = OutlookMail(item, store_id, self.app.cfg, sent=sent)
        ev, reason = classify_explain(mail, self.app.cfg)
        self.seen.add(eid)
        self.to_mark.append((eid, iso(when or mail.time)))
        result = ""
        if ev:
            result = self.app.db.apply(ev) or ""
            self.health["totals"][ev["kind"]] = self.health["totals"].get(ev["kind"], 0) + 1
        try:
            rec = mail._recips  # 只記錄已經讀取過的收件人，避免額外讀取
            self.app.db.log_scan({
                "entry_id": eid, "time": iso(mail.time), "folder": folder_path + ("（寄件備份）" if sent else ""),
                "sender": mail.sender, "sender_name": mail.sender_name, "internal": 1 if mail.sender_internal else 0,
                "recipients": "; ".join(("CC:" if t == 2 else "BCC:" if t == 3 else "") + a for t, a in rec) if rec else "",
                "subject": mail.subject, "kind": ev["kind"] if ev else "", "reason": reason, "result": result,
                "replier": (ev or {}).get("replier"), "replier_how": (ev or {}).get("replier_how")})
        except Exception:
            traceback.print_exc()
        return bool(ev)

    def flush_marks(self):
        rows, self.to_mark = self.to_mark, []
        self.app.db.mark_processed(rows)

    # ---------- 要掃描的資料夾（只計算一次，設定改變時才重新計算） ----------
    def get_scan_folders(self, ns):
        if self._scan_folders is not None:
            return self._scan_folders
        cfg = self.app.cfg
        out = []
        paths = list(cfg.get("scan_folders") or [])
        if not paths:
            store = ns.DefaultStore
            try:
                out.extend((f, False) for f in iter_folders(store.GetDefaultFolder(6), cfg["include_subfolders"]))
            except Exception:
                out.append((ns.GetDefaultFolder(6), False))
            if cfg.get("scan_sent_items"):
                try:
                    out.append((store.GetDefaultFolder(5), True))
                except Exception:
                    pass
            paths = list(cfg.get("extra_folders") or [])
        for p in paths:
            try:
                f = get_folder_by_path(ns, p)
            except Exception:
                self.health["last_error"] = f"找不到資料夾：{p}"
                self.health["last_error_at"] = iso(now())
                continue
            sent = False
            try:
                sent = f.EntryID == f.Store.GetDefaultFolder(5).EntryID
            except Exception:
                pass
            if sent:
                out.append((f, True))
            else:
                out.extend((x, False) for x in iter_folders(f, cfg["include_subfolders"]))
        res, ids = [], set()
        for f, sent in out:
            try:
                eid = f.EntryID
                if eid in ids:
                    continue
                ids.add(eid)
                res.append({"folder": f, "sent": sent, "entry_id": eid, "store_id": f.StoreID,
                            "path": f.FolderPath})
            except Exception:
                continue
        self._scan_folders = res
        self.health["folders"] = [d["path"] + ("  (寄件備份)" if d["sent"] else "") for d in res]
        return res

    def hook(self, folders):
        ids = tuple(d["entry_id"] for d in folders)
        if ids == self.hooked_ids:
            return
        self.hooks = []
        ok = 0
        for d in folders:
            try:
                h = win32com.client.WithEvents(d["folder"].Items, ItemsEvents)
                h.sent = d["sent"]
                h.path = d["path"]
                self.hooks.append(h)
                ok += 1
            except Exception:
                traceback.print_exc()
        self.hooked_ids = ids
        self.health["instant"] = ok > 0

    def new_rows(self, folder, time_col, since):
        """用 Table 由新到舊列出郵件（只讀 EntryID / 時間 / 類型三個欄位），遇到比 since 舊就停止。"""
        table, _ = open_table(folder, ["EntryID", time_col, "MessageClass"], sort_col=time_col)
        stop = since - timedelta(minutes=1)
        count = [0]

        def on_batch():
            self.heartbeat(self.health.get("progress", 0))
            self.pump()

        for row in read_table(table, 3, batch=50, on_batch=on_batch):
            eid, when, mclass = (list(row) + [None, None, None])[:3]
            if not eid or when is None or not _is_dt(when):
                continue
            when = to_naive(when)
            if when < stop:
                return
            count[0] += 1
            yield eid, when, mclass or ""

    def scan_once(self, deep=False):
        cfg = self.app.cfg
        h = self.health
        h["running"] = True
        self.heartbeat(0)
        t0 = time.time()
        opened = new_events = rows = 0
        first = not self.app.db.meta("first_scan_done")
        h["scan_kind"] = "首次掃描" if first else ("完整覆查" if deep else "增量掃描")
        throttle = 0.03 if (first or deep) else 0.0     # 大量讀取時放慢，讓 Outlook 保持順暢
        try:
            ns = self.session()
            folders = self.get_scan_folders(ns)
            self.hook(folders)
            cutoff = now() - timedelta(days=cfg["days_back"])
            ok_folders = 0
            self.scan_opened = 0
            for d in folders:
                try:
                    r = self.scan_folder(ns, d, cutoff, deep, throttle)
                    self.scan_opened += r[0]
                    opened += r[0]; new_events += r[1]; rows += r[2]
                    h["folder_errors"].pop(d["path"], None)
                    ok_folders += 1
                except Exception as fe:
                    if self.app.stopping:
                        return
                    if is_session_error(fe):
                        raise                                  # Outlook 連線本身出問題
                    h["folder_errors"][d["path"]] = {"error": friendly_error(fe), "at": iso(now())}
                    traceback.print_exc()
                self.pump()
            if folders and not ok_folders:
                raise RuntimeError("全部資料夾都未能讀取：" + "；".join(v["error"] for v in h["folder_errors"].values()))
            h["last_ok"] = iso(now())
            h["scans"] += 1
            h["remembered"] = len(self.seen)
            if deep:
                self.last_deep = time.time()
                self.app.db.prune_scanlog()
            if first:
                self.app.db.set_meta("first_scan_done", iso(now()))
        except Exception as e:
            h["last_error"] = friendly_error(e) if not isinstance(e, RuntimeError) else str(e)
            h["last_error_at"] = iso(now())
            self.drop_session()
            traceback.print_exc()
        finally:
            self.flush_marks()
            h["running"] = False
            h["busy"] = ""
            h["duration"] = round(time.time() - t0, 1)
            h["scanned"] = opened
            h["rows_read"] = rows
            h["new_events"] = new_events
            self.heartbeat(opened)

    def scan_folder(self, ns, d, cutoff, deep, throttle):
        """掃描一個資料夾；回傳 (打開, 事件, 列出)。"""
        h = self.health
        opened = new_events = rows = 0
        key = "wm:" + d["entry_id"]
        wm = parse_iso(self.app.db.meta(key))
        since = cutoff if (deep or wm is None) else max(cutoff, wm - timedelta(minutes=15))
        started = now()
        h["busy"] = d["path"]
        retry_later = False
        for eid, when, mclass in self.new_rows(d["folder"], "SentOn" if d["sent"] else "ReceivedTime", since):
            if self.app.stopping:
                break
            rows += 1
            if eid in self.seen:
                continue
            if not is_mail_class(mclass):          # 會議邀請、退信報告等不用打開
                self.seen.add(eid)
                self.to_mark.append((eid, iso(when)))
                continue
            try:
                item = ns.GetItemFromID(eid, d["store_id"])
                opened += 1
                if self.handle(item, d["store_id"], d["sent"], when, d["path"]):
                    new_events += 1
                self.item_fail.pop(eid, None)
            except Exception as ie:
                if is_session_error(ie):
                    raise
                # 暫時讀取失敗（例如網絡問題）：下次再試；連續失敗 3 次才略過，以免漏掉客戶郵件
                n = self.item_fail.get(eid, 0) + 1
                self.item_fail[eid] = n
                if n >= 3:
                    self.seen.add(eid)
                    self.to_mark.append((eid, iso(when)))
                    self.item_fail.pop(eid, None)
                else:
                    retry_later = True
                h["last_item_error"] = friendly_error(ie)
                traceback.print_exc()
            if opened % 10 == 0:
                self.heartbeat(self.scan_opened + opened)
                self.flush_marks()
                self.pump()
            if throttle:
                time.sleep(throttle)
        self.flush_marks()
        if not retry_later:
            self.app.db.set_meta(key, iso(started))
        return opened, new_events, rows

    # ---------- Outlook 分類 ----------
    def get_cat_folders(self, ns):
        if self._cat_folders is not None:
            return self._cat_folders
        cfg = self.app.cfg
        out = []
        paths = list(cfg.get("category_folders") or [])
        if paths:
            for p in paths:
                try:
                    f = get_folder_by_path(ns, p)
                    kind = "task" if getattr(f, "DefaultItemType", 0) == 3 else "mail"
                    out.append((f, kind))
                except Exception:
                    self.categories["error"] = f"找不到資料夾：{p}"
        else:
            store = ns.DefaultStore
            try:
                inbox = store.GetDefaultFolder(6)
                out.append((inbox, "mail"))
                # 收件匣的子資料夾，但不包括大資料夾（例如 Alerts、Archive）
                for sub in iter_folders(inbox, True):
                    if sub.EntryID == inbox.EntryID:
                        continue
                    n = folder_count(sub)
                    if 0 <= n <= BIG_FOLDER:
                        out.append((sub, "mail"))
            except Exception:
                traceback.print_exc()
            try:
                out.append((store.GetDefaultFolder(13), "task"))
            except Exception:
                pass
        res = []
        for f, kind in out:
            try:
                res.append({"folder": f, "kind": kind, "entry_id": f.EntryID, "store_id": f.StoreID,
                            "path": f.FolderPath})
            except Exception:
                pass
        self._cat_folders = res
        self.health["cat_folders"] = [d["path"] for d in res]
        self.cat_hooks = []
        ok = 0
        for d in res:
            try:
                self.cat_hooks.append(win32com.client.WithEvents(d["folder"].Items, CatEvents))
                ok += 1
            except Exception:
                traceback.print_exc()
        self.health["cat_live"] = ok > 0
        return res

    def refresh_categories(self):
        cfg = self.app.cfg
        hidden = set(cfg.get("hidden_categories", []))
        t0 = time.time()
        try:
            ns = self.session()
            master = []
            for c in ns.Categories:
                try:
                    master.append({"name": c.Name, "color": OUTLOOK_COLORS.get(int(c.Color), "#9ca3af")})
                except Exception:
                    pass
            self.cat_master = master
            items, errs = {}, []
            for d in self.get_cat_folders(ns):
                self.heartbeat()
                try:
                    for rec in self.cat_table(d):
                        items[rec["entry_id"]] = rec
                except Exception as fe:
                    if is_session_error(fe):
                        raise
                    errs.append(f"{d['path']}：{friendly_error(fe)}")
                    for eid, rec in self.cat_items.items():    # 保留上次讀到的結果
                        if rec.get("folder") == d["path"]:
                            items[eid] = rec
                self.pump()
            self.cat_items = items
            self.categories["error"] = "；".join(errs) or None
            self.health["cat_refresh_at"] = iso(now())
            self.health["cat_duration"] = round(time.time() - t0, 1)
            self.publish_categories(hidden)
        except Exception as e:
            self.categories = dict(self.categories, error=friendly_error(e), scanned_at=iso(now()))
            self.drop_session()
            traceback.print_exc()

    def cat_table(self, d):
        """用 Table 讀取一個資料夾內有分類的項目（不用逐封打開）。"""
        task = d["kind"] == "task"
        tcol = "CreationTime" if task else "ReceivedTime"
        cols = ["EntryID", tcol, "Subject", "Categories", "MessageClass"]
        cols += [PID_TASK_DUE, PID_TASK_COMPLETE] if task else ["SenderName", "UnRead", PR_FLAG_STATUS, PID_TASK_DUE]
        table, names = open_table(d["folder"], cols, filt=CAT_FILTER)
        off = datetime.now().astimezone().utcoffset() or timedelta(0)
        out = []
        for row in read_table(table, len(names), batch=100, on_batch=self.heartbeat):
            r = dict(zip(names, row))
            cats = [c.strip() for c in re.split(r"[;,]", r.get("Categories") or "") if c.strip()]
            if not cats or not r.get("EntryID"):
                continue
            when = r.get(tcol)
            due = r.get(PID_TASK_DUE)
            due_s = None
            if due is not None and _is_dt(due) and due.year < 4000:
                due_s = iso(datetime(due.year, due.month, due.day, due.hour, due.minute) + off)
            flag = r.get(PR_FLAG_STATUS)
            subject = r.get("Subject") or ""
            out.append({
                "entry_id": r["EntryID"], "store_id": d["store_id"], "folder": d["path"], "subject": subject,
                "categories": cats, "cls": 48 if task else 43,
                "sender": "工作" if task else (r.get("SenderName") or ""),
                "received": iso(to_naive(when)) if when is not None and _is_dt(when) else None,
                "unread": bool(r.get("UnRead")) if not task else False,
                "flagged": True if task else flag == 2,
                "complete": bool(r.get(PID_TASK_COMPLETE)) if task else flag == 1,
                "due": due_s if (task or flag == 2) else None,
                "ticket_id": ticket_ref(subject)})
        return out

    def publish_categories(self, hidden):
        items = []
        for rec in self.cat_items.values():
            cats = [c for c in rec["categories"] if c not in hidden]
            if cats:
                items.append(dict(rec, categories=cats))
        master = list(self.cat_master)
        known = {c["name"] for c in master}
        for rec in items:
            for c in rec["categories"]:
                if c not in known:
                    known.add(c)
                    master.append({"name": c, "color": "#9ca3af"})
        self.categories = {"master": master, "items": items, "scanned_at": iso(now()),
                           "error": self.categories.get("error"), "folders": len(self._cat_folders or [])}

    @staticmethod
    def cat_record(it, folder):
        cats = [c.strip() for c in re.split(r"[;,]", it.Categories or "") if c.strip()]
        if not cats:
            return None
        cls = it.Class
        rec = {"entry_id": it.EntryID, "store_id": folder.StoreID, "folder": folder.FolderPath,
               "subject": it.Subject or "", "categories": cats, "cls": cls, "sender": "", "received": None,
               "unread": False, "flagged": False, "due": None, "complete": False}
        if cls == 43:
            rec["sender"] = it.SenderName or ""
            rec["received"] = iso(to_naive(it.ReceivedTime))
            rec["unread"] = bool(it.UnRead)
            try:
                rec["flagged"] = it.FlagStatus == 2
                rec["complete"] = it.FlagStatus == 1
                if rec["flagged"]:
                    d = it.TaskDueDate
                    if d and d.year < 4000:
                        rec["due"] = iso(to_naive(d))
            except Exception:
                pass
        elif cls == 48:  # 工作
            rec["sender"] = "工作"
            try:
                rec["received"] = iso(to_naive(it.CreationTime))
                d = it.DueDate
                if d and d.year < 4000:
                    rec["due"] = iso(to_naive(d))
                rec["complete"] = bool(it.Complete)
                rec["flagged"] = True
            except Exception:
                pass
        else:
            try:
                rec["received"] = iso(to_naive(it.CreationTime))
            except Exception:
                pass
        rec["ticket_id"] = ticket_ref(rec["subject"])
        return rec

    # ---------- 視窗關閉時用 Windows 彈窗提醒 ----------
    def check_popups(self, force=False):
        cfg = self.app.cfg
        if not cfg.get("desktop_popup") or os.name != "nt" or self.popup_open:
            return
        if time.time() - self.app.last_ui_poll < 20:
            return
        n = now()
        base = parse_iso(self.app.db.meta("baseline_at")) or n
        urgent, top = [], 0
        for tk in self.app.db.q("SELECT id, subject, company, requester, sla_start, origin FROM tickets WHERE ignored=0 AND "
                                "first_response_at IS NULL AND sla_start IS NOT NULL AND sla_start>=?",
                                (iso(n - timedelta(hours=cfg["unverified_hours"])),)):
            st = parse_iso(tk["sla_start"])
            if tk["origin"] == "email" and st < base - timedelta(minutes=30):
                continue
            el = (n - st).total_seconds() / 60
            lvl = 4 if el >= cfg["sla_minutes"] else 3 if el >= cfg["critical_minutes"] else 2 if el >= cfg["warning_minutes"] else 1
            if self.popup_alerted.get(tk["id"], 0) < lvl:
                self.popup_alerted[tk["id"]] = lvl
                top = max(top, lvl)
                label = f"#{tk['id']}" if tk["origin"] != "email" else "新郵件"
                urgent.append(f"{label}  {tk['company'] or tk['requester'] or ''}  {(tk['subject'] or '')[:40]}  ({int(el)} 分鐘)")
        if urgent:
            head = "有新客戶郵件，請在 %d 分鐘內回覆：" % cfg["sla_minutes"] if top == 1 else "以下郵件尚未首次回覆："
            msg = head + "\n\n" + "\n".join(urgent[:8]) + "\n\n按「確定」打開 PulseDesk。"
            self.popup_open = True

            def show():
                import ctypes
                try:
                    r = ctypes.windll.user32.MessageBoxW(0, msg, "PulseDesk 提醒", 0x40000 | 0x1000 | 0x30 | 0x1)
                    if r == 1:
                        open_ui(self.app.url)
                finally:
                    self.popup_open = False
            threading.Thread(target=show, daemon=True).start()


# =====================================================================
#  Outlook 動作（開郵件 / 通知同事草稿）
# =====================================================================
def with_outlook(fn):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return fn(outlook, outlook.GetNamespace("MAPI"))
    finally:
        pythoncom.CoUninitialize()


def open_mail(entry_id, store_id):
    def fn(outlook, ns):
        item = ns.GetItemFromID(entry_id, store_id) if store_id else ns.GetItemFromID(entry_id)
        item.Display()
        try:
            item.GetInspector.Activate()
        except Exception:
            pass
    with_outlook(fn)


KIND_ZH = {"customer_new": "客戶來信", "new_ticket": "Freshdesk 建立工單", "ack": "Ticket Received 自動回覆",
           "agent_reply": "同事回覆客戶", "direct_reply": "同事直接用 Outlook 回覆", "customer_reply": "客戶再回覆", "": "（不處理）"}
KEEP_RE = re.compile(r"(Tkt\s*#\s*\d+|\[#\d+\]|^\s*(?:RE|FW|FWD|回覆|轉寄)\s*[:：]|\[External\]|Ticket Received|"
                     r"has created a new Freshdesk ticket|New comment)", re.I)


def mask_addr(addr, cfg):
    if not addr or "@" not in addr or is_internal_addr(addr, cfg):
        return addr or ""
    local, dom = addr.split("@", 1)
    return local[:1] + "***@" + dom


def mask_name(name):
    name = (name or "").strip()
    return name[:1] + "***" if name else ""


def mask_subject(subj):
    out, pos = [], 0
    for m in KEEP_RE.finditer(subj or ""):
        out.append(re.sub(r"[^\s\-:：/()\[\]#.,]", "•", subj[pos:m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(re.sub(r"[^\s\-:：/()\[\]#.,]", "•", (subj or "")[pos:]))
    return "".join(out)


def diagnostics_csv(app, days=3, m_email=True, m_subject=False):
    import csv
    import io
    cfg = app.cfg
    lo = iso(now() - timedelta(days=days))
    rows = app.db.q("SELECT s.*, m.ticket_id FROM scanlog s LEFT JOIN mails m ON m.entry_id=s.entry_id "
                    "WHERE s.time>=? ORDER BY s.time DESC", (lo,))
    tickets = {t["id"]: t for t in app.db.q("SELECT * FROM tickets")}
    corr = {}
    for c in app.db.q("SELECT * FROM corrections ORDER BY id"):
        corr.setdefault(c["ticket_id"], []).append(f"{c['action']}：{c['detail']}")
        m2 = re.search(r"合併到 (\S+)|Tkt#(\d+)", c["detail"] or "")
        if m2:
            corr.setdefault(m2.group(1) or m2.group(2), []).append(f"{c['action']}（來自 {c['ticket_id']}）")

    def status(t):
        if not t:
            return ""
        if t["ignored"]:
            return "已忽略"
        if not t["first_response_at"]:
            return "未回覆"
        st, fr = parse_iso(t["sla_start"]), parse_iso(t["first_response_at"])
        if st and fr:
            mins = (fr - st).total_seconds() / 60
            return f"已回覆（{mins:.0f} 分鐘，{'達標' if mins <= cfg['sla_minutes'] else '超時'}）"
        return "已回覆"

    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["收到時間", "資料夾", "寄件人名稱", "寄件人電郵", "內部寄件人", "收件人", "主旨", "判斷類型",
                "判斷原因", "處理結果", "回覆同事", "辨認依據", "工單", "工單現況", "人手修正", "判斷正確？(Y/N)",
                "正確應該是", "備註"])
    for r in rows:
        tid = r["ticket_id"] or ""
        t = tickets.get(tid)
        rec = r["recipients"] or ""
        if m_email:
            rec = "; ".join(("CC:" if x.startswith("CC:") else "BCC:" if x.startswith("BCC:") else "")
                            + mask_addr(x.split(":", 1)[1] if x[:3] in ("CC:", "BCC") else x, cfg)
                            for x in rec.split("; ") if x)
        w.writerow([r["time"].replace("T", " "), r["folder"],
                    mask_name(r["sender_name"]) if m_email and not r["internal"] else r["sender_name"],
                    mask_addr(r["sender"], cfg) if m_email else r["sender"],
                    "是" if r["internal"] else "", rec,
                    mask_subject(r["subject"]) if m_subject else r["subject"],
                    KIND_ZH.get(r["kind"] or "", r["kind"]), r["reason"], r["result"],
                    r.get("replier") or "", r.get("replier_how") or "",
                    ("#" + tid) if tid and not tid.startswith("E") else ("新郵件 " + tid if tid else ""),
                    status(t), "；".join(corr.get(tid, [])), "", "", ""])
    info = [[], ["# PulseDesk %s 診斷資料，匯出時間 %s，最近 %d 日，共 %d 封" % (VERSION, iso(now()), days, len(rows))],
            ["# 內部網域：" + ", ".join(cfg["internal_domains"]) + "；支援信箱：" + ", ".join(cfg["support_addresses"])],
            ["# 忽略寄件人：" + ", ".join(cfg["ignore_senders"])], ["# 忽略主旨：" + ", ".join(cfg["ignore_subjects"])]]
    for line in info:
        w.writerow(line)
    return ("﻿" + buf.getvalue()).encode("utf-8")


def list_folders():
    """列出可以選擇的資料夾（每個信箱的收件匣、其子資料夾、寄件備份及工作），連同郵件數量。"""
    def fn(outlook, ns):
        out = []
        default_id = ns.DefaultStore.StoreID
        for store in ns.Stores:
            try:
                name = store.DisplayName
                is_default = store.StoreID == default_id
            except Exception:
                continue
            for fid, kind in ((6, "inbox"), (5, "sent"), (13, "task")):
                if kind == "task" and not is_default:
                    continue
                try:
                    f = store.GetDefaultFolder(fid)
                except Exception:
                    continue
                out.append({"path": f.FolderPath, "name": f.Name, "store": name, "kind": kind,
                            "default": is_default, "depth": 0, "count": folder_count(f)})
                if kind == "inbox":
                    def walk(folder, depth):
                        if depth > 2:
                            return
                        try:
                            for sub in folder.Folders:
                                out.append({"path": sub.FolderPath, "name": sub.Name, "store": name,
                                            "kind": "sub", "default": is_default, "depth": depth,
                                            "count": folder_count(sub)})
                                walk(sub, depth + 1)
                        except Exception:
                            pass
                    walk(f, 1)
        return out
    return with_outlook(fn)


def draft_to_colleague(app, tk, colleague):
    cfg = app.cfg
    is_fd = tk.get("origin") != "email"
    link = f"{cfg['freshdesk_url']}/helpdesk/tickets/{tk['id']}"
    head = f"<b>Tkt#{tk['id']}</b> — " if is_fd else ""
    intro = (f"<p>Hi {colleague['name']},</p><p>請跟進以下客戶請求：</p>"
             f"<p>{head}{tk['subject']}<br>客戶：{tk['company'] or ''} {tk['requester'] or ''} "
             f"{tk['requester_email'] or ''}" + (f"<br><a href=\"{link}\">{link}</a>" if is_fd else "") + "</p>")
    if tk.get("note"):
        intro += f"<p>備註：{tk['note']}</p>"
    intro += "<p>Thanks!</p><hr>"

    def fn(outlook, ns):
        src = tk.get("customer_entry") or tk.get("notif_entry") or tk.get("ack_entry")
        mail = None
        if src:
            try:
                orig = ns.GetItemFromID(src, tk["store_id"]) if tk.get("store_id") else ns.GetItemFromID(src)
                mail = orig.Forward()
            except Exception:
                mail = None
        if mail is None:
            mail = outlook.CreateItem(0)
            mail.HTMLBody = ""
        mail.To = colleague.get("email", "")
        mail.Subject = (f"[跟進] Tkt#{tk['id']} " if is_fd else "[跟進] ") + (tk["subject"] or "")
        mail.HTMLBody = intro + (mail.HTMLBody or "")
        mail.Display()  # 只開草稿，由使用者自己傳送
    with_outlook(fn)


# =====================================================================
#  應用程式 / HTTP API
# =====================================================================
class App:
    def __init__(self, demo=False):
        self.demo = demo
        self.stopping = False
        self.last_ui_poll = 0
        self.cfg = load_config()
        if demo:
            path = os.path.join(DATA_DIR, "demo.db")
            if os.path.exists(path):
                os.remove(path)
            self.db = DB(path)
            self.db.set_meta("baseline_at", iso(now() - timedelta(days=10)))
            if not self.cfg["colleagues"]:
                self.cfg["colleagues"] = [{"name": n, "email": ""} for n in ("陳大文", "李小欣", "黃志明", "張美玲")]
            seed_demo(self.db)
            self.demo_categories = demo_categories()
            self.scanner = None
        else:
            self.db = DB(os.path.join(DATA_DIR, "pulsedesk.db"))
            self.scanner = Scanner(self) if HAS_COM else None
        self.url = f"http://127.0.0.1:{self.cfg['port']}/"

    def health(self):
        if self.demo:
            return {"mode": "demo", "last_ok": iso(now()), "running": False, "scans": 1, "totals": {},
                    "folders": ["\\\\kwong@premier-technology.com\\Inbox", "\\\\kwong@premier-technology.com\\Sent Items  (寄件備份)"],
                    "scanned": 3, "rows_read": 5, "new_events": 1, "duration": 0.4, "instant": True,
                    "scan_kind": "增量掃描", "remembered": 612, "outlook_user": "Kelvin Wong", "alive_at": iso(now()),
                    "cat_folders": ["\\\\kwong@premier-technology.com\\Inbox", "\\\\kwong@premier-technology.com\\Inbox\\Need to Follow up",
                                    "\\\\kwong@premier-technology.com\\Tasks"],
                    "cat_live": True, "cat_refresh_at": iso(now()), "cat_duration": 0.8}
        if not self.scanner:
            return {"mode": "no_outlook", "last_ok": None, "last_error": "未安裝 pywin32 或不是 Windows，無法連接 Outlook。",
                    "last_error_at": iso(now()), "totals": {}, "folders": []}
        return self.scanner.health

    def categories(self):
        if self.demo:
            return self.demo_categories
        if not self.scanner:
            return {"master": [], "items": [], "scanned_at": None, "error": None}
        return self.scanner.categories

    def state(self):
        self.last_ui_poll = time.time()
        pub = {k: self.cfg[k] for k in DEFAULT_CONFIG if k != "port"}
        return {"server_now": iso(now()), "version": VERSION, "demo": self.demo,
                "baseline_at": self.db.meta("baseline_at"), "config": pub, "tickets": self.db.tickets(),
                "followups": self.db.followups(), "recent": self.db.recent_mails(),
                "corrections": self.db.q("SELECT * FROM corrections ORDER BY id DESC LIMIT 40"),
                "health": self.health(), "categories": self.categories()}

    def add_ignore(self, key, value):
        value = (value or "").strip().lower()
        if value and value not in self.cfg[key]:
            self.cfg[key] = self.cfg[key] + [value]
            if not self.demo:
                save_config(self.cfg)

    def bulk_action(self, ids, body):
        done = 0
        for tid in ids[:1000]:
            try:
                self.ticket_action(str(tid), body)
                done += 1
            except Exception:
                pass
        return {"ok": True, "done": done}

    def ticket_action(self, tid, body):
        tk = self.db.get(tid)
        if not tk:
            raise ValueError("找不到工單")
        a = body.get("action")
        n = iso(now())
        if a == "not_request":
            scope = body.get("scope", "this")
            addr = (tk["requester_email"] or "").lower()
            self.db.update(tid, ignored=1)
            detail = "只此一張"
            if scope == "sender" and addr:
                self.add_ignore("ignore_senders", addr)
                self.db.x("UPDATE tickets SET ignored=1, updated_at=? WHERE requester_email=? AND first_response_at IS NULL", (n, addr))
                detail = f"以後忽略寄件人 {addr}"
            elif scope == "domain" and "@" in addr:
                dom = "@" + domain_of(addr)
                if dom[1:] in self.cfg["internal_domains"]:
                    raise ValueError("不能忽略公司內部網域")
                self.add_ignore("ignore_senders", dom)
                self.db.x("UPDATE tickets SET ignored=1, updated_at=? WHERE requester_email LIKE ? AND first_response_at IS NULL",
                          (n, "%" + dom))
                detail = f"以後忽略網域 {dom}"
            elif scope == "subject":
                subj = clean_subject(tk["subject"]).lower()
                if len(subj) < 4:
                    raise ValueError("主旨太短，不能用作忽略規則")
                self.add_ignore("ignore_subjects", subj)
                detail = f"以後忽略主旨「{subj}」"
            self.db.log(tid, "人手修正：不是客戶請求（" + detail + "）")
            self.db.add_correction(tk, "不是客戶請求", detail)
        elif a == "merge":
            dst = str(body.get("target") or "").strip()
            if not dst or dst == tid or not self.db.get(dst):
                raise ValueError("請選擇要合併到的工單")
            self.db.merge_into(tid, dst, f"人手合併：{tk['subject'] or tid} 已合併到這張工單")
            self.db.add_correction(tk, "合併工單", f"合併到 {dst}")
            return {"ok": True, "ticket": self.db.get(dst), "moved_to": dst}
        elif a == "link":
            num = re.sub(r"\D", "", str(body.get("number") or ""))
            if len(num) < 3:
                raise ValueError("請輸入正確的 Freshdesk 工單號碼")
            if num == tid:
                raise ValueError("已經是這張工單")
            self.db.merge_into(tid, num, f"人手連結到 Freshdesk 工單 Tkt#{num}")
            self.db.add_correction(tk, "連結 Freshdesk 工單", f"Tkt#{num}")
            return {"ok": True, "ticket": self.db.get(num), "moved_to": num}
        elif a == "not_reply":
            self.db.update(tid, first_response_at=None, response_source=None, reply_entry=None,
                           first_responder=None, first_responder_how=None)
            self.db.log(tid, "人手修正：偵測到的郵件不是回覆客戶，重新計時")
            self.db.add_correction(tk, "不是回覆", f"原本偵測的回覆時間 {tk['first_response_at']}")
        elif a == "respond":
            st = parse_iso(tk["sla_start"])
            late = st is None or (now() - st) > timedelta(hours=self.cfg["unverified_hours"])
            unknown = bool(body.get("unknown")) or late
            # 事後補記的「已回覆」不知道真正回覆時間，不計入達標率及平均回覆時間
            self.db.update(tid, first_response_at=n, response_source="manual_unknown" if unknown else "manual",
                           first_responder=None, first_responder_how="人手標記")
            self.db.log(tid, "人手標記為已首次回覆" + ("（回覆時間不明，不計入統計）" if unknown else ""))
            self.db.add_correction(tk, "人手標記已回覆", "回覆時間不明" if unknown else "未偵測到回覆郵件")
        elif a == "unrespond":
            self.db.update(tid, first_response_at=None, response_source=None, first_responder=None, first_responder_how=None)
            self.db.log(tid, "取消「已回覆」標記")
        elif a == "assign":
            who = (body.get("assignee") or "").strip()
            if who:
                upd = {"assignee": who, "assigned_at": n}
                if tk["status"] in (None, "open"):
                    upd.update(status="progress", status_at=n)
                self.db.update(tid, **upd)
                self.db.log(tid, f"分派給 {who}")
            else:
                self.db.update(tid, assignee=None, assigned_at=None, status="open", status_at=n)
                self.db.log(tid, "取消分派")
        elif a == "status":
            st = body.get("status")
            if st not in ("open", "progress", "waiting", "resolved", "closed"):
                raise ValueError("狀態不正確")
            self.db.update(tid, status=st, status_at=n)
            self.db.log(tid, "狀態改為 " + {"open": "待處理", "progress": "跟進中", "waiting": "等待客戶", "resolved": "已解決",
                                           "closed": "已處理（不需分派）"}[st])
        elif a == "set_responder":
            who = (body.get("responder") or "").strip()
            self.db.update(tid, first_responder=who or None, first_responder_how="人手修正" if who else None)
            self.db.log(tid, f"人手修正首次回覆同事：{who or '（清除）'}")
            self.db.add_correction(tk, "修正回覆同事", f"{tk.get('first_responder') or '未知'} → {who or '（清除）'}")
        elif a == "note":
            self.db.update(tid, note=str(body.get("note", ""))[:4000])
        elif a in ("ignore", "ignore_sender"):
            self.db.update(tid, ignored=1)
            self.db.log(tid, "標記為忽略（不計 SLA）")
            self.db.add_correction(tk, "忽略", "以後忽略此寄件人" if a == "ignore_sender" else "只此一張")
            if a == "ignore_sender" and tk["requester_email"]:
                addr = tk["requester_email"].lower()
                if addr not in self.cfg["ignore_senders"]:
                    self.cfg["ignore_senders"] = self.cfg["ignore_senders"] + [addr]
                    if not self.demo:
                        save_config(self.cfg)
                self.db.x("UPDATE tickets SET ignored=1, updated_at=? WHERE requester_email=? AND first_response_at IS NULL",
                          (n, addr))
        elif a == "unignore":
            self.db.update(tid, ignored=0)
            self.db.log(tid, "取消忽略")
        else:
            raise ValueError("未知操作")
        return {"ok": True, "ticket": self.db.get(tid)}


def seed_demo(db):
    n = now()
    samples = [
        ("E1A2B3C4D5E", "Copilot Enterprise User Transfer", "", "Aico Li", "ali@coreviewcap.com", 1, None, None, "open", "email"),
        ("254333", "Susan and Risk email group", "Ocean Arete", "Iyrus Lam", "ilam@aretefunds.com", 3, None, None, "open", "freshdesk"),
        ("E9F8E7D6C5B", "Derek's Shanghai Tang access issue", "", "Ada Tse", "atse@lunargp.com", 7, None, None, "open", "email"),
        ("254335", "VPN keeps disconnecting", "Arete Funds", "Michael Edgar", "michael.edgar@cim.com.hk", 11, None, None, "open", "freshdesk"),
        ("254336", "新同事電腦設定 (下星期一入職)", "GS Client Tech", "Joanne Ho", "jho@gsclient.com", 14, None, None, "open", "freshdesk"),
        ("254337", "Printer on 12/F offline", "Ocean Arete", "Tom Lee", "tlee@aretefunds.com", 17, None, None, "open", "freshdesk"),
        ("254330", "I cannot open password protected zip on my PC", "CIM", "Michael Edgar", "michael.edgar@cim.com.hk", 42, 9, None, "open", "freshdesk"),
        ("254328", "Fund admin email - Edith", "Arete Funds", "Edith Chan", "echan@aretefunds.com", 70, 12, None, "open", "freshdesk"),
        ("254320", "Teams 會議室顯示屏無畫面", "Cheetah Investment", "Kelvin Lau", "klau@cheetah-inv.com", 150, 6, "Steven Li", "progress", "freshdesk"),
        ("254318", "Report delay enquiry", "GS Client Tech", "Ada Ng", "ada@gsclient.com", 200, 18, "Sai Ho", "waiting", "freshdesk"),
        ("254310", "Laptop replacement request", "Ocean Arete", "Sam Yip", "syip@aretefunds.com", 300, 7, "Benny Chan", "progress", "freshdesk"),
        ("254305", "Mailbox full warning", "CIM", "Peter Cheung", "pcheung@cim.com.hk", 260, 4, "Chester Choi", "resolved", "freshdesk"),
        ("254301", "非管理員也能安裝程序", "Premier Client", "Leo Chan", "leo@client.com", 330, 11, "Steven Li", "resolved", "freshdesk"),
    ]
    for tid, subj, comp, req, email, ago, resp, who, st, origin in samples:
        start = n - timedelta(minutes=ago)
        db.x("INSERT INTO tickets(id,subject,company,requester,requester_email,sla_start,first_response_at,"
             "response_source,assignee,assigned_at,status,status_at,seen_at,updated_at,origin) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
             (tid, subj, comp, req, email, iso(start), iso(start + timedelta(minutes=resp)) if resp else None,
              "email" if resp else None, who, iso(start + timedelta(minutes=(resp or 0) + 20)) if who else None,
              st, iso(n - timedelta(minutes=30)) if st == "resolved" else None, iso(start), iso(n), origin))
        db.log(tid, "收到客戶郵件，開始計時", start)
        if resp:
            db.log(tid, "已在 Freshdesk 回覆客戶（Re: Tkt# 郵件）", start + timedelta(minutes=resp))
        if who:
            db.log(tid, f"分派給 {who}", start + timedelta(minutes=resp + 20))
    db.x("UPDATE tickets SET last_customer_reply_at=? WHERE id='254318'", (iso(n - timedelta(minutes=12)),))
    import random
    rnd = random.Random(7)
    k = 254100
    for d in range(1, 7):
        for _ in range(rnd.randint(6, 14)):
            k += 1
            start = (n - timedelta(days=d)).replace(hour=rnd.randint(9, 18), minute=rnd.randint(0, 59))
            resp = rnd.choice([3, 5, 6, 8, 9, 11, 12, 13, 14, 16, 22])
            who = rnd.choice(["Steven Li", "Sai Ho", "Benny Chan", "Chester Choi"])
            db.x("INSERT INTO tickets(id,subject,company,requester,sla_start,first_response_at,response_source,"
                 "assignee,assigned_at,status,status_at,seen_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                 (str(k), rnd.choice(["Password reset", "Email issue", "New user setup", "VPN access", "Printer issue"]),
                  rnd.choice(["Ocean Arete", "CIM", "Cheetah Investment", "GS Client Tech"]), "User",
                  iso(start), iso(start + timedelta(minutes=resp)), "email", who,
                  iso(start + timedelta(minutes=resp + 15)), "resolved", iso(start + timedelta(hours=3)),
                  iso(start), iso(start + timedelta(hours=3))))
    import random as _r
    rr = _r.Random(3)
    staff = [c["name"] for c in DEFAULT_STAFF if c["name"] != "Alan Wong 2"]
    for row in db.q("SELECT id FROM tickets WHERE first_response_at IS NOT NULL"):
        if rr.random() < 0.08:
            db.x("UPDATE tickets SET first_responder_how='未能辨認簽名（最後一行：Premier Technology Support Team）' WHERE id=?", (row["id"],))
        else:
            who = rr.choice(staff[:8])
            db.x("UPDATE tickets SET first_responder=?, first_responder_how=? WHERE id=?",
                 (who, "簽名「%s」" % next(c["signature"].replace("\n", " / ") for c in DEFAULT_STAFF if c["name"] == who), row["id"]))
    demo_log = [
        (1, "ali@coreviewcap.com", "Aico Li", 0, "support@premier-technology.com", "Copilot Enterprise User Transfer",
         "customer_new", "外部寄件人，收件人包括 support@premier-technology.com", "新請求 E1A2B3C4D5E，開始計時"),
        (3, "ilam@aretefunds.com", "Iyrus Lam", 0, "support@premier-technology.com", "Susan and Risk email group",
         "customer_new", "外部寄件人，收件人包括 support@premier-technology.com", "合併到 254333，計時由這封郵件開始"),
        (2, "support@premier-technology.com", "Premier Technology Support", 1, "servicedesk@premier-technology.com",
         "Ocean Arete - Iyrus Lam has created a new Freshdesk ticket: Susan and Risk email group", "new_ticket",
         "主旨是「has created a new Freshdesk ticket」，工單號碼來自郵件連結", "建立 / 更新工單 #254333"),
        (13, "johnpttester@gmail.com", "John PTTester", 0, "support@premier-technology.com",
         "Test to support@premier-technology.com - Please ignore - A01", "", "忽略主旨「please ignore」", ""),
        (20, "mea_premier@lenovo.com", "MEA Premier Support", 0, "servicedesk@premier-technology.com",
         "[External] RE: Email to ionpacific@pt Lenovo Case 2032973157", "customer_new",
         "外部寄件人，收件人包括 servicedesk@premier-technology.com", "新請求 E7777777777，開始計時"),
        (25, "newsletter@vendor.com", "Vendor News", 0, "kwong@premier-technology.com", "October product update",
         "", "外部郵件，但收件人沒有 support / servicedesk", ""),
    ]
    for ago, snd, name, internal, rcp, subj, kind, reason, result in demo_log:
        db.log_scan({"entry_id": f"demo-log-{ago}", "time": iso(n - timedelta(minutes=ago)),
                     "folder": "\\\\kwong@premier-technology.com\\Inbox", "sender": snd, "sender_name": name,
                     "internal": internal, "recipients": rcp, "subject": subj, "kind": kind, "reason": reason,
                     "result": result})
    db.x("INSERT INTO mails(entry_id,kind,ticket_id,time,subject,sender,sender_name) VALUES(?,?,?,?,?,?,?)",
         ("demo-f1", "customer_reply", "254299", iso(n - timedelta(minutes=15)),
          "Proxy for connecting Singapore Server (Urgent)", "tchau@centerlineim.com", "Thomas Chau"))


def demo_categories():
    n = now()
    master = [("Assigned need to follow up", 4), ("Email sample", 9), ("Need to follow up", 8),
              ("Need to follow up TODAY!", 24), ("To Do", 1), ("Tools", 5), ("Wait for reply", 14),
              ("Waiting for assign", 2)]
    data = {
        "Assigned need to follow up": ["Morgan Stanley - KUARK CAPITAL onboarding", "Tkt#249400 RE: ACTION REQUIRED",
                                       "Centerline FW: Tkt#249276 NAS backup", "Tkt#249466 CoreView - Cloud migration",
                                       "Tenucia FW: Tkt#249809 Fw: printer", "Tkt#249672 TruMed - New laptop"],
        "Need to follow up": ["AIIM Investment Management - Teams phone", "New comment - [#249539] Proxy issue",
                              "Tkt#246974 Sharing Drive - \"Finance\"", "Tkt#251789 FW: Potential Collab",
                              "Tkt#252750 Nezu Tokyo annual renewal", "Peach Creek Capital - Network upgrade",
                              "Server PW issue", "Redwood Peak - Risky User alert", "Tkt#254126 新加坡办公室2026",
                              "Tkt#254206 Amazon Quick"],
        "Need to follow up TODAY!": ["Tkt#254323 RE: ACTION REQUIRED - MFA"],
        "To Do": ["Update asset register", "Renew SSL cert for portal"],
        "Wait for reply": ["CIML IT: Diana Lee - WFH - Printer"],
        "Waiting for assign": ["Tkt#253211 account creation", "Lunar - Quickbooks server access",
                               "Tkt#253730 FW: Pending Document", "New Colleague email set up",
                               "Tkt#254108 Call regarding change"],
    }
    items = []
    i = 0
    for cat, subs in data.items():
        for s in subs:
            i += 1
            age = [0.2, 1, 2, 4, 6, 9, 12, 20][i % 8]
            items.append({"entry_id": f"demo-c{i}", "store_id": "", "folder": "\\\\Premier Servicedesk\\Inbox",
                          "subject": s, "categories": [cat], "cls": 48 if cat == "To Do" else 43,
                          "sender": "工作" if cat == "To Do" else ["Aico Li", "Thomas Chau", "Ka Yee Lam", "Diana Lee"][i % 4],
                          "received": iso(n - timedelta(days=age)), "unread": i % 5 == 0, "flagged": True,
                          "due": iso(n - timedelta(days=1)) if i % 6 == 0 else None, "complete": False,
                          "ticket_id": ticket_ref(s)})
    return {"master": [{"name": m, "color": OUTLOOK_COLORS[c]} for m, c in master], "items": items,
            "scanned_at": iso(n), "error": None, "folders": 3}


def open_ui(url):
    if os.name == "nt":
        for exe in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
            if os.path.exists(exe):
                try:
                    subprocess.Popen([exe, f"--app={url}", "--window-size=1480,940"])
                    return
                except Exception:
                    pass
    webbrowser.open(url)


def make_handler(app):
    ui_path = os.path.join(BASE_DIR, "ui.html")

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _local(self):
            host = (self.headers.get("Host") or "").rsplit(":", 1)[0]
            origin = self.headers.get("Origin")
            if host not in ("127.0.0.1", "localhost"):
                return False
            if origin and urlparse(origin).hostname not in ("127.0.0.1", "localhost"):
                return False
            return True

        def _send(self, code, data, ctype="application/json; charset=utf-8"):
            raw = data if isinstance(data, bytes) else json.dumps(data, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(raw)

        def do_GET(self):
            if not self._local():
                return self._send(403, {"error": "forbidden"})
            u = urlparse(self.path)
            try:
                if u.path in ("/", "/index.html"):
                    with open(ui_path, "rb") as f:
                        return self._send(200, f.read(), "text/html; charset=utf-8")
                if u.path == "/api/ping":
                    return self._send(200, {"app": APP, "version": VERSION})
                if u.path == "/api/state":
                    return self._send(200, app.state())
                if u.path == "/api/folders":
                    if app.demo or not HAS_COM:
                        K = "\\\\kwong@premier-technology.com"
                        demo_f = [("Inbox", "inbox", 0, 18115), ("Need to Follow up", "sub", 1, 2),
                                  ("Assigned need to follow up", "sub", 1, 1), ("Wait for arrange", "sub", 1, 0),
                                  ("Alerts", "sub", 1, 75919), ("Mailbox Archive", "sub", 1, 18878),
                                  ("Sent Items", "sent", 0, 9210), ("Tasks", "task", 0, 14)]
                        fl = []
                        for nm, kind, depth, cnt in demo_f:
                            path = K + "\\" + ("Inbox\\" + nm if kind == "sub" else nm)
                            fl.append({"path": path, "name": nm, "store": "kwong@premier-technology.com", "kind": kind,
                                       "default": True, "depth": depth, "count": cnt})
                        fl += [{"path": "\\\\PTClientforHD\\Inbox", "name": "Inbox", "store": "PTClientforHD", "kind": "inbox",
                                "default": False, "depth": 0, "count": 3120},
                               {"path": "\\\\PTClientforHD\\Sent Items", "name": "Sent Items", "store": "PTClientforHD",
                                "kind": "sent", "default": False, "depth": 0, "count": 880}]
                        return self._send(200, {"folders": fl})
                    return self._send(200, {"folders": list_folders()})
                if u.path == "/api/diagnostics":
                    qs = parse_qs(u.query)
                    days = max(1, min(14, int((qs.get("days") or ["3"])[0])))
                    data = diagnostics_csv(app, days, (qs.get("mask_email") or ["1"])[0] == "1",
                                           (qs.get("mask_subject") or ["0"])[0] == "1")
                    self.send_response(200)
                    self.send_header("Content-Type", "text/csv; charset=utf-8")
                    self.send_header("Content-Disposition",
                                     f'attachment; filename="PulseDesk_diagnostics_{now():%Y%m%d_%H%M}.csv"')
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                    return
                if u.path == "/api/timeline":
                    tid = parse_qs(u.query).get("id", [""])[0]
                    return self._send(200, app.db.timeline(tid))
                return self._send(404, {"error": "not found"})
            except Exception as e:
                traceback.print_exc()
                return self._send(500, {"error": str(e)})

        def do_POST(self):
            if not self._local():
                return self._send(403, {"error": "forbidden"})
            try:
                n = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(n) or b"{}")
                p = urlparse(self.path).path
                if p == "/api/ticket":
                    return self._send(200, app.ticket_action(str(body.get("id")), body))
                if p == "/api/bulk":
                    return self._send(200, app.bulk_action(list(body.get("ids") or []), body))
                if p == "/api/sync":
                    busy = ""
                    if app.scanner:
                        h = app.scanner.health
                        busy = h.get("scan_kind", "") if h.get("running") else ""
                        full = bool(body.get("full"))
                        app.scanner.rescan(full=full, cats=bool(body.get("cats")), manual=not full)
                    return self._send(200, {"ok": True, "busy": busy,
                                            "scans": app.scanner.health.get("scans", 0) if app.scanner else 0})
                if p == "/api/open_mail":
                    if app.demo or not HAS_COM:
                        raise ValueError("示範模式不能打開 Outlook 郵件")
                    open_mail(body["entry_id"], body.get("store_id"))
                    return self._send(200, {"ok": True})
                if p == "/api/notify":
                    if app.demo or not HAS_COM:
                        raise ValueError("示範模式不能建立 Outlook 草稿")
                    tk = app.db.get(str(body.get("id")))
                    col = next((c for c in app.cfg["colleagues"] if c["name"] == body.get("colleague")), None)
                    if not tk or not col:
                        raise ValueError("找不到工單或同事")
                    draft_to_colleague(app, tk, col)
                    app.db.log(tk["id"], f"已開啟 Outlook 草稿通知 {col['name']}")
                    return self._send(200, {"ok": True})
                if p == "/api/test_signature":
                    class _M:
                        pass
                    m = _M()
                    m.sender = str(body.get("sender") or (app.cfg["support_addresses"] or [""])[0]).strip().lower()
                    m.sender_name = str(body.get("sender_name") or "")
                    text = str(body.get("text") or "")
                    m.body = lambda: text
                    who, how = identify_replier(m, app.cfg)
                    return self._send(200, {"replier": who, "how": how, "reply_lines": len([l for l in reply_part(text) if l.strip()])})
                if p == "/api/dismiss_mail":
                    app.db.x("UPDATE mails SET dismissed=1 WHERE entry_id=?", (body.get("entry_id"),))
                    return self._send(200, {"ok": True})
                if p == "/api/config":
                    old = app.cfg
                    app.cfg = sanitize_config(body, old)
                    if not app.demo:
                        save_config(app.cfg)
                    if app.scanner:
                        changed = lambda keys: any(old.get(k) != app.cfg.get(k) for k in keys)
                        if changed(("internal_domains", "support_addresses")):
                            app.scanner.rescan(full=True)          # 判斷規則改變：重新判斷最近的郵件
                        elif changed(("include_subfolders", "scan_folders", "extra_folders", "days_back", "scan_sent_items")):
                            app.scanner.rescan(folders=True)       # 只重新計算資料夾，新資料夾會自動補掃
                        if changed(("category_folders",)):
                            app.scanner.rescan(cat_folders=True)
                        elif changed(("hidden_categories",)):
                            app.scanner.rescan(cats=True)
                    return self._send(200, {"ok": True, "config": app.cfg})
                if p == "/api/shutdown":
                    self._send(200, {"ok": True})
                    app.stopping = True
                    threading.Thread(target=app.server.shutdown, daemon=True).start()
                    return
                return self._send(404, {"error": "not found"})
            except Exception as e:
                return self._send(400, {"error": str(e)})

    return H


class Server(ThreadingHTTPServer):
    """Windows 上不允許兩個程式同時使用同一個連接埠（避免同時運行兩個 PulseDesk）。"""
    allow_reuse_address = os.name != "nt"
    daemon_threads = True

    def server_bind(self):
        if os.name == "nt":
            import socket
            if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
                self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def ping(url):
    try:
        return json.loads(urlopen(url + "api/ping", timeout=2).read())
    except Exception:
        return None


def request_shutdown(url):
    from urllib.request import Request
    try:
        urlopen(Request(url + "api/shutdown", data=b"{}", headers={"Content-Type": "application/json"}), timeout=3).read()
    except Exception:
        pass


def take_over(url, log=print):
    """如果已有 PulseDesk 在運行：同一版本 → 回傳 "same"；其他版本 → 關閉舊版本後回傳 "replaced"。"""
    info = ping(url)
    if not info or info.get("app") != APP:
        return None
    if info.get("version") == VERSION:
        return "same"
    for _ in range(6):                      # 舊版本可能同時運行了幾個
        log(f"關閉舊版本 PulseDesk {info.get('version')} …")
        request_shutdown(url)
        for _ in range(20):
            time.sleep(0.5)
            info = ping(url)
            if not info:
                break
        if not info or info.get("app") != APP:
            return "replaced"
        if info.get("version") == VERSION:
            return "same"
    return "replaced"


def main():
    # 以 --windowed 打包成 EXE 時沒有主控台，把輸出寫到記錄檔
    if sys.stdout is None or sys.stderr is None:
        os.makedirs(DATA_DIR, exist_ok=True)
        log = open(os.path.join(DATA_DIR, "pulsedesk.log"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or log
        sys.stderr = sys.stderr or log
    ap = argparse.ArgumentParser()
    ap.add_argument("--demo", action="store_true", help="用示範資料預覽界面")
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()

    port = load_config()["port"] + (1 if args.demo else 0)
    url = f"http://127.0.0.1:{port}/"
    state = take_over(url)
    if state == "same":                     # 已經在運行：只打開視窗
        if not args.no_browser:
            open_ui(url)
        return

    server = None
    for _ in range(40):                     # 等舊版本釋放連接埠
        try:
            server = Server(("127.0.0.1", port), None)
            break
        except OSError:
            if take_over(url) == "same":
                if not args.no_browser:
                    open_ui(url)
                return
            time.sleep(0.5)
    if server is None:
        raise SystemExit(f"連接埠 {port} 被其他程式佔用，請重新開機後再試。")

    app = App(demo=args.demo)
    app.url = url
    server.RequestHandlerClass = make_handler(app)
    app.server = server
    if app.scanner:
        app.scanner.start()
    print(f"{APP} {VERSION} running at {app.url}" + ("  [DEMO]" if app.demo else "")
          + ("（已取代舊版本）" if state == "replaced" else ""))
    if not args.no_browser:
        threading.Timer(0.8, open_ui, args=(app.url,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    app.stopping = True
    server.server_close()


if __name__ == "__main__":
    main()
