"""
Outlook MCP Server für Claude Desktop.

Greift über das COM-Objektmodell auf das klassische Outlook zu und
stellt Werkzeuge zum Lesen und Durchsuchen von E-Mails,
Kalendereinträgen und Anhängen bereit. Schreibend werden nur
E-Mail-Entwürfe angelegt (neu, Antwort, Weiterleitung) und
ungesendete Entwürfe im Ordner Entwürfe bearbeitet. Besprechungs-
einladungen werden vorbefüllt in einem Outlook-Fenster geöffnet,
aber nicht gespeichert. Es wird nichts
gesendet, verschoben oder gelöscht; dafür gibt es bewusst kein
Werkzeug. Anhänge werden zum Lesen kurz in ein temporäres
Verzeichnis gespeichert und danach wieder gelöscht.

Voraussetzungen: Windows, klassisches Outlook (das neue Outlook
unterstützt kein COM), Python-Pakete "mcp", "pywin32", "pypdf",
"openpyxl", "olefile", "xlrd" und "pillow".
"""

from __future__ import annotations

import asyncio
import email
import html
import io
import locale
import logging
import math
import os
import re
import shutil
import struct
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from email import policy
from itertools import groupby
from typing import Any, Callable

import pythoncom
import pywintypes
import win32com.client
from mcp.server.fastmcp import FastMCP, Image

from . import __version__

# Protokoll ausschliesslich auf stderr: stdout gehört dem MCP-Protokoll.
logging.basicConfig(
    stream=sys.stderr,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("outlook-mcp")


# ============================================================
# Konfiguration (über Umgebungsvariablen anpassbar)
# ============================================================

def _env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except ValueError:
        return default


MAX_RESULTS = _env_int("OUTLOOK_MCP_MAX_RESULTS", 50)
MAX_BODY_CHARS = _env_int("OUTLOOK_MCP_MAX_BODY_CHARS", 8000)
MAX_SCAN = _env_int("OUTLOOK_MCP_MAX_SCAN", 5000)
MAX_CALENDAR_DAYS = _env_int("OUTLOOK_MCP_MAX_CALENDAR_DAYS", 366)
MAX_FREEBUSY_DAYS = _env_int("OUTLOOK_MCP_MAX_FREEBUSY_DAYS", 62)
PREVIEW_CHARS = _env_int("OUTLOOK_MCP_PREVIEW_CHARS", 200)
MAX_ATTACHMENT_BYTES = _env_int("OUTLOOK_MCP_MAX_ATTACHMENT_BYTES", 30_000_000)
# Textmenge pro Aufruf; längere Anhänge werden mit 'offset' weitergelesen.
MAX_ATTACHMENT_CHARS = _env_int("OUTLOOK_MCP_MAX_ATTACHMENT_CHARS", 30000)
# Die Claude-API nimmt Bilder bis 5 MB an, base64 macht sie ein Drittel
# grösser. Grössere Bilder werden verkleinert; Claude skaliert Bilder mit
# mehr als ca. 1568 Pixel Kantenlänge ohnehin herunter.
MAX_IMAGE_BYTES = _env_int("OUTLOOK_MCP_MAX_IMAGE_BYTES", 3_500_000)
MAX_IMAGE_EDGE = _env_int("OUTLOOK_MCP_MAX_IMAGE_EDGE", 1568)
# Schrift für den Text von Entwürfen, direkt am Absatz gesetzt, damit
# sie in jedem Outlook gleich erscheint. Leere Schriftart = Formatvorlage
# der Mail (bei Antworten oft die Schrift der Originalmail).
DRAFT_FONT = os.environ.get("OUTLOOK_MCP_DRAFT_FONT", "Aptos").strip().strip('"')
DRAFT_FONT_SIZE = _env_int("OUTLOOK_MCP_DRAFT_FONT_SIZE", 12)
# Datumsformat für Kalenderfilter. Leer = kurzes Datumsformat von Windows.
# Beispiel bei Problemen: "%d.%m.%Y %H:%M"
CALENDAR_DATE_FORMAT = os.environ.get("OUTLOOK_MCP_DATE_FORMAT", "")


# ============================================================
# Outlook-Konstanten
# ============================================================

DEFAULT_FOLDERS = {
    "posteingang": 6, "inbox": 6,
    "gesendet": 5, "gesendete elemente": 5, "sent": 5,
    "entwürfe": 16, "drafts": 16,
    "gelöscht": 3, "gelöschte elemente": 3, "deleted": 3,
    "junk": 23, "junk-e-mail": 23,
    "kalender": 9, "calendar": 9,
}
FOLDER_INBOX = 6
FOLDER_CALENDAR = 9
ITEM_TYPE_MAIL = 0
ITEM_TYPE_APPOINTMENT = 1
CLASS_MAIL = 43
FORMAT_HTML = 2
CLOSE_DISCARD = 1
RECIPIENT_TO, RECIPIENT_CC, RECIPIENT_BCC = 1, 2, 3
RECIPIENT_REQUIRED, RECIPIENT_OPTIONAL, RECIPIENT_RESOURCE = 1, 2, 3
MEETING_ORGANIZER = 1

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_LAST_VERB = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
PR_LAST_VERB_TIME = "http://schemas.microsoft.com/mapi/proptag/0x10820040"
PR_ROOM_CAPACITY = "http://schemas.microsoft.com/mapi/proptag/0x08070003"
# PR_TODO_ITEM_FLAGS: Exchange filtert zuverlässig darauf, nicht auf
# PR_FLAG_STATUS. Erledigte Kennzeichnungen werden in Python aussortiert.
DASL_TODO_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003"
TEAMS_URL_PROP = ("http://schemas.microsoft.com/mapi/string/"
                  "{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl")

ATT_BY_VALUE = 1
ATT_BY_REFERENCE = 4
ATT_EMBEDDED = 5
ATT_OLE = 6
ATTACHMENT_TYPE = {ATT_BY_VALUE: "Datei", ATT_BY_REFERENCE: "Verknüpfung",
                   ATT_EMBEDDED: "Outlook-Element", ATT_OLE: "OLE-Objekt"}

IMPORTANCE = {0: "niedrig", 1: "normal", 2: "hoch"}
BUSY_STATUS = {0: "frei", 1: "mit Vorbehalt", 2: "beschäftigt",
               3: "abwesend", 4: "anderswo tätig"}
MEETING_STATUS = {0: "Termin", 1: "Besprechung (Organisator)",
                  3: "Besprechung (Einladung erhalten)", 5: "abgesagt",
                  7: "abgesagt (Absage erhalten)"}
RESPONSE_STATUS = {0: "keine", 1: "Organisator", 2: "mit Vorbehalt",
                   3: "zugesagt", 4: "abgelehnt", 5: "noch nicht beantwortet"}
RECIPIENT_TYPE = {1: "erforderlich", 2: "optional", 3: "Ressource"}
IMPORTANCE_VALUES = {"niedrig": 0, "low": 0, "normal": 1, "hoch": 2, "high": 2}
SENSITIVITY = {0: "normal", 1: "persönlich", 2: "privat", 3: "vertraulich"}
SENSITIVITY_PRIVATE = 2
FLAG_STATUS = {0: "keine", 1: "erledigt", 2: "gekennzeichnet"}
FLAG_MARKED = 2
# PR_LAST_VERB_EXECUTED: letzte Aktion auf einer Mail.
LAST_VERB = {102: "beantwortet", 103: "allen geantwortet", 104: "weitergeleitet"}
RESPONSE_NOT_RESPONDED = 5
MEETING_RECEIVED = 3
RECURRENCE_TYPE = {0: "täglich", 1: "wöchentlich", 2: "monatlich",
                   3: "monatlich (n-ter Wochentag)", 5: "jährlich",
                   6: "jährlich (n-ter Wochentag)"}
WEEKDAY_MASK = ((2, "Montag"), (4, "Dienstag"), (8, "Mittwoch"),
                (16, "Donnerstag"), (32, "Freitag"), (64, "Samstag"),
                (1, "Sonntag"))
INSTANCE = {1: "ersten", 2: "zweiten", 3: "dritten", 4: "vierten", 5: "letzten"}
MONTHS = ("Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
          "August", "September", "Oktober", "November", "Dezember")
# Name der Exchange-Adressliste mit allen Räumen, je nach Sprache.
ROOM_LIST_NAMES = {"alle räume", "all rooms", "toutes les salles",
                   "tutte le sale", "salles"}
MAX_EXCEPTIONS = 100
MAX_LINKS = 50

# Zeichen in Recipient.FreeBusy(..., CompleteFormat=True). Für
# "anderswo tätig" liefert Outlook in der Praxis 5 statt 4.
FREEBUSY_STATUS = {"0": "frei", "1": "mit Vorbehalt", "2": "beschäftigt",
                   "3": "abwesend", "4": "anderswo tätig", "5": "anderswo tätig"}
FREEBUSY_INTERVALS = (5, 10, 15, 30, 60)
# Recipient.FreeBusy liefert pro Aufruf 29 Tage ab dem Startdatum.
FREEBUSY_DAYS_PER_CALL = 28


# ============================================================
# COM-Zugriff in einem eigenen Thread
# COM-Objekte sind an den Thread gebunden, in dem sie erzeugt
# wurden. Alle Outlook-Aufrufe laufen daher in genau einem
# Worker-Thread mit initialisiertem COM.
# ============================================================

_executor = ThreadPoolExecutor(
    max_workers=1,
    thread_name_prefix="outlook-com",
    initializer=pythoncom.CoInitialize,
)
_state: dict[str, Any] = {}


def _namespace():
    ns = _state.get("ns")
    if ns is None:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        _state["ns"] = ns
        log.info("Verbindung zu Outlook hergestellt")
    return ns


def _call(fn: Callable, *args):
    try:
        return fn(*args)
    except pywintypes.com_error as exc:
        # Outlook wurde evtl. neu gestartet: einmal neu verbinden.
        log.warning("COM-Fehler, neuer Verbindungsversuch: %s", exc)
        _state.clear()
        try:
            return fn(*args)
        except pywintypes.com_error as exc2:
            raise RuntimeError(f"Outlook-Zugriff fehlgeschlagen: {exc2}") from exc2


async def _run(fn: Callable, *args):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(_executor, _call, fn, *args)


# ============================================================
# Hilfsfunktionen
# ============================================================

def _safe(obj, attr: str, default=None):
    try:
        value = getattr(obj, attr)
        return default if value is None else value
    except Exception:
        return default


def _dt(value) -> datetime | None:
    """COM-Datum in ein naives lokales datetime umwandeln."""
    if value is None:
        return None
    try:
        dt = datetime(value.year, value.month, value.day,
                      value.hour, value.minute, value.second)
    except Exception:
        return None
    # Outlook verwendet 01.01.4501 als "kein Datum".
    return None if dt.year >= 4500 else dt


def _fmt(dt: datetime | None) -> str | None:
    return dt.strftime("%Y-%m-%d %H:%M") if dt else None


def _parse_date(text: str, name: str) -> datetime | None:
    if not text or not text.strip():
        return None
    text = text.strip()
    for pattern in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%d.%m.%Y %H:%M", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            continue
    raise ValueError(
        f"Ungültiges Datum für '{name}': '{text}'. "
        "Erlaubt: JJJJ-MM-TT, JJJJ-MM-TT HH:MM, TT.MM.JJJJ"
    )


def _has_time(text: str) -> bool:
    return ":" in (text or "")


def _period(start_date: str, end_date: str, max_days: int) -> tuple[datetime, datetime]:
    """Zeitraum aus start_date (Standard heute) und end_date (inklusive,
    Standard start_date + 7 Tage)."""
    start = _parse_date(start_date, "start_date")
    if start is None:
        start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end = _parse_date(end_date, "end_date")
    if end is None:
        end = start + timedelta(days=7)
    elif not _has_time(end_date):
        end += timedelta(days=1)  # Enddatum inklusive
    if end <= start:
        raise ValueError("end_date muss nach start_date liegen.")
    if (end - start).days > max_days:
        raise ValueError(f"Zeitraum zu gross (max. {max_days} Tage).")
    return start, end


def _truncate(text: str, limit: int) -> str:
    text = (text or "").replace("\r\n", "\n").strip()
    if len(text) > limit:
        rest = len(text) - limit
        return text[:limit] + f"\n[… gekürzt, {rest} weitere Zeichen]"
    return text


def _limit(value: int) -> int:
    return max(1, min(int(value or MAX_RESULTS), MAX_RESULTS))


def _dasl_text(value: str) -> str:
    # Hochkommas verdoppeln, LIKE-Platzhalter entfernen.
    return value.replace("'", "''").replace("%", "")


def _folder_path(folder) -> str:
    return _safe(folder, "FolderPath", "") or _safe(folder, "Name", "")


def _resolve_folder(path: str, default_id: int):
    """Ordner über Kurzname (z. B. 'Posteingang') oder vollständigen
    Pfad (z. B. '\\\\max@firma.ch\\Posteingang\\Projekte') auflösen."""
    ns = _namespace()
    if not path or not path.strip():
        return ns.GetDefaultFolder(default_id)

    key = path.strip().lower()
    if key in DEFAULT_FOLDERS:
        return ns.GetDefaultFolder(DEFAULT_FOLDERS[key])

    parts = [p for p in path.strip().strip("\\").split("\\") if p]
    current = None
    collection = ns.Folders
    for part in parts:
        match = None
        for candidate in collection:
            if (_safe(candidate, "Name", "") or "").lower() == part.lower():
                match = candidate
                break
        if match is None:
            raise ValueError(
                f"Ordner nicht gefunden: '{part}' (Pfad '{path}'). "
                "Gültige Pfade liefern list_folders bzw. list_calendars."
            )
        current = match
        collection = match.Folders
    if current is None:
        raise ValueError(f"Ungültiger Ordnerpfad: '{path}'")
    return current


def _shared_calendar(mailbox: str):
    ns = _namespace()
    recipient = ns.CreateRecipient(mailbox)
    recipient.Resolve()
    if not recipient.Resolved:
        raise ValueError(f"Postfach konnte nicht aufgelöst werden: '{mailbox}'")
    try:
        return ns.GetSharedDefaultFolder(recipient, FOLDER_CALENDAR)
    except pywintypes.com_error as exc:
        raise ValueError(
            f"Kein Zugriff auf den Kalender von '{mailbox}'. "
            "Fehlt die Freigabe (mindestens Leserechte)? "
            "Belegte Zeiten liefert get_free_busy."
        ) from exc


def _collect_folders(root, item_type: int, max_depth: int) -> list:
    result = []

    def walk(folder, depth):
        try:
            if _safe(folder, "DefaultItemType") == item_type:
                result.append(folder)
            if depth < max_depth:
                for sub in folder.Folders:
                    walk(sub, depth + 1)
        except pywintypes.com_error:
            pass  # z. B. Ordner ohne Leserechte

    walk(root, 0)
    return result


def _smtp_of_sender(item) -> str:
    address = _safe(item, "SenderEmailAddress", "") or ""
    if _safe(item, "SenderEmailType", "") == "EX":
        try:
            address = item.PropertyAccessor.GetProperty(PR_SENDER_SMTP) or address
        except Exception:
            pass
    return address


def _attachment_count(item) -> int:
    try:
        return int(item.Attachments.Count)
    except Exception:
        return 0


def _open_item(entry_id: str, store_id: str):
    ns = _namespace()
    return (ns.GetItemFromID(entry_id, store_id) if store_id
            else ns.GetItemFromID(entry_id))


def _attachment_name(att) -> str:
    return _safe(att, "FileName", "") or _safe(att, "DisplayName", "")


def _attachment_info(att, nr: int, html_body: str) -> dict:
    info = {
        "nr": nr,
        "name": _attachment_name(att),
        "groesse_bytes": _safe(att, "Size"),
        "typ": ATTACHMENT_TYPE.get(_safe(att, "Type"), ""),
    }
    content_id = ""
    try:
        content_id = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID) or ""
    except Exception:
        pass
    if content_id and f"cid:{content_id}" in html_body:
        info["im_text_eingebettet"] = True  # z. B. Logo in der Signatur
    return info


def _attachment_list(item) -> list[dict]:
    html_body = _safe(item, "HTMLBody", "") or ""
    result = []
    try:
        attachments = item.Attachments
        for nr in range(1, int(attachments.Count) + 1):
            result.append(_attachment_info(attachments.Item(nr), nr, html_body))
    except Exception:
        pass
    return result


# ============================================================
# E-Mail
# ============================================================

def _mail_summary(item, folder) -> dict:
    summary = {
        "entry_id": _safe(item, "EntryID", ""),
        "store_id": _safe(folder, "StoreID", ""),
        "betreff": _safe(item, "Subject", ""),
        "von": _safe(item, "SenderName", ""),
        "von_email": _smtp_of_sender(item),
        "empfangen": _fmt(_dt(_safe(item, "ReceivedTime"))),
        "ungelesen": bool(_safe(item, "UnRead", False)),
        "anhaenge": _attachment_count(item),
        "ordner": _folder_path(folder),
        "vorschau": _truncate(_safe(item, "Body", ""), PREVIEW_CHARS),
    }
    # Nur Abweichungen vom Normalfall ausgeben, damit Trefferlisten kurz bleiben.
    flag = _safe(item, "FlagStatus", 0)
    if flag:
        summary["kennzeichnung"] = FLAG_STATUS.get(flag, "")
    importance = _safe(item, "Importance", 1)
    if importance != 1:
        summary["wichtigkeit"] = IMPORTANCE.get(importance, "")
    categories = _safe(item, "Categories", "")
    if categories:
        summary["kategorien"] = categories
    return summary


def _mail_matches(item, category: str, recipient: str, attachment_name: str,
                  flagged: bool = False) -> bool:
    """Filter, die DASL nicht zuverlässig abbildet, in Python prüfen."""
    if flagged and _safe(item, "FlagStatus", 0) != FLAG_MARKED:
        return False
    if category and category not in (_safe(item, "Categories", "") or "").lower():
        return False
    if attachment_name:
        names = []
        try:
            attachments = item.Attachments
            for nr in range(1, int(attachments.Count) + 1):
                names.append(_attachment_name(attachments.Item(nr)).lower())
        except Exception:
            pass
        if not any(attachment_name in n for n in names):
            return False
    if recipient:
        display = " ".join((_safe(item, "To", "") or "", _safe(item, "CC", "") or "")).lower()
        if recipient in display:
            return True
        try:
            for r in item.Recipients:
                if _safe(r, "Type") not in (RECIPIENT_TO, RECIPIENT_CC):
                    continue
                if recipient in _recipient_smtp(r).lower():
                    return True
        except Exception:
            pass
        return False
    return True


def _recipient_smtp(recipient) -> str:
    try:
        return recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) or ""
    except Exception:
        return _safe(recipient, "Address", "") or ""


def _list_folders_impl(max_depth: int) -> dict:
    ns = _namespace()
    folders = []
    for store_root in ns.Folders:
        for folder in _collect_folders(store_root, ITEM_TYPE_MAIL, max_depth):
            folders.append({
                "pfad": _folder_path(folder),
                "name": _safe(folder, "Name", ""),
                "elemente": _safe(_safe(folder, "Items"), "Count"),
                "ungelesen": _safe(folder, "UnReadItemCount"),
            })
    return {"anzahl": len(folders), "ordner": folders}


def _mail_dasl(query, search_body, use_index, sender, unread_only,
               has_attachments, flagged, importance) -> str | None:
    conditions = []
    if query and query.strip():
        q = _dasl_text(query.strip())
        fields = ["urn:schemas:httpmail:subject"]
        if search_body:
            fields.append("urn:schemas:httpmail:textdescription")
        # ci_phrasematch nutzt den Suchindex von Outlook (Windows Search
        # bzw. Exchange): schneller, findet aber nur ganze Wörter.
        parts = [(f"\"{f}\" ci_phrasematch '{q}'" if use_index
                  else f"\"{f}\" LIKE '%{q}%'") for f in fields]
        conditions.append("(" + " OR ".join(parts) + ")")
    if sender and sender.strip():
        s = _dasl_text(sender.strip())
        conditions.append(
            f"(\"urn:schemas:httpmail:fromname\" LIKE '%{s}%' OR "
            f"\"urn:schemas:httpmail:fromemail\" LIKE '%{s}%')"
        )
    if unread_only:
        conditions.append("\"urn:schemas:httpmail:read\" = 0")
    if has_attachments:
        conditions.append("\"urn:schemas:httpmail:hasattachment\" = 1")
    if flagged:
        conditions.append(f"\"{DASL_TODO_FLAGS}\" = 1")
    if importance is not None:
        conditions.append(f"\"urn:schemas:httpmail:importance\" = {importance}")
    return "@SQL=" + " AND ".join(conditions) if conditions else None


def _search_mails_impl(query, folder, include_subfolders, sender, date_from,
                       date_to, unread_only, has_attachments, search_body,
                       limit, flagged=False, category="", importance="",
                       recipient="", attachment_name="", use_index=False) -> dict:
    limit = _limit(limit)
    dt_from = _parse_date(date_from, "date_from")
    dt_to = _parse_date(date_to, "date_to")
    if dt_to and not _has_time(date_to):
        dt_to += timedelta(days=1)  # Enddatum inklusive

    importance_value = None
    if importance and importance.strip():
        key = importance.strip().lower()
        if key not in IMPORTANCE_VALUES:
            raise ValueError("importance muss 'niedrig', 'normal' oder 'hoch' sein.")
        importance_value = IMPORTANCE_VALUES[key]
    category = (category or "").strip().lower()
    recipient = (recipient or "").strip().lower()
    attachment_name = (attachment_name or "").strip().lower()

    root = _resolve_folder(folder, FOLDER_INBOX)
    folders = (_collect_folders(root, ITEM_TYPE_MAIL, 10)
               if include_subfolders else [root])

    dasl_args = (query, search_body, use_index, sender, unread_only,
                 has_attachments or bool(attachment_name), flagged, importance_value)
    dasl = _mail_dasl(*dasl_args)
    index_fallback = False

    hits: list[dict] = []
    scanned = 0
    truncated = False
    for current in folders:
        try:
            items = current.Items
            if dasl:
                try:
                    items = items.Restrict(dasl)
                except pywintypes.com_error:
                    if not use_index:
                        raise
                    # Ordner ohne Suchindex: normale Suche verwenden.
                    index_fallback = True
                    items = current.Items.Restrict(
                        _mail_dasl(query, search_body, False, *dasl_args[3:]))
            items.Sort("[ReceivedTime]", True)  # neueste zuerst
        except pywintypes.com_error as exc:
            log.warning("Ordner übersprungen (%s): %s", _folder_path(current), exc)
            continue

        found_here = 0
        item = items.GetFirst()
        while item is not None:
            if scanned >= MAX_SCAN:
                truncated = True
                break
            scanned += 1
            received = _dt(_safe(item, "ReceivedTime"))
            if received and dt_from and received < dt_from:
                break  # sortiert: ältere folgen nur noch
            if (not (received and dt_to and received >= dt_to)
                    and _mail_matches(item, category, recipient, attachment_name,
                                      flagged)):
                hits.append(_mail_summary(item, current))
                found_here += 1
                if found_here >= limit:
                    break
            item = items.GetNext()
        if truncated:
            break

    hits.sort(key=lambda h: h["empfangen"] or "", reverse=True)
    result = {"anzahl": min(len(hits), limit), "mails": hits[:limit]}
    if truncated or len(hits) > limit:
        result["hinweis"] = ("Weitere Treffer möglich. Suche eingrenzen "
                             "(Zeitraum, Ordner, Absender).")
    if index_fallback:
        result["hinweis_index"] = ("Für mindestens einen Ordner ist kein Suchindex "
                                   "verfügbar, dort wurde normal gesucht.")
    return result


def _last_action(item) -> dict | None:
    try:
        verb = item.PropertyAccessor.GetProperty(PR_LAST_VERB)
    except Exception:
        return None
    if verb not in LAST_VERB:
        return None
    action = {"aktion": LAST_VERB[verb]}
    try:
        # Der Wert ist UTC; pywin32 liefert ihn als zeitzonenbehaftetes Datum.
        when = item.PropertyAccessor.GetProperty(PR_LAST_VERB_TIME)
        if getattr(when, "tzinfo", None) is not None:
            when = when.astimezone()
        action["am"] = _fmt(_dt(when))
    except Exception:
        pass
    return action


def _flag_info(item) -> dict | None:
    status = _safe(item, "FlagStatus", 0)
    if not status:
        return None
    info = {"status": FLAG_STATUS.get(status, ""),
            "text": _safe(item, "FlagRequest", "")}
    for key, attr in (("beginn", "TaskStartDate"), ("faellig", "TaskDueDate"),
                      ("erledigt_am", "TaskCompletedDate")):
        value = _fmt(_dt(_safe(item, attr)))
        if value:
            info[key] = value[:10]
    return info


def _bcc_of(item) -> str:
    names = []
    try:
        for r in item.Recipients:
            if _safe(r, "Type") == RECIPIENT_BCC:
                names.append(_safe(r, "Name", "") or _recipient_smtp(r))
    except Exception:
        pass
    return "; ".join(n for n in names if n)


def _links(html_body: str) -> list[dict]:
    result, seen = [], set()
    for match in re.finditer(r"(?is)<a\b[^>]*?href\s*=\s*([\"'])(.*?)\1[^>]*>(.*?)</a\s*>",
                             html_body or ""):
        url = html.unescape(match.group(2)).strip()
        if not url or url.lower().startswith(("cid:", "#", "javascript:")) or url in seen:
            continue
        seen.add(url)
        text = _html_to_text(match.group(3)).replace("\n", " ").strip()
        result.append({"text": text[:200], "url": url})
        if len(result) >= MAX_LINKS:
            break
    return result


def _auth_results(headers: str) -> dict:
    """SPF, DKIM und DMARC aus Authentication-Results auslesen."""
    unfolded = re.sub(r"\r?\n[ \t]+", " ", headers or "")
    result = {}
    for line in re.findall(r"(?im)^Authentication-Results:(.*)$", unfolded):
        for key in ("spf", "dkim", "dmarc", "compauth"):
            match = re.search(rf"\b{key}=(\w+)", line, re.I)
            if match and key not in result:
                result[key] = match.group(1).lower()
    return result


def _get_mail_impl(entry_id, store_id, max_body_chars, body_format="text",
                   include_headers=False) -> dict:
    fmt = (body_format or "text").strip().lower()
    if fmt not in ("text", "html"):
        raise ValueError("body_format muss 'text' oder 'html' sein.")
    item = _open_item(entry_id, store_id)
    attachments = _attachment_list(item)
    limit = max(500, min(int(max_body_chars or MAX_BODY_CHARS), MAX_BODY_CHARS))
    html_body = _safe(item, "HTMLBody", "") or ""
    data = {
        "entry_id": _safe(item, "EntryID", ""),
        "betreff": _safe(item, "Subject", ""),
        "von": _safe(item, "SenderName", ""),
        "von_email": _smtp_of_sender(item),
        "an": _safe(item, "To", ""),
        "cc": _safe(item, "CC", ""),
        "bcc": _bcc_of(item),
        "empfangen": _fmt(_dt(_safe(item, "ReceivedTime"))),
        "gesendet": _fmt(_dt(_safe(item, "SentOn"))),
        "wichtigkeit": IMPORTANCE.get(_safe(item, "Importance"), "normal"),
        "vertraulichkeit": SENSITIVITY.get(_safe(item, "Sensitivity"), "normal"),
        "lesebestaetigung_angefordert": bool(_safe(item, "ReadReceiptRequested", False)),
        "zustellbestaetigung_angefordert": bool(
            _safe(item, "OriginatorDeliveryReportRequested", False)),
        "kategorien": _safe(item, "Categories", ""),
        "kennzeichnung": _flag_info(item),
        "letzte_aktion": _last_action(item),
        "konversation": _safe(item, "ConversationTopic", ""),
        "ordner": _folder_path(_safe(item, "Parent")),
        "anhaenge": attachments,
    }
    if fmt == "html" and html_body:
        data["html"] = _truncate(html_body, limit)
    else:
        data["text"] = _truncate(_safe(item, "Body", ""), limit)
        links = _links(html_body)
        if links:
            data["links"] = links
    if include_headers:
        try:
            headers = item.PropertyAccessor.GetProperty(PR_TRANSPORT_HEADERS) or ""
        except Exception:
            headers = ""
        if headers:
            data["authentifizierung"] = _auth_results(headers)
            data["header"] = _truncate(headers, MAX_BODY_CHARS)
        else:
            data["header"] = None
            data["hinweis_header"] = ("Keine Internet-Header vorhanden (z. B. "
                                      "Entwurf, gesendete oder interne Mail).")
    return {k: v for k, v in data.items() if v not in (None, "") or k in
            ("an", "cc", "text", "header")}


def _conversation_items(item) -> list[tuple[Any, int]]:
    """Alle Elemente eines Gesprächs mit Tiefe in der Antwortkette."""
    conversation = item.GetConversation()
    if conversation is None:
        return []
    result = []

    def walk(node, depth):
        if len(result) >= MAX_SCAN:
            return
        result.append((node, depth))
        try:
            for child in conversation.GetChildren(node):
                walk(child, depth + 1)
        except pywintypes.com_error:
            pass

    for root in conversation.GetRootItems():
        walk(root, 0)
    return result


def _get_conversation_impl(entry_id, store_id, limit) -> dict:
    limit = max(1, min(int(limit or MAX_RESULTS), 200))
    item = _open_item(entry_id, store_id)
    topic = _safe(item, "ConversationTopic", "") or _safe(item, "Subject", "")
    try:
        nodes = _conversation_items(item)
        method = "Gespräch"
    except pywintypes.com_error:
        nodes = []
    if not nodes:
        # Postfach ohne Gesprächsfunktion: über das Gesprächsthema in
        # Posteingang und Gesendete Elemente suchen.
        method = "Gesprächsthema"
        ns = _namespace()
        flt = ("@SQL=\"urn:schemas:httpmail:thread-topic\" = "
               f"'{_dasl_text(topic)}'")
        for folder_id in (FOLDER_INBOX, 5):
            items = ns.GetDefaultFolder(folder_id).Items.Restrict(flt)
            for found in items:
                nodes.append((found, None))

    mails = []
    for node, depth in nodes:
        if _safe(node, "Class") != CLASS_MAIL:
            continue  # z. B. Besprechungsanfragen, Lesebestätigungen
        folder = _safe(node, "Parent")
        summary = _mail_summary(node, folder)
        summary["gesendet"] = _fmt(_dt(_safe(node, "SentOn")))
        if depth is not None:
            summary["ebene"] = depth
        mails.append(summary)
    mails.sort(key=lambda m: m.get("gesendet") or m.get("empfangen") or "")
    total = len(mails)
    result = {"thema": topic, "ermittelt_ueber": method, "anzahl": total,
              "mails": mails[-limit:]}
    if total > limit:
        result["hinweis"] = f"Nur die neuesten {limit} von {total} Mails."
    return result


# ============================================================
# Kalender
# ============================================================

def _calendar_filter_date(dt: datetime) -> str:
    if CALENDAR_DATE_FORMAT:
        return dt.strftime(CALENDAR_DATE_FORMAT)
    # Kurzes Datumsformat von Windows, so wie es Outlook erwartet.
    return dt.strftime("%x %H:%M")


def _event_summary(item, folder) -> dict:
    return {
        "entry_id": _safe(item, "EntryID", ""),
        "store_id": _safe(folder, "StoreID", ""),
        "betreff": _safe(item, "Subject", ""),
        "start": _fmt(_dt(_safe(item, "Start"))),
        "ende": _fmt(_dt(_safe(item, "End"))),
        "ganztaegig": bool(_safe(item, "AllDayEvent", False)),
        "ort": _safe(item, "Location", ""),
        "organisator": _safe(item, "Organizer", ""),
        "serie": bool(_safe(item, "IsRecurring", False)),
        "anzeigen_als": BUSY_STATUS.get(_safe(item, "BusyStatus"), ""),
        "art": MEETING_STATUS.get(_safe(item, "MeetingStatus"), ""),
        "meine_antwort": RESPONSE_STATUS.get(_safe(item, "ResponseStatus"), ""),
        "kategorien": _safe(item, "Categories", ""),
        "privat": _safe(item, "Sensitivity", 0) == SENSITIVITY_PRIVATE,
    }


def _iter_events(folder, start: datetime, end: datetime):
    """Termine (Serien aufgelöst), die den Zeitraum überschneiden,
    nach Beginn sortiert."""
    # Reihenfolge ist wichtig: Sort -> IncludeRecurrences -> Restrict,
    # sonst werden Serientermine nicht korrekt aufgelöst.
    items = folder.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    flt = (f"[Start] < '{_calendar_filter_date(end)}' AND "
           f"[End] > '{_calendar_filter_date(start)}'")
    restricted = items.Restrict(flt)

    scanned = 0
    item = restricted.GetFirst()
    while item is not None and scanned < MAX_SCAN:
        scanned += 1
        s = _dt(_safe(item, "Start"))
        e = _dt(_safe(item, "End"))
        if s and s >= end:
            break  # nach Start sortiert: keine weiteren Treffer
        # Zusätzliche Prüfung in Python, unabhängig vom Datumsformat.
        if s and e and s < end and e > start:
            yield item
        item = restricted.GetNext()


def _list_calendars_impl() -> dict:
    ns = _namespace()
    default = ns.GetDefaultFolder(FOLDER_CALENDAR)
    default_id = _safe(default, "EntryID", "")
    calendars = []
    for store_root in ns.Folders:
        for folder in _collect_folders(store_root, ITEM_TYPE_APPOINTMENT, 6):
            calendars.append({
                "pfad": _folder_path(folder),
                "name": _safe(folder, "Name", ""),
                "standard": _safe(folder, "EntryID", "") == default_id,
            })
    return {
        "anzahl": len(calendars),
        "kalender": calendars,
        "hinweis": ("Kalender von Kolleginnen und Kollegen können über den "
                    "Parameter 'shared_mailbox' von list_events abgefragt "
                    "werden, sofern eine Freigabe mit Leserechten besteht. "
                    "Für belegte Zeiten genügt Frei/Gebucht: get_free_busy "
                    "bzw. find_meeting_times."),
    }


def _list_events_impl(start_date, end_date, calendar, shared_mailbox, query,
                      limit) -> dict:
    limit = _limit(limit)
    start, end = _period(start_date, end_date, MAX_CALENDAR_DAYS)

    folder = (_shared_calendar(shared_mailbox) if shared_mailbox and shared_mailbox.strip()
              else _resolve_folder(calendar, FOLDER_CALENDAR))

    needle = (query or "").strip().lower()
    events = []
    for item in _iter_events(folder, start, end):
        summary = _event_summary(item, folder)
        haystack = " ".join(str(summary.get(k) or "") for k in
                            ("betreff", "ort", "organisator", "kategorien")).lower()
        if not needle or needle in haystack:
            events.append(summary)
            if len(events) >= limit:
                break

    result = {
        "kalender": _folder_path(folder),
        "zeitraum": {"von": _fmt(start), "bis": _fmt(end)},
        "anzahl": len(events),
        "termine": events,
    }
    if len(events) >= limit:
        result["hinweis"] = "Limit erreicht, Zeitraum eingrenzen für weitere Termine."
    elif not events and shared_mailbox and shared_mailbox.strip():
        result["hinweis"] = ("Keine Termine gefunden. Ist der Kalender nur für "
                             "Frei/Gebucht freigegeben (auch mit Betreff und Ort), "
                             "liefert Outlook hier keine Termine. Belegte Zeiten "
                             "liefert get_free_busy, gemeinsame freie Zeiten "
                             "find_meeting_times.")
    return result


def _get_event_impl(entry_id, store_id, max_body_chars) -> dict:
    item = _open_item(entry_id, store_id)
    folder = _safe(item, "Parent")
    data = _event_summary(item, folder)

    attendees = []
    try:
        for recipient in item.Recipients:
            attendees.append({
                "name": _safe(recipient, "Name", ""),
                "email": _recipient_smtp(recipient),
                "typ": RECIPIENT_TYPE.get(_safe(recipient, "Type"), ""),
                "antwort": RESPONSE_STATUS.get(_safe(recipient, "MeetingResponseStatus"), ""),
            })
    except Exception:
        pass

    if data["serie"]:
        try:
            pattern = item.GetRecurrencePattern()
            data["serienmuster"] = _recurrence_info(pattern)
            exceptions = _recurrence_exceptions(pattern)
            if exceptions:
                data["ausnahmen"] = exceptions
        except Exception as exc:
            log.warning("Serienmuster nicht lesbar: %s", exc)

    data["vertraulichkeit"] = SENSITIVITY.get(_safe(item, "Sensitivity"), "normal")
    if _safe(item, "ReminderSet", False):
        data["erinnerung_minuten_vorher"] = _safe(item, "ReminderMinutesBeforeStart")
    else:
        data["erinnerung_minuten_vorher"] = None
    body = _safe(item, "Body", "") or ""
    online = _online_meeting(item, body)
    if online:
        data["online_besprechung"] = online

    limit = max(500, min(int(max_body_chars or MAX_BODY_CHARS), MAX_BODY_CHARS))
    data.update({
        "erforderliche_teilnehmende": _safe(item, "RequiredAttendees", ""),
        "optionale_teilnehmende": _safe(item, "OptionalAttendees", ""),
        "teilnehmende": attendees,
        "kalender": _folder_path(folder),
        "anhaenge": _attachment_list(item),
        "text": _truncate(body, limit),
    })
    return data


def _weekdays(mask: int) -> list[str]:
    return [name for bit, name in WEEKDAY_MASK if mask & bit]


def _recurrence_text(info: dict, rtype: int, interval: int, mask: int,
                     day: int, month: int, instance: int) -> str:
    days = _weekdays(mask)
    month_name = MONTHS[month - 1] if 1 <= month <= 12 else ""
    nth = INSTANCE.get(instance, "")
    if rtype == 0:
        if mask and len(days) == 5 and "Samstag" not in days and "Sonntag" not in days:
            text = "Jeden Werktag"
        else:
            text = "Jeden Tag" if interval <= 1 else f"Alle {interval} Tage"
    elif rtype == 1:
        text = ("Jede Woche" if interval <= 1 else f"Alle {interval} Wochen")
        if days:
            text += " am " + ", ".join(days)
    elif rtype == 2:
        text = f"Am {day}. " + ("jedes Monats" if interval <= 1 else f"alle {interval} Monate")
    elif rtype == 3:
        text = (f"Am {nth} {' oder '.join(days)} "
                + ("jedes Monats" if interval <= 1 else f"alle {interval} Monate"))
    elif rtype == 5:
        text = ("Jedes Jahr" if interval <= 1 else f"Alle {interval} Jahre") + \
               f" am {day}. {month_name}"
    elif rtype == 6:
        text = ("Jedes Jahr" if interval <= 1 else f"Alle {interval} Jahre") + \
               f" am {nth} {' oder '.join(days)} im {month_name}"
    else:
        text = "Unbekanntes Muster"
    if info.get("ganztaegig"):
        text += ", ganztägig"
    elif info.get("uhrzeit_von") and info.get("uhrzeit_bis"):
        text += f", {info['uhrzeit_von']} bis {info['uhrzeit_bis']}"
    if info.get("beginn"):
        text += f", ab {info['beginn']}"
    if info.get("ohne_ende"):
        text += ", ohne Enddatum"
    elif info.get("anzahl_vorkommen"):
        text += f", {info['anzahl_vorkommen']} Mal (bis {info.get('ende')})"
    elif info.get("ende"):
        text += f", bis {info['ende']}"
    return text


def _recurrence_info(pattern) -> dict:
    rtype = _safe(pattern, "RecurrenceType", -1)
    interval = int(_safe(pattern, "Interval", 1) or 1)
    if rtype in (5, 6) and interval >= 12:
        interval //= 12  # Outlook gibt jährliche Intervalle in Monaten an
    mask = int(_safe(pattern, "DayOfWeekMask", 0) or 0)
    day = int(_safe(pattern, "DayOfMonth", 0) or 0)
    month = int(_safe(pattern, "MonthOfYear", 0) or 0)
    instance = int(_safe(pattern, "Instance", 0) or 0)
    start_time = _dt(_safe(pattern, "StartTime"))
    end_time = _dt(_safe(pattern, "EndTime"))
    no_end = bool(_safe(pattern, "NoEndDate", False))
    info = {
        "typ": RECURRENCE_TYPE.get(rtype, str(rtype)),
        "intervall": interval,
        "wochentage": _weekdays(mask) or None,
        "tag_im_monat": day if rtype in (2, 5) else None,
        "monat": MONTHS[month - 1] if rtype in (5, 6) and 1 <= month <= 12 else None,
        "woche_im_monat": INSTANCE.get(instance) if rtype in (3, 6) else None,
        "uhrzeit_von": start_time.strftime("%H:%M") if start_time else None,
        "uhrzeit_bis": end_time.strftime("%H:%M") if end_time else None,
        "dauer_minuten": _safe(pattern, "Duration"),
        "beginn": (_fmt(_dt(_safe(pattern, "PatternStartDate"))) or "")[:10] or None,
        "ende": None if no_end else (_fmt(_dt(_safe(pattern, "PatternEndDate"))) or "")[:10] or None,
        "ohne_ende": no_end,
        "anzahl_vorkommen": None if no_end else _safe(pattern, "Occurrences"),
    }
    info["beschreibung"] = _recurrence_text(info, rtype, interval, mask, day,
                                            month, instance)
    return {k: v for k, v in info.items() if v is not None}


def _recurrence_exceptions(pattern) -> list[dict]:
    result = []
    try:
        exceptions = pattern.Exceptions
        count = int(exceptions.Count)
    except Exception:
        return result
    for nr in range(1, count + 1):
        try:
            ex = exceptions.Item(nr)
            entry = {"urspruenglich": _fmt(_dt(_safe(ex, "OriginalDate")))}
            if _safe(ex, "Deleted", False):
                entry["status"] = "abgesagt bzw. entfernt"
            else:
                appt = ex.AppointmentItem
                entry.update({
                    "status": "geändert",
                    "start": _fmt(_dt(_safe(appt, "Start"))),
                    "ende": _fmt(_dt(_safe(appt, "End"))),
                    "betreff": _safe(appt, "Subject", ""),
                    "ort": _safe(appt, "Location", ""),
                })
            result.append(entry)
        except Exception as exc:
            log.warning("Serienausnahme %s nicht lesbar: %s", nr, exc)
    result.sort(key=lambda e: e.get("urspruenglich") or "")
    if len(result) > MAX_EXCEPTIONS:
        # Die neuesten Ausnahmen sind meist die relevanten.
        result = result[-MAX_EXCEPTIONS:]
    return result


ONLINE_PATTERNS = (
    ("Teams", r"https://teams\.(?:microsoft|live)\.com/(?:l/meetup-join|meet)/[^\s<>\"]+"),
    ("Zoom", r"https://[\w.-]*zoom\.us/(?:j|my|w)/[^\s<>\"]+"),
    ("Webex", r"https://[\w.-]*webex\.com/[^\s<>\"]+"),
    ("Google Meet", r"https://meet\.google\.com/[^\s<>\"]+"),
)


def _online_meeting(item, body: str) -> dict | None:
    url, provider = "", ""
    try:
        url = item.PropertyAccessor.GetProperty(TEAMS_URL_PROP) or ""
        provider = "Teams" if url else ""
    except Exception:
        pass
    if not url:
        for name, pattern in ONLINE_PATTERNS:
            match = re.search(pattern, body or "")
            if match:
                url, provider = match.group(0).rstrip(">).,;"), name
                break
    if not url:
        return None
    info = {"anbieter": provider, "link": url}
    match = re.search(r"(?i)(?:Besprechungs-ID|Meeting[- ]ID|ID der Besprechung)"
                      r"\s*:?\s*(\d[\d ]{5,}\d)", body or "")
    if match:
        info["besprechungs_id"] = match.group(1).strip()
    match = re.search(r"(?i)(?:Kenncode|Passcode|Passwort|Password)\s*:?\s*(\S+)",
                      body or "")
    if match:
        info["kenncode"] = match.group(1).strip()
    return info


def _list_open_invitations_impl(start_date, end_date, limit) -> dict:
    limit = _limit(limit)
    if not (end_date or "").strip():
        start, _ = _period(start_date, "", MAX_CALENDAR_DAYS)
        end = start + timedelta(days=60)
    else:
        start, end = _period(start_date, end_date, MAX_CALENDAR_DAYS)
    folder = _resolve_folder("", FOLDER_CALENDAR)

    invitations: dict[str, dict] = {}
    for item in _iter_events(folder, start, end):
        if (_safe(item, "ResponseStatus") != RESPONSE_NOT_RESPONDED
                or _safe(item, "MeetingStatus") != MEETING_RECEIVED):
            continue
        key = _safe(item, "GlobalAppointmentID", "") or _safe(item, "EntryID", "")
        if key in invitations:
            # Weiteres Vorkommen derselben Serie: nur zählen.
            invitations[key]["vorkommen_im_zeitraum"] += 1
            continue
        if len(invitations) >= limit:
            break
        summary = _event_summary(item, folder)
        summary["vorkommen_im_zeitraum"] = 1
        invitations[key] = summary

    result = {
        "zeitraum": {"von": _fmt(start), "bis": _fmt(end)},
        "anzahl": len(invitations),
        "einladungen": list(invitations.values()),
    }
    if len(invitations) >= limit:
        result["hinweis"] = "Limit erreicht, Zeitraum eingrenzen für weitere Einladungen."
    return result


# ============================================================
# Frei/Gebucht
# Recipient.FreeBusy liefert die Verfügbarkeit wie die
# Terminplanung in Outlook. Dafür genügt die Freigabe
# "Frei/Gebucht", Leserechte auf den Kalender sind nicht nötig.
# ============================================================

def _interval(value: int) -> int:
    interval = int(value or 15)
    if interval not in FREEBUSY_INTERVALS:
        raise ValueError("interval_minutes muss einer dieser Werte sein: "
                         + ", ".join(str(v) for v in FREEBUSY_INTERVALS))
    return interval


def _align(dt: datetime, interval: int) -> datetime:
    """Auf den Beginn des Intervalls abrunden."""
    minutes = dt.hour * 60 + dt.minute
    return dt.replace(hour=0, minute=0, second=0, microsecond=0) + \
        timedelta(minutes=minutes - minutes % interval)


def _parse_clock(text: str, name: str) -> int:
    """Uhrzeit HH:MM in Minuten seit Mitternacht."""
    try:
        t = datetime.strptime((text or "").strip(), "%H:%M")
    except ValueError:
        raise ValueError(f"Ungültige Uhrzeit für '{name}': '{text}'. Erlaubt: HH:MM")
    return t.hour * 60 + t.minute


def _mailboxes(emails) -> list[str]:
    """Adressen aus einer Liste oder einem Text (getrennt durch Komma,
    Semikolon oder Leerzeichen), ohne Duplikate."""
    if isinstance(emails, str):
        emails = [emails]
    result: list[str] = []
    for entry in emails or []:
        for part in re.split(r"[,;\s]+", entry or ""):
            if part and part.lower() not in (m.lower() for m in result):
                result.append(part)
    return result


def _free_busy(recipient, start: datetime, end: datetime, interval: int) -> str:
    """Status-Zeichen pro Intervall von start (auf das Intervall
    ausgerichtet) bis end."""
    day = start.replace(hour=0, minute=0, second=0, microsecond=0)
    per_call = FREEBUSY_DAYS_PER_CALL * 1440 // interval
    data = ""
    current = day
    while current < end:
        # Mittag statt Mitternacht übergeben: pywin32 rechnet die Zeit
        # nach UTC um, Outlook nimmt davon nur das Datum und würde sonst
        # beim Vortag beginnen.
        chunk = recipient.FreeBusy(current.replace(hour=12), interval, True)
        data += chunk[:per_call].ljust(per_call, "?")
        current += timedelta(days=FREEBUSY_DAYS_PER_CALL)
    first = int((start - day).total_seconds()) // 60 // interval
    last = math.ceil((end - day).total_seconds() / 60 / interval)
    return data[first:last]


def _free_busy_people(mailboxes: list[str], include_self: bool,
                      start: datetime, end: datetime,
                      interval: int) -> tuple[list[dict], list[dict]]:
    ns = _namespace()
    targets = [("", ns.CurrentUser)] if include_self else []
    for mailbox in mailboxes:
        recipient = ns.CreateRecipient(mailbox)
        recipient.Resolve()
        targets.append((mailbox, recipient))

    people, errors = [], []
    for mailbox, recipient in targets:
        if not recipient.Resolved:
            errors.append({"email": mailbox, "fehler": "Adresse konnte nicht aufgelöst werden."})
            continue
        try:
            slots = _free_busy(recipient, start, end, interval)
        except pywintypes.com_error as exc:
            log.warning("Frei/Gebucht für %s nicht verfügbar: %s", mailbox or "mich", exc)
            errors.append({"email": mailbox, "fehler": (
                "Frei/Gebucht-Informationen nicht verfügbar (Adresse prüfen, "
                "externe Adressen werden nicht unterstützt).")})
            continue
        people.append({"email": mailbox, "name": _safe(recipient, "Name", mailbox),
                       "slots": slots})
    return people, errors


def _busy_blocks(slots: str, start: datetime, interval: int) -> list[dict]:
    blocks = []
    pos = 0
    for char, group in groupby(slots):
        length = len(list(group))
        if char != "0":
            blocks.append({
                "von": _fmt(start + timedelta(minutes=pos * interval)),
                "bis": _fmt(start + timedelta(minutes=(pos + length) * interval)),
                "status": FREEBUSY_STATUS.get(char, "unbekannt"),
            })
        pos += length
    return blocks


def _get_free_busy_impl(emails, start_date, end_date, interval_minutes) -> dict:
    interval = _interval(interval_minutes)
    mailboxes = _mailboxes(emails)
    if not mailboxes:
        raise ValueError("Mindestens eine E-Mail-Adresse angeben.")
    start, end = _period(start_date, end_date, MAX_FREEBUSY_DAYS)
    start = _align(start, interval)

    people, errors = _free_busy_people(mailboxes, False, start, end, interval)
    result = {
        "zeitraum": {"von": _fmt(start), "bis": _fmt(end)},
        "intervall_minuten": interval,
        "personen": [{"email": p["email"], "name": p["name"],
                      "belegt": _busy_blocks(p["slots"], start, interval)}
                     for p in people],
        "hinweis": "Aufgeführt sind nur belegte Zeiten, alle übrigen sind frei.",
    }
    if errors:
        result["ohne_daten"] = errors
    return result


def _find_meeting_times_impl(emails, duration_minutes, start_date, end_date,
                             day_start, day_end, include_self, include_weekends,
                             allow_tentative, interval_minutes, limit) -> dict:
    interval = _interval(interval_minutes)
    limit = _limit(limit)
    duration = int(duration_minutes or 0)
    if duration < interval:
        raise ValueError(f"duration_minutes muss mindestens {interval} betragen.")
    hours_from = _parse_clock(day_start, "day_start")
    hours_to = _parse_clock(day_end, "day_end")
    if hours_to <= hours_from:
        raise ValueError("day_end muss nach day_start liegen.")
    mailboxes = _mailboxes(emails)
    if not mailboxes and not include_self:
        raise ValueError("Mindestens eine E-Mail-Adresse angeben.")
    start, end = _period(start_date, end_date, MAX_FREEBUSY_DAYS)
    start = _align(start, interval)

    people, errors = _free_busy_people(mailboxes, include_self, start, end, interval)
    if not people:
        raise ValueError("Für keine Person sind Frei/Gebucht-Informationen verfügbar: "
                         + "; ".join(f"{e['email']}: {e['fehler']}" for e in errors))

    free = {"0", "1"} if allow_tentative else {"0"}
    need = math.ceil(duration / interval)
    now = datetime.now()

    def usable(i: int) -> bool:
        t = start + timedelta(minutes=i * interval)
        minute = t.hour * 60 + t.minute
        return (t >= now
                and (include_weekends or t.weekday() < 5)
                and hours_from <= minute and minute + interval <= hours_to
                and all(p["slots"][i] in free for p in people))

    windows = []
    run_start = None
    count = min(len(p["slots"]) for p in people)
    for i in range(count + 1):
        if i < count and usable(i):
            if run_start is None:
                run_start = i
            continue
        if run_start is not None and i - run_start >= need:
            windows.append({
                "von": _fmt(start + timedelta(minutes=run_start * interval)),
                "bis": _fmt(start + timedelta(minutes=i * interval)),
                "minuten": (i - run_start) * interval,
            })
        run_start = None

    result = {
        "teilnehmende": [p["name"] for p in people],
        "zeitraum": {"von": _fmt(start), "bis": _fmt(end)},
        "dauer_minuten": duration,
        "anzahl": min(len(windows), limit),
        "zeitfenster": windows[:limit],
    }
    notes = []
    if errors:
        result["ohne_daten"] = errors
        notes.append("Personen unter 'ohne_daten' sind nicht berücksichtigt.")
    elsewhere = [p["name"] for p in people if {"4", "5"} & set(p["slots"])]
    if elsewhere:
        # Ein ganztägiger Eintrag "anderswo tätig" überdeckt in Outlook
        # alle anderen Termine des Tages.
        notes.append("'Anderswo tätig' gilt als belegt und überdeckt in Outlook "
                     "andere Termine (betrifft: " + ", ".join(elsewhere) + "). "
                     "Details mit get_free_busy prüfen.")
    if len(windows) > limit:
        notes.append("Limit erreicht, Zeitraum eingrenzen für weitere Zeitfenster.")
    if notes:
        result["hinweis"] = " ".join(notes)
    return result


# ============================================================
# Räume
# Exchange führt alle Räume in einer eigenen Adressliste
# ("Alle Räume"). Deren Name hängt von der Sprache ab und kann
# mit OUTLOOK_MCP_ROOM_LIST festgelegt werden.
# ============================================================

ROOM_LIST_OVERRIDE = os.environ.get("OUTLOOK_MCP_ROOM_LIST", "").strip().lower()
MAX_ROOM_AVAILABILITY = _env_int("OUTLOOK_MCP_MAX_ROOM_AVAILABILITY", 40)


def _room_address_list():
    names = {ROOM_LIST_OVERRIDE} if ROOM_LIST_OVERRIDE else ROOM_LIST_NAMES
    available = []
    for address_list in _namespace().AddressLists:
        name = (_safe(address_list, "Name", "") or "")
        available.append(name)
        if name.lower() in names:
            return address_list
    raise ValueError("Keine Raumliste gefunden. Vorhandene Adresslisten: "
                     + ", ".join(available) + ". Name der Raumliste mit der "
                     "Umgebungsvariablen OUTLOOK_MCP_ROOM_LIST festlegen.")


ROOM_CACHE_SECONDS = _env_int("OUTLOOK_MCP_ROOM_CACHE_SECONDS", 3600)
# Ist bei den ersten Räumen keine Kapazität gepflegt, wird sie nicht
# weiter abgefragt (spart eine Serveranfrage pro Raum).
ROOM_CAPACITY_PROBE = 10


def _room_info(entry, with_capacity: bool) -> dict:
    info = {"name": _safe(entry, "Name", ""), "email": ""}
    try:
        user = entry.GetExchangeUser()
    except Exception:
        user = None
    if user is not None:
        info["email"] = _safe(user, "PrimarySmtpAddress", "") or ""
        for key, attr in (("standort", "OfficeLocation"), ("ort", "City")):
            value = _safe(user, attr, "")
            if value:
                info[key] = value
    if with_capacity:
        try:
            capacity = entry.PropertyAccessor.GetProperty(PR_ROOM_CAPACITY)
            if isinstance(capacity, int) and capacity > 0:
                info["kapazitaet"] = capacity
        except Exception:
            pass
    return info


def _all_rooms() -> tuple[str, list[dict]]:
    """Alle Räume der Raumliste, eine Stunde zwischengespeichert, da das
    Lesen pro Raum eine Anfrage an Exchange braucht."""
    cached = _state.get("rooms")
    if cached and time.monotonic() - cached[0] < ROOM_CACHE_SECONDS:
        return cached[1], cached[2]
    address_list = _room_address_list()
    entries = address_list.AddressEntries
    rooms = []
    with_capacity = True
    for nr in range(1, int(entries.Count) + 1):
        rooms.append(_room_info(entries.Item(nr), with_capacity))
        if (with_capacity and nr == ROOM_CAPACITY_PROBE
                and not any("kapazitaet" in r for r in rooms)):
            with_capacity = False
    rooms.sort(key=lambda r: r["name"].lower())
    name = _safe(address_list, "Name", "")
    _state["rooms"] = (time.monotonic(), name, rooms)
    return name, rooms


def _list_rooms_impl(query, min_capacity, start, end, duration_minutes,
                     only_free, limit) -> dict:
    limit = max(1, min(int(limit or 100), 500))
    words = [w for w in (query or "").lower().split() if w]
    min_capacity = max(0, int(min_capacity or 0))

    check_from = check_to = None
    if (start or "").strip():
        if not _has_time(start):
            raise ValueError("start braucht Datum und Uhrzeit (JJJJ-MM-TT HH:MM).")
        check_from, check_to = _meeting_period(start, end, duration_minutes)
    elif only_free:
        raise ValueError("only_free braucht einen Zeitraum (start).")

    list_name, all_rooms = _all_rooms()
    rooms = []
    for room in all_rooms:
        haystack = " ".join(str(v) for v in room.values()).lower()
        if not all(w in haystack for w in words):
            continue
        if min_capacity and room.get("kapazitaet", 0) < min_capacity:
            continue
        rooms.append(dict(room))

    result = {"raumliste": list_name, "anzahl": len(rooms)}
    if min_capacity and not any("kapazitaet" in r for r in all_rooms):
        result["hinweis_kapazitaet"] = ("In der Raumliste ist keine Kapazität "
                                        "gepflegt, min_capacity ist daher nicht nutzbar.")

    if check_from:
        interval = 5
        aligned = _align(check_from, interval)
        ns = _namespace()
        checked = rooms[:MAX_ROOM_AVAILABILITY]
        for room in checked:
            if not room.get("email"):
                room["verfuegbarkeit"] = "unbekannt"
                continue
            try:
                recipient = ns.CreateRecipient(room["email"])
                recipient.Resolve()
                slots = _free_busy(recipient, aligned, check_to, interval)
                room["verfuegbarkeit"] = ("frei" if slots and set(slots) <= {"0"}
                                          else "belegt")
            except Exception as exc:
                log.warning("Frei/Gebucht für Raum %s nicht verfügbar: %s",
                            room["email"], exc)
                room["verfuegbarkeit"] = "unbekannt"
        if only_free:
            rooms = [r for r in checked if r["verfuegbarkeit"] == "frei"]
        result["zeitraum"] = {"von": _fmt(check_from), "bis": _fmt(check_to)}
        if len(checked) < result["anzahl"]:
            result["hinweis_verfuegbarkeit"] = (
                f"Verfügbarkeit nur für die ersten {len(checked)} Räume geprüft. "
                "Suche mit query oder min_capacity eingrenzen.")
        result["anzahl"] = len(rooms)

    result["raeume"] = rooms[:limit]
    if len(rooms) > limit:
        result["hinweis"] = f"Nur die ersten {limit} Räume, Suche eingrenzen."
    return result


# ============================================================
# Entwürfe
# Entwürfe werden nur gespeichert (Save), nie gesendet. Senden,
# Verschieben und Löschen bleiben der Benutzerin bzw. dem
# Benutzer in Outlook vorbehalten.
# ============================================================

# Beginn des Inhalts: nach <body>, bei Mails aus dem Outlook-Editor
# zusätzlich nach <div class=WordSection1>.
BODY_START = re.compile(
    r"<body\b[^>]*>(\s*<div\s+class=\"?WordSection1\"?[^>]*>)?", re.I)
BLOCK_TAG = re.compile(r"<(p|li|td|th|div|h[1-6])\b([^>]*)>", re.I)
STYLE_ATTR = re.compile(r"""\bstyle\s*=\s*(["'])""", re.I)


def _font_css() -> str:
    if not DRAFT_FONT:
        return ""
    return (f"font-family:&quot;{html.escape(DRAFT_FONT)}&quot;,sans-serif;"
            f"font-size:{max(6, min(DRAFT_FONT_SIZE, 72))}.0pt")


def _with_font(fragment: str) -> str:
    """Schrift an jedem Block setzen. Die Formatvorlage der Mail legt
    die Schrift für Absätze fest, ein umschliessendes Element genügt
    deshalb nicht. Eigene Angaben im Fragment haben Vorrang."""
    css = _font_css()
    if not css:
        return fragment

    def add(match):
        tag, attrs = match.group(1), match.group(2)
        style = STYLE_ATTR.search(attrs)
        if style:
            attrs = attrs[:style.end()] + css + ";" + attrs[style.end():]
        else:
            attrs = f' style="{css}"' + attrs
        return f"<{tag}{attrs}>"

    return f'<div style="{css}">' + BLOCK_TAG.sub(add, fragment) + "</div>"


def _body_html(text: str, is_html: bool) -> str:
    if is_html:
        return _with_font(text or "")
    # Absätze im Stil der Outlook-Vorlage (Abstände), Schrift aus _with_font.
    lines = (text or "").replace("\r\n", "\n").strip("\n").split("\n")
    return _with_font("".join(
        f"<p class=MsoNormal>{html.escape(line) if line.strip() else '&nbsp;'}</p>"
        for line in lines))


def _insert_body(item, fragment: str) -> None:
    """Text am Anfang des Inhalts einfügen (vor Signatur und Zitat).
    Kopf und Formatvorlagen bleiben."""
    # Format nur bei Bedarf setzen: Erneutes Setzen verwirft eine
    # gerade eingefügte Signatur.
    if _safe(item, "BodyFormat") != FORMAT_HTML:
        item.BodyFormat = FORMAT_HTML
    body = _safe(item, "HTMLBody", "") or ""
    match = BODY_START.search(body)
    if match:
        body = body[:match.end()] + fragment + body[match.end():]
    else:
        body = fragment + body
    item.HTMLBody = body


def _with_signature(item):
    """HTML-Format setzen und die Standardsignatur einfügen lassen.
    Outlook fügt sie ein, sobald das Element ein Fenster-Objekt
    erhält, auch wenn dieses nicht angezeigt wird. Der Inhalt wird
    danach sofort gelesen, damit spätere Änderungen sie nicht verwerfen."""
    item.BodyFormat = FORMAT_HTML
    try:
        item.GetInspector
        item.HTMLBody
    except Exception as exc:
        log.warning("Signatur konnte nicht eingefügt werden: %s", exc)
    return item


def _set_recipients(item, kind: int, addresses) -> None:
    for address in _mailboxes(addresses):
        item.Recipients.Add(address).Type = kind


def _unresolved(item) -> list[str]:
    try:
        item.Recipients.ResolveAll()
        return [_safe(r, "Name", "") for r in item.Recipients
                if not _safe(r, "Resolved", False)]
    except Exception:
        return []


def _open_mail(entry_id: str, store_id: str):
    item = _open_item(entry_id, store_id)
    if _safe(item, "Class") != CLASS_MAIL:
        raise ValueError("Das Element ist keine E-Mail. Entwürfe sind nur "
                         "für E-Mails möglich.")
    return item


def _save_draft(item, open_in_outlook: bool) -> dict:
    unresolved = _unresolved(item)
    item.Save()
    folder = _safe(item, "Parent")
    result = {
        "entry_id": _safe(item, "EntryID", ""),
        "store_id": _safe(folder, "StoreID", ""),
        "betreff": _safe(item, "Subject", ""),
        "an": _safe(item, "To", ""),
        "cc": _safe(item, "CC", ""),
        "bcc": _safe(item, "BCC", ""),
        "ordner": _folder_path(folder),
        "anhaenge": [a["name"] for a in _attachment_list(item)],
        "vorschau": _truncate(_safe(item, "Body", ""), 1000),
        "hinweis": ("Entwurf gespeichert, nicht gesendet. Gesendet wird "
                    "nur von Hand in Outlook."),
    }
    if unresolved:
        result["nicht_aufgeloest"] = unresolved
    if open_in_outlook:
        item.Display(False)
    else:
        try:
            item.Close(CLOSE_DISCARD)  # unsichtbares Fenster-Objekt freigeben
        except Exception:
            pass
    return result


def _create_draft_impl(to, cc, bcc, subject, body, body_is_html,
                       open_in_outlook) -> dict:
    item = _with_signature(_namespace().Application.CreateItem(ITEM_TYPE_MAIL))
    item.Subject = subject or ""
    _set_recipients(item, RECIPIENT_TO, to)
    _set_recipients(item, RECIPIENT_CC, cc)
    _set_recipients(item, RECIPIENT_BCC, bcc)
    if body:
        _insert_body(item, _body_html(body, body_is_html))
    return _save_draft(item, open_in_outlook)


def _create_reply_draft_impl(entry_id, store_id, body, body_is_html, reply_all,
                             cc, bcc, open_in_outlook) -> dict:
    original = _open_mail(entry_id, store_id)
    item = _with_signature(original.ReplyAll() if reply_all else original.Reply())
    _set_recipients(item, RECIPIENT_CC, cc)
    _set_recipients(item, RECIPIENT_BCC, bcc)
    if body:
        _insert_body(item, _body_html(body, body_is_html))
    return _save_draft(item, open_in_outlook)


def _create_forward_draft_impl(entry_id, store_id, to, cc, bcc, body,
                               body_is_html, open_in_outlook) -> dict:
    item = _with_signature(_open_mail(entry_id, store_id).Forward())
    _set_recipients(item, RECIPIENT_TO, to)
    _set_recipients(item, RECIPIENT_CC, cc)
    _set_recipients(item, RECIPIENT_BCC, bcc)
    if body:
        _insert_body(item, _body_html(body, body_is_html))
    return _save_draft(item, open_in_outlook)


# ============================================================
# Einladungen
# Der Termin wird nur in einem Outlook-Fenster geöffnet, weder
# gespeichert noch gesendet. Speichern oder Senden erfolgt dort
# von Hand.
# ============================================================

def _com_time(dt: datetime) -> datetime:
    # pywin32 rechnet naive Zeiten nach UTC um, eine als UTC markierte
    # Zeit übernimmt es unverändert. Outlook liest sie als Ortszeit.
    return dt.replace(tzinfo=timezone.utc)


def _meeting_period(start: str, end: str,
                    duration_minutes: int) -> tuple[datetime, datetime]:
    begin = _parse_date(start, "start")
    if begin is None or not _has_time(start):
        raise ValueError("start braucht Datum und Uhrzeit (JJJJ-MM-TT HH:MM).")
    if end and end.strip():
        if not _has_time(end):
            raise ValueError("end braucht Datum und Uhrzeit (JJJJ-MM-TT HH:MM).")
        finish = _parse_date(end, "end")
    else:
        finish = begin + timedelta(minutes=int(duration_minutes or 0))
    if finish <= begin:
        raise ValueError("Das Ende muss nach dem Beginn liegen.")
    return begin, finish


def _open_meeting_draft_impl(subject, start, end, duration_minutes, location,
                             required, optional, rooms, body) -> dict:
    begin, finish = _meeting_period(start, end, duration_minutes)
    item = _namespace().Application.CreateItem(ITEM_TYPE_APPOINTMENT)
    item.Subject = subject or ""
    item.Start = _com_time(begin)
    item.End = _com_time(finish)
    if location:
        item.Location = location
    if body:
        item.Body = body
    _set_recipients(item, RECIPIENT_REQUIRED, required)
    _set_recipients(item, RECIPIENT_OPTIONAL, optional)
    _set_recipients(item, RECIPIENT_RESOURCE, rooms)
    unresolved = []
    if int(item.Recipients.Count):
        item.MeetingStatus = MEETING_ORGANIZER
        unresolved = _unresolved(item)
    item.Display(False)

    result = {
        "betreff": _safe(item, "Subject", ""),
        "beginn": _fmt(_dt(_safe(item, "Start"))),
        "ende": _fmt(_dt(_safe(item, "End"))),
        "ort": _safe(item, "Location", ""),
        "erforderlich": _safe(item, "RequiredAttendees", ""),
        "optional": _safe(item, "OptionalAttendees", ""),
        "raeume": _safe(item, "Resources", ""),
        "hinweis": ("In Outlook geöffnet, nicht gespeichert und nicht "
                    "gesendet. Prüfen, bei Bedarf über die Schaltfläche eine "
                    "Teams-Besprechung hinzufügen, dann selbst senden oder "
                    "speichern. Wird das Fenster ohne Speichern geschlossen, "
                    "bleibt nichts zurück."),
    }
    if unresolved:
        result["nicht_aufgeloest"] = unresolved
    return result


# ============================================================
# Anhänge
# Der Anhang wird in ein eigenes temporäres Verzeichnis unter
# festem Namen gespeichert (kein Pfad aus dem Dateinamen), der
# Inhalt gelesen und das Verzeichnis danach wieder gelöscht.
# ============================================================

TEXT_EXTENSIONS = {".txt", ".csv", ".tsv", ".log", ".json", ".xml", ".md",
                   ".ics", ".vcf", ".ini", ".cfg", ".yaml", ".yml", ".sql",
                   ".ps1", ".py", ".js", ".css", ".rtf"}
HTML_EXTENSIONS = {".html", ".htm"}
IMAGE_FORMATS = {".png": "png", ".jpg": "jpeg", ".jpeg": "jpeg",
                 ".gif": "gif", ".webp": "webp"}
SUPPORTED_HINT = ("Unterstützt: PDF, DOCX, XLSX, PPTX, DOC, XLS, PPT, MSG, EML, ZIP "
                  "(Inhaltsverzeichnis), HTML, Textdateien, Bilder (PNG, JPG, GIF, WEBP).")

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
MAX_ZIP_MEMBER_BYTES = 50_000_000


def _read_bytes(path: str) -> bytes:
    with open(path, "rb") as fh:
        return fh.read()


def _decode(data: bytes) -> str:
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16", errors="replace")
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


def _html_to_text(markup: str) -> str:
    markup = re.sub(r"(?is)<(script|style|head)\b.*?</\1\s*>", " ", markup)
    markup = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|li|h[1-6]|table)\s*>", "\n", markup)
    markup = re.sub(r"(?i)</t[dh]\s*>", "\t", markup)
    markup = re.sub(r"<[^>]+>", "", markup)
    text = html.unescape(markup).replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n\s*\n+", "\n\n", text).strip()


def _pdf_text(path: str, limit: int) -> tuple[str, str | None]:
    from pypdf import PdfReader

    reader = PdfReader(path)
    if reader.is_encrypted:
        try:
            if not reader.decrypt(""):
                raise ValueError
        except Exception:
            raise ValueError("Das PDF ist passwortgeschützt und kann nicht gelesen werden.")
    parts, total = [], 0
    page_count = len(reader.pages)
    for nr, page in enumerate(reader.pages, 1):
        text = (page.extract_text() or "").strip()
        if text:
            parts.append(f"--- Seite {nr} ---\n{text}")
            total += len(text)
        if total > limit:
            break
    if not parts:
        return "", (f"{page_count} Seiten, aber kein Text gefunden "
                    "(vermutlich ein gescanntes PDF ohne Texterkennung).")
    return "\n\n".join(parts), f"{page_count} Seiten"


def _xlsx_text(path: str, limit: int) -> tuple[str, str | None]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        parts, total = [], 0
        sheets = workbook.worksheets
        for sheet in sheets:
            lines = [f"--- Blatt: {sheet.title} ---"]
            for row in sheet.iter_rows(values_only=True):
                cells = ["" if v is None else str(v) for v in row]
                while cells and not cells[-1]:
                    cells.pop()
                if not cells:
                    continue
                line = "\t".join(cells)
                lines.append(line)
                total += len(line) + 1
                if total > limit:
                    break
            parts.append("\n".join(lines))
            if total > limit:
                break
        return "\n\n".join(parts), f"{len(sheets)} Blätter, Zellen tabulatorgetrennt"
    finally:
        workbook.close()


# ---- Alte Office-Formate (Word, Excel, PowerPoint 97-2003) ----
# Gelesen werden nur die Daten, ohne Office zu starten: Makros und
# eingebettete Objekte werden nie ausgeführt.

def _ole_stream(ole, name: str) -> bytes:
    if not ole.exists(name):
        raise ValueError(f"Datei ist beschädigt oder kein Office-Dokument (Teil '{name}' fehlt).")
    if ole.get_size(name) > MAX_ZIP_MEMBER_BYTES:
        raise ValueError("Dokument zu gross zum Lesen.")
    return ole.openstream(name).read()


def _doc_clean(text: str) -> str:
    # Feldfunktionen: \x13 Code \x14 Ergebnis \x15 -> nur das Ergebnis behalten.
    out, depth_stack = [], []
    for ch in text:
        if ch == "\x13":
            depth_stack.append(True)   # im Feldcode
            continue
        if ch == "\x14":
            if depth_stack:
                depth_stack[-1] = False  # ab hier Feldergebnis
            continue
        if ch == "\x15":
            if depth_stack:
                depth_stack.pop()
            continue
        if any(depth_stack):
            continue
        out.append(ch)
    text = "".join(out)
    # \x07 beendet Tabellenzellen, zweimal hintereinander eine Tabellenzeile.
    text = text.replace("\x07\x07", "\n").replace("\x07", "\t")
    text = re.sub(r"[\r\x0b\x0c\x0e]", "\n", text)
    text = text.replace("\x1e", "-").replace("\xa0", " ")
    text = re.sub(r"[\x00-\x08\x1c\x1d\x1f]", "", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _doc_text(path: str) -> str:
    """Text aus Word 97-2003 (.doc) über die Stückliste (Piece Table)."""
    import olefile

    if not olefile.isOleFile(path):
        raise ValueError("Keine gültige Word-97-2003-Datei.")
    with olefile.OleFileIO(path) as ole:
        word = _ole_stream(ole, "WordDocument")
        if len(word) < 0x1AA:
            raise ValueError("Word-Dokument ist zu kurz oder beschädigt.")
        ident, nfib = struct.unpack_from("<HH", word, 0)
        flags = struct.unpack_from("<H", word, 0x0A)[0]
        if ident != 0xA5EC or nfib < 0x00C1:
            raise ValueError("Word-Format vor Word 97 wird nicht unterstützt.")
        if flags & 0x0100:
            raise ValueError("Das Word-Dokument ist passwortgeschützt und kann nicht gelesen werden.")
        table = _ole_stream(ole, "1Table" if flags & 0x0200 else "0Table")
        ccp_text = struct.unpack_from("<i", word, 0x4C)[0]
        fc_clx, lcb_clx = struct.unpack_from("<II", word, 0x1A2)
        clx = table[fc_clx:fc_clx + lcb_clx]

    pos = 0
    while pos < len(clx) and clx[pos] == 0x01:  # Prc-Einträge überspringen
        pos += 3 + struct.unpack_from("<H", clx, pos + 1)[0]
    if pos >= len(clx) or clx[pos] != 0x02:
        raise ValueError("Textstruktur des Word-Dokuments nicht gefunden.")
    lcb = struct.unpack_from("<I", clx, pos + 1)[0]
    plc = clx[pos + 5:pos + 5 + lcb]
    count = (lcb - 4) // 12
    cps = struct.unpack_from(f"<{count + 1}I", plc, 0)
    parts = []
    for i in range(count):
        cp_start, cp_end = cps[i], min(cps[i + 1], ccp_text)
        if cp_start >= cp_end:
            break
        fc = struct.unpack_from("<I", plc, (count + 1) * 4 + i * 8 + 2)[0]
        length = cp_end - cp_start
        if fc & 0x40000000:
            offset = (fc & 0x3FFFFFFF) // 2
            parts.append(word[offset:offset + length].decode("cp1252", errors="replace"))
        else:
            parts.append(word[fc:fc + 2 * length].decode("utf-16-le", errors="replace"))
    return _doc_clean("".join(parts))


def _xls_text(path: str, limit: int) -> tuple[str, str | None]:
    import xlrd

    try:
        book = xlrd.open_workbook(path, on_demand=True)
    except xlrd.biffh.XLRDError as exc:
        if "encrypt" in str(exc).lower():
            raise ValueError("Die Excel-Datei ist passwortgeschützt und kann nicht gelesen werden.")
        raise ValueError(f"Excel-Datei nicht lesbar: {exc}")
    try:
        parts, total = [], 0
        names = book.sheet_names()
        for index, title in enumerate(names):
            sheet = book.sheet_by_index(index)
            lines = [f"--- Blatt: {title} ---"]
            for r in range(sheet.nrows):
                cells = []
                for cell in sheet.row(r):
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            dt = xlrd.xldate.xldate_as_datetime(value, book.datemode)
                            value = (dt.strftime("%Y-%m-%d") if dt.time() == datetime.min.time()
                                     else dt.strftime("%Y-%m-%d %H:%M"))
                        except Exception:
                            pass
                    elif cell.ctype == xlrd.XL_CELL_NUMBER and float(value).is_integer():
                        value = int(value)
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        value = bool(value)
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK,
                                        xlrd.XL_CELL_ERROR):
                        value = ""
                    cells.append(str(value))
                while cells and not cells[-1]:
                    cells.pop()
                if not cells:
                    continue
                line = "\t".join(cells)
                lines.append(line)
                total += len(line) + 1
                if total > limit:
                    break
            book.unload_sheet(index)
            parts.append("\n".join(lines))
            if total > limit:
                break
        return "\n\n".join(parts), f"{len(names)} Blätter, Zellen tabulatorgetrennt"
    finally:
        book.release_resources()


# Datensatztypen im Stream "PowerPoint Document".
PPT_TEXT_CHARS, PPT_TEXT_BYTES = 0x0FA0, 0x0FA8
PPT_SLIDE, PPT_NOTES, PPT_MAIN_MASTER, PPT_SLIDE_LIST = 0x03EE, 0x03F0, 0x03F8, 0x0FF0
PPT_SLIDE_PERSIST, PPT_HANDOUT, PPT_NOTES_ATOM = 0x03F3, 0x0FC9, 0x03F1
PPT_SKIP = {PPT_MAIN_MASTER, PPT_HANDOUT, 0x0FD9}  # Master, Handzettel, Kopf-/Fusszeilen


def _ppt_text(path: str) -> tuple[str, str | None]:
    """Text aus PowerPoint 97-2003 (.ppt) durch Lesen der Textdatensätze."""
    import olefile

    if not olefile.isOleFile(path):
        raise ValueError("Keine gültige PowerPoint-97-2003-Datei.")
    with olefile.OleFileIO(path) as ole:
        if ole.exists("EncryptedSummary"):
            raise ValueError("Die PowerPoint-Datei ist passwortgeschützt und kann nicht gelesen werden.")
        data = _ole_stream(ole, "PowerPoint Document")

    blocks: list[tuple[str, list[str]]] = []

    def new_block(label: str):
        blocks.append((label, []))

    def walk(start: int, end: int, in_slide_list: bool):
        pos = start
        while pos + 8 <= end:
            ver_inst, rtype, length = struct.unpack_from("<HHI", data, pos)
            body, nxt = pos + 8, min(pos + 8 + length, end)
            if (ver_inst & 0x0F) == 0x0F:  # Container
                if rtype in PPT_SKIP:
                    pass
                elif rtype == PPT_SLIDE:
                    new_block("Folie")
                    walk(body, nxt, False)
                elif rtype == PPT_NOTES:
                    # NotesAtom mit slideIdRef 0 bzw. 0x80000000 (so von
                    # PowerPoint geschrieben) kennzeichnet den Notizen-Master.
                    child_type = struct.unpack_from("<H", data, body + 2)[0] \
                        if body + 12 <= nxt else 0
                    if not (child_type == PPT_NOTES_ATOM
                            and struct.unpack_from("<I", data, body + 8)[0]
                            in (0, 0x80000000)):
                        new_block("Notizen")
                        walk(body, nxt, False)
                elif rtype == PPT_SLIDE_LIST:
                    walk(body, nxt, (ver_inst >> 4) in (0, 2))
                else:
                    walk(body, nxt, in_slide_list)
            elif rtype == PPT_SLIDE_PERSIST and in_slide_list:
                new_block("Folie")
            elif rtype in (PPT_TEXT_CHARS, PPT_TEXT_BYTES):
                raw = data[body:nxt]
                text = (raw.decode("utf-16-le", errors="replace") if rtype == PPT_TEXT_CHARS
                        else raw.decode("latin-1"))
                text = text.replace("\r", "\n").replace("\x0b", "\n").strip()
                if text and text != "*":
                    if not blocks:
                        new_block("Text")
                    blocks[-1][1].append(text)
            pos = nxt

    walk(0, len(data), False)

    parts, seen, nr = [], set(), 0
    for label, texts in blocks:
        content = "\n".join(texts).strip()
        if not content or content in seen:
            continue  # leer oder Duplikat aus früherem Speichern
        seen.add(content)
        if label == "Folie":
            nr += 1
            label = f"Folie {nr}"
        parts.append(f"--- {label} ---\n{content}")
    return "\n\n".join(parts), ("Textfelder der Folien in Dateireihenfolge; "
                                "die Zuordnung zu Foliennummern ist ungefähr.")


def _zip_member(archive: zipfile.ZipFile, name: str) -> bytes:
    if archive.getinfo(name).file_size > MAX_ZIP_MEMBER_BYTES:
        raise ValueError("Dokument zu gross zum Lesen.")
    return archive.read(name)


def _ooxml_paragraphs(xml_data: bytes, ns: str) -> list[str]:
    root = ET.fromstring(xml_data)
    lines = []
    for paragraph in root.iter(ns + "p"):
        chunks = []
        for node in paragraph.iter():
            if node.tag == ns + "t":
                chunks.append(node.text or "")
            elif node.tag == ns + "tab":
                chunks.append("\t")
            elif node.tag in (ns + "br", ns + "cr"):
                chunks.append("\n")
        line = "".join(chunks).strip()
        if line:
            lines.append(line)
    return lines


def _docx_text(path: str) -> str:
    with zipfile.ZipFile(path) as archive:
        return "\n".join(_ooxml_paragraphs(
            _zip_member(archive, "word/document.xml"), W_NS))


def _pptx_text(path: str) -> tuple[str, str | None]:
    with zipfile.ZipFile(path) as archive:
        slides = sorted(
            (n for n in archive.namelist()
             if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"(\d+)\.xml$", n).group(1)),
        )
        parts = []
        for nr, name in enumerate(slides, 1):
            lines = _ooxml_paragraphs(_zip_member(archive, name), A_NS)
            parts.append(f"--- Folie {nr} ---\n" + "\n".join(lines))
    return "\n\n".join(parts), f"{len(slides)} Folien"


def _zip_listing(path: str) -> str:
    with zipfile.ZipFile(path) as archive:
        lines = [f"{info.filename} ({info.file_size} Bytes)"
                 for info in archive.infolist() if not info.is_dir()]
    return "Inhalt des ZIP-Archivs:\n" + "\n".join(lines)


def _eml_text(path: str) -> str:
    message = email.message_from_bytes(_read_bytes(path), policy=policy.default)
    lines = [
        f"Betreff: {message['subject'] or ''}",
        f"Von: {message['from'] or ''}",
        f"An: {message['to'] or ''}",
        f"Cc: {message['cc'] or ''}",
        f"Datum: {message['date'] or ''}",
    ]
    names = [p.get_filename() for p in message.iter_attachments() if p.get_filename()]
    if names:
        lines.append("Anhänge: " + ", ".join(names))
    body = message.get_body(preferencelist=("plain", "html"))
    text = ""
    if body is not None:
        text = body.get_content()
        if body.get_content_type() == "text/html":
            text = _html_to_text(text)
    return "\n".join(lines) + "\n\n" + text


def _msg_text(path: str) -> str:
    item = _namespace().OpenSharedItem(path)
    try:
        lines = [
            f"Betreff: {_safe(item, 'Subject', '')}",
            f"Von: {_safe(item, 'SenderName', '')} <{_smtp_of_sender(item)}>",
            f"An: {_safe(item, 'To', '')}",
            f"Cc: {_safe(item, 'CC', '')}",
            f"Gesendet: {_fmt(_dt(_safe(item, 'SentOn'))) or ''}",
        ]
        names = [a["name"] for a in _attachment_list(item)]
        if names:
            lines.append("Anhänge: " + ", ".join(names) +
                         " (nicht direkt lesbar, da in der angehängten Mail)")
        return "\n".join(lines) + "\n\n" + (_safe(item, "Body", "") or "")
    finally:
        try:
            item.Close(1)  # olDiscard: nichts speichern, Datei freigeben
        except Exception:
            pass
        del item


def _image_content(path: str, ext: str, info: dict) -> Image | None:
    """Bild unverändert übergeben, wenn es klein genug ist, sonst
    verkleinern. Claude skaliert grosse Bilder ohnehin herunter."""
    from PIL import Image as PILImage

    data = _read_bytes(path)
    with PILImage.open(path) as img:
        width, height = img.size
        info["abmessungen"] = f"{width}x{height}"
        if len(data) <= MAX_IMAGE_BYTES and max(width, height) <= MAX_IMAGE_EDGE:
            return Image(data=data, format=IMAGE_FORMATS[ext])

        img.seek(0)  # bei animierten Bildern das erste Bild
        img = img.copy()
        img.thumbnail((MAX_IMAGE_EDGE, MAX_IMAGE_EDGE))
        transparent = img.mode in ("RGBA", "LA", "P") and (
            img.mode != "P" or "transparency" in img.info)

    if transparent:
        buffer = io.BytesIO()
        img.convert("RGBA").save(buffer, format="PNG", optimize=True)
        if buffer.tell() <= MAX_IMAGE_BYTES:
            info["hinweis"] = f"Bild verkleinert auf {img.width}x{img.height}."
            return Image(data=buffer.getvalue(), format="png")
    rgb = img.convert("RGBA").convert("RGB") if transparent else img.convert("RGB")
    for quality in (85, 70, 50):
        buffer = io.BytesIO()
        rgb.save(buffer, format="JPEG", quality=quality, optimize=True)
        if buffer.tell() <= MAX_IMAGE_BYTES:
            info["hinweis"] = f"Bild verkleinert auf {img.width}x{img.height} (JPEG)."
            return Image(data=buffer.getvalue(), format="jpeg")
    return None


def _extract_attachment(path: str, ext: str, offset: int, limit: int, info: dict):
    if ext in IMAGE_FORMATS:
        try:
            image = _image_content(path, ext, info)
        except Exception as exc:
            info["hinweis"] = f"Bild konnte nicht gelesen werden: {exc}"
            return info
        if image is None:
            info["hinweis"] = "Bild ist auch verkleinert zu gross für die Übertragung."
            return info
        info["inhalt"] = "Bild, siehe nachfolgende Bilddaten."
        return [info, image]

    # Formate mit vielen Seiten werden nur so weit gelesen wie nötig.
    need = offset + limit + 1
    note = None
    if ext == ".pdf":
        text, note = _pdf_text(path, need)
    elif ext in (".xlsx", ".xlsm"):
        text, note = _xlsx_text(path, need)
    elif ext in (".docx", ".docm"):
        text = _docx_text(path)
    elif ext in (".pptx", ".pptm"):
        text, note = _pptx_text(path)
    elif ext in (".doc", ".dot"):
        text = _doc_text(path)
    elif ext in (".xls", ".xlt"):
        text, note = _xls_text(path, need)
    elif ext in (".ppt", ".pps", ".pot"):
        text, note = _ppt_text(path)
    elif ext == ".msg":
        text = _msg_text(path)
    elif ext == ".eml":
        text = _eml_text(path)
    elif ext == ".zip":
        text = _zip_listing(path)
    elif ext in HTML_EXTENSIONS:
        text = _html_to_text(_decode(_read_bytes(path)))
    elif ext in TEXT_EXTENSIONS:
        text = _decode(_read_bytes(path))
    else:
        info["hinweis"] = (f"Format '{ext or 'ohne Endung'}' wird nicht "
                           f"unterstützt. {SUPPORTED_HINT}")
        return info

    text = (text or "").replace("\r\n", "\n").strip()
    if offset and offset >= len(text):
        raise ValueError(f"offset {offset} liegt hinter dem Textende "
                         f"({len(text)} Zeichen).")
    end = offset + limit
    info["text"] = text[offset:end]
    info["von_zeichen"] = offset
    if len(text) > end:
        info["naechster_offset"] = end
        info["fortsetzung"] = (f"Text geht weiter. Für den nächsten Teil "
                               f"get_attachment mit offset={end} aufrufen.")
    if note:
        info["hinweis"] = note
    return info


def _get_attachment_impl(entry_id, store_id, index, name, offset, max_chars):
    item = _open_item(entry_id, store_id)
    attachments = item.Attachments
    count = int(attachments.Count)
    if count == 0:
        raise ValueError("Das Element hat keine Anhänge.")

    if index:
        if not 1 <= index <= count:
            raise ValueError(f"Ungültige Anhangnummer {index} (vorhanden: 1 bis {count}).")
        att = attachments.Item(index)
    elif name and name.strip():
        wanted = name.strip().lower()
        att = None
        for nr in range(1, count + 1):
            candidate = attachments.Item(nr)
            if _attachment_name(candidate).lower() == wanted:
                att, index = candidate, nr
                break
        if att is None:
            raise ValueError(f"Anhang '{name}' nicht gefunden. "
                             "Die Liste der Anhänge liefert get_mail bzw. get_event.")
    elif count == 1:
        att, index = attachments.Item(1), 1
    else:
        raise ValueError(f"Das Element hat {count} Anhänge: bitte 'index' oder "
                         "'name' angeben (Liste über get_mail bzw. get_event).")

    info = _attachment_info(att, index, _safe(item, "HTMLBody", "") or "")
    att_type = _safe(att, "Type")
    if att_type == ATT_BY_REFERENCE:
        info["hinweis"] = "Verknüpfung auf eine externe Datei, Inhalt ist nicht in der Mail gespeichert."
        return info
    size = _safe(att, "Size", 0) or 0
    if size > MAX_ATTACHMENT_BYTES:
        info["hinweis"] = f"Anhang zu gross ({size} Bytes, max. {MAX_ATTACHMENT_BYTES})."
        return info

    ext = ".msg" if att_type == ATT_EMBEDDED else os.path.splitext(info["name"])[1].lower()
    if not re.fullmatch(r"\.[a-z0-9]{1,10}", ext):
        ext = ""
    limit = max(500, min(int(max_chars or MAX_ATTACHMENT_CHARS), MAX_ATTACHMENT_CHARS))
    offset = max(0, int(offset or 0))

    tmpdir = tempfile.mkdtemp(prefix="outlook-mcp-")
    try:
        path = os.path.join(tmpdir, "anhang" + ext)
        att.SaveAsFile(path)
        return _extract_attachment(path, ext, offset, limit, info)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# MCP-Werkzeuge
# ============================================================

mcp = FastMCP(
    "outlook",
    instructions=(
        "Lesender Zugriff auf E-Mails und Kalender im klassischen Outlook, "
        "dazu E-Mail-Entwürfe. Erst mit search_mails bzw. list_events "
        "suchen, dann Details mit get_mail bzw. get_event abrufen, den "
        "ganzen Gesprächsverlauf mit get_conversation. Anhänge "
        "mit get_attachment lesen (Nummer 'nr' aus get_mail). Für die "
        "Verfügbarkeit anderer Personen get_free_busy, für gemeinsame freie "
        "Zeiten find_meeting_times verwenden. Unbeantwortete Einladungen "
        "liefert list_open_invitations, Besprechungsräume list_rooms. "
        "Entwürfe mit create_draft, "
        "create_reply_draft und create_forward_draft anlegen. "
        "Besprechungseinladungen öffnet "
        "open_meeting_draft vorbefüllt in einem Outlook-Fenster, ohne zu "
        "speichern. Senden und Löschen ist nicht möglich: Entwürfe und "
        "Einladungen sendet die Benutzerin bzw. der Benutzer selbst in "
        "Outlook. Datumsangaben im Format JJJJ-MM-TT."
    ),
)


@mcp.tool()
async def list_folders(max_depth: int = 3) -> dict:
    """Listet die E-Mail-Ordner aller eingebundenen Postfächer mit Pfad,
    Anzahl Elemente und Anzahl ungelesener Elemente auf.

    max_depth: Wie tief Unterordner durchsucht werden (Standard 3).
    """
    return await _run(_list_folders_impl, max(0, min(max_depth, 10)))


@mcp.tool()
async def search_mails(
    query: str = "",
    folder: str = "",
    include_subfolders: bool = False,
    sender: str = "",
    date_from: str = "",
    date_to: str = "",
    unread_only: bool = False,
    has_attachments: bool = False,
    search_body: bool = True,
    limit: int = 20,
    flagged: bool = False,
    category: str = "",
    importance: str = "",
    recipient: str = "",
    attachment_name: str = "",
    use_index: bool = False,
) -> dict:
    """Durchsucht E-Mails, neueste zuerst. Liefert Kurzinfos mit entry_id
    und store_id für get_mail.

    query: Suchbegriff in Betreff und (bei search_body) im Text.
    folder: Leer = Posteingang. Kurzname ('Posteingang', 'Gesendet',
        'Entwürfe', 'Gelöscht') oder vollständiger Pfad aus list_folders.
        Für das ganze Postfach den Stammpfad (z. B. '\\\\max@firma.ch')
        mit include_subfolders=true angeben.
    include_subfolders: Unterordner mit durchsuchen.
    sender: Teil des Absendernamens oder der Absenderadresse.
    date_from / date_to: Zeitraum (JJJJ-MM-TT), beide inklusive.
    unread_only: Nur ungelesene Mails.
    has_attachments: Nur Mails mit Anhang.
    search_body: Auch im Nachrichtentext suchen (langsamer).
    limit: Maximale Anzahl Treffer.
    flagged: Nur gekennzeichnete (nicht erledigte) Mails.
    category: Teil eines Kategorienamens (z. B. 'Rot').
    importance: 'niedrig', 'normal' oder 'hoch'.
    recipient: Teil von Name oder Adresse in An oder Cc.
    attachment_name: Teil eines Anhangnamens (z. B. '.pdf' oder 'Rechnung').
    use_index: Den Outlook-Suchindex verwenden (schneller, besonders mit
        search_body; findet aber nur ganze Wörter bzw. Wortanfänge).
    """
    return await _run(_search_mails_impl, query, folder, include_subfolders,
                      sender, date_from, date_to, unread_only,
                      has_attachments, search_body, limit, flagged, category,
                      importance, recipient, attachment_name, use_index)


@mcp.tool()
async def get_mail(entry_id: str, store_id: str = "",
                   max_body_chars: int = MAX_BODY_CHARS,
                   body_format: str = "text",
                   include_headers: bool = False) -> dict:
    """Liest eine einzelne E-Mail vollständig: Kopfdaten inkl. BCC,
    Wichtigkeit, Vertraulichkeit, angeforderte Lese- und
    Zustellbestätigung, Kennzeichnung (Flag mit Fälligkeit), letzte
    Aktion (beantwortet/weitergeleitet mit Zeitpunkt), Text, Links und
    Liste der Anhänge mit Nummer 'nr'. Die Mail wird dabei nicht als
    gelesen markiert. Den Inhalt eines Anhangs liefert get_attachment,
    den ganzen Gesprächsverlauf get_conversation.

    entry_id / store_id: Werte aus search_mails.
    max_body_chars: Maximale Textlänge.
    body_format: 'text' (Standard, mit Liste der Links) oder 'html'
        (HTML-Quelltext, z. B. für Tabellen und Formatierung).
    include_headers: Internet-Header mitliefern, dazu SPF-, DKIM- und
        DMARC-Ergebnis (z. B. zur Prüfung auf Phishing).
    """
    return await _run(_get_mail_impl, entry_id, store_id, max_body_chars,
                      body_format, include_headers)


@mcp.tool()
async def get_conversation(entry_id: str, store_id: str = "",
                           limit: int = 50) -> dict:
    """Liefert alle Mails des Gesprächs (Threads), zu dem eine Mail gehört,
    über alle Ordner hinweg (z. B. Posteingang und Gesendete Elemente),
    chronologisch. 'ebene' gibt die Tiefe in der Antwortkette an. Details
    einzelner Mails mit get_mail.

    entry_id / store_id: Werte einer Mail aus search_mails.
    limit: Maximale Anzahl Mails (die neuesten), höchstens 200.
    """
    return await _run(_get_conversation_impl, entry_id, store_id, limit)


@mcp.tool()
async def get_attachment(entry_id: str, store_id: str = "", index: int = 0,
                         name: str = "", offset: int = 0,
                         max_chars: int = MAX_ATTACHMENT_CHARS):
    """Liest den Inhalt eines Anhangs einer E-Mail oder eines Termins.
    Text wird aus PDF, DOCX, XLSX, PPTX, den alten Office-Formaten DOC,
    XLS und PPT, angehängten Mails (MSG, EML),
    HTML- und Textdateien extrahiert, bei ZIP wird das Inhaltsverzeichnis
    geliefert. Bilder (PNG, JPG, GIF, WEBP) werden als Bild übergeben,
    grosse Bilder vorher verkleinert. Die Mail wird dabei nicht verändert.

    Lange Texte werden in Teilen geliefert: Enthält die Antwort
    'naechster_offset', mit diesem Wert als offset erneut aufrufen.

    entry_id / store_id: Werte aus search_mails bzw. list_events.
    index: Nummer 'nr' des Anhangs aus get_mail bzw. get_event.
    name: Alternativ der Dateiname des Anhangs. Bei nur einem Anhang
        können index und name weggelassen werden.
    offset: Startposition im Text (Zeichen), Standard 0.
    max_chars: Maximale Textlänge pro Aufruf.
    """
    return await _run(_get_attachment_impl, entry_id, store_id, index,
                      name, offset, max_chars)


@mcp.tool()
async def list_calendars() -> dict:
    """Listet alle Kalenderordner der eingebundenen Postfächer auf."""
    return await _run(_list_calendars_impl)


@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    calendar: str = "",
    shared_mailbox: str = "",
    query: str = "",
    limit: int = 50,
) -> dict:
    """Listet Kalendertermine in einem Zeitraum auf, chronologisch.
    Serientermine werden in einzelne Vorkommen aufgelöst.

    start_date: Beginn (JJJJ-MM-TT), Standard heute.
    end_date: Ende inklusive (JJJJ-MM-TT), Standard start_date + 7 Tage.
    calendar: Leer = eigener Standardkalender, sonst Pfad aus list_calendars.
    shared_mailbox: E-Mail-Adresse einer Person oder eines Raums, deren
        freigegebener Standardkalender gelesen werden soll.
    query: Filter auf Betreff, Ort, Organisator und Kategorien.
    limit: Maximale Anzahl Termine.
    """
    return await _run(_list_events_impl, start_date, end_date, calendar,
                      shared_mailbox, query, limit)


@mcp.tool()
async def get_event(entry_id: str, store_id: str = "",
                    max_body_chars: int = MAX_BODY_CHARS) -> dict:
    """Liest einen Termin mit Teilnehmenden, Antworten, Beschreibung,
    Erinnerung, Vertraulichkeit (privat) und Online-Besprechung (Teams-,
    Zoom- oder Webex-Link, Besprechungs-ID, Kenncode). Bei Serienterminen
    wird der Serienmaster geliefert, mit genauem Serienmuster (Rhythmus,
    Wochentage, Ende) und den Ausnahmen (geänderte oder abgesagte
    einzelne Vorkommen).

    entry_id / store_id: Werte aus list_events.
    max_body_chars: Maximale Textlänge der Beschreibung.
    """
    return await _run(_get_event_impl, entry_id, store_id, max_body_chars)


@mcp.tool()
async def get_free_busy(
    emails: list[str],
    start_date: str = "",
    end_date: str = "",
    interval_minutes: int = 15,
) -> dict:
    """Liefert die belegten Zeiten von Personen oder Räumen (frei,
    mit Vorbehalt, beschäftigt, abwesend, anderswo tätig), wie sie die
    Terminplanung in Outlook zeigt. Funktioniert auch ohne Leserechte
    auf den Kalender, die Freigabe "Frei/Gebucht" genügt. Betreff und
    Ort werden nicht geliefert.

    emails: E-Mail-Adressen der Personen oder Räume.
    start_date: Beginn (JJJJ-MM-TT oder JJJJ-MM-TT HH:MM), Standard heute.
    end_date: Ende inklusive (JJJJ-MM-TT), Standard start_date + 7 Tage.
    interval_minutes: Auflösung in Minuten (5, 10, 15, 30 oder 60).
    """
    return await _run(_get_free_busy_impl, emails, start_date, end_date,
                      interval_minutes)


@mcp.tool()
async def find_meeting_times(
    emails: list[str],
    duration_minutes: int = 30,
    start_date: str = "",
    end_date: str = "",
    day_start: str = "08:00",
    day_end: str = "17:00",
    include_self: bool = True,
    include_weekends: bool = False,
    allow_tentative: bool = False,
    interval_minutes: int = 15,
    limit: int = 20,
) -> dict:
    """Sucht Zeitfenster, in denen alle Teilnehmenden frei sind, auf
    Basis der Frei/Gebucht-Informationen (keine Leserechte nötig).
    Liefert zusammenhängende freie Zeitfenster, die mindestens
    duration_minutes lang sind, frühestens ab jetzt. "Anderswo tätig"
    gilt als belegt.

    emails: E-Mail-Adressen der weiteren Teilnehmenden oder Räume.
    duration_minutes: Mindestdauer des Termins.
    start_date / end_date: Suchzeitraum (JJJJ-MM-TT), Ende inklusive,
        Standard heute bis heute + 7 Tage.
    day_start / day_end: Tageszeitfenster (HH:MM), Standard 08:00 bis 17:00.
    include_self: Eigenen Kalender mit berücksichtigen.
    include_weekends: Auch Samstag und Sonntag durchsuchen.
    allow_tentative: Zeiten "mit Vorbehalt" als frei werten.
    interval_minutes: Raster in Minuten (5, 10, 15, 30 oder 60).
    limit: Maximale Anzahl Zeitfenster.
    """
    return await _run(_find_meeting_times_impl, emails, duration_minutes,
                      start_date, end_date, day_start, day_end, include_self,
                      include_weekends, allow_tentative, interval_minutes, limit)


@mcp.tool()
async def list_open_invitations(start_date: str = "", end_date: str = "",
                                limit: int = 50) -> dict:
    """Listet Besprechungseinladungen im eigenen Kalender, die noch nicht
    beantwortet sind (weder zugesagt, mit Vorbehalt noch abgelehnt).
    Serien erscheinen einmal, mit Anzahl Vorkommen im Zeitraum.

    start_date: Beginn (JJJJ-MM-TT), Standard heute.
    end_date: Ende inklusive (JJJJ-MM-TT), Standard start_date + 60 Tage.
    limit: Maximale Anzahl Einladungen.
    """
    return await _run(_list_open_invitations_impl, start_date, end_date, limit)


@mcp.tool()
async def list_rooms(
    query: str = "",
    min_capacity: int = 0,
    start: str = "",
    end: str = "",
    duration_minutes: int = 60,
    only_free: bool = False,
    limit: int = 100,
) -> dict:
    """Listet Besprechungsräume aus der Exchange-Raumliste ("Alle Räume")
    mit E-Mail-Adresse, Standort und Kapazität. Mit start wird zusätzlich
    geprüft, ob jeder Raum im Zeitraum frei ist (Frei/Gebucht). Die
    Adresse eines Raums kann für open_meeting_draft (rooms) verwendet
    werden.

    query: Suchwörter in Name, Standort oder Adresse (z. B. 'TN' oder
        'Winterthur'), alle müssen vorkommen.
    min_capacity: Mindestanzahl Plätze (Räume ohne Angabe fallen weg).
    start: Beginn der Verfügbarkeitsprüfung (JJJJ-MM-TT HH:MM).
    end: Ende (JJJJ-MM-TT HH:MM). Leer = start + duration_minutes.
    duration_minutes: Dauer, wenn end leer ist (Standard 60).
    only_free: Nur Räume liefern, die im Zeitraum frei sind.
    limit: Maximale Anzahl Räume.
    """
    return await _run(_list_rooms_impl, query, min_capacity, start, end,
                      duration_minutes, only_free, limit)


@mcp.tool()
async def create_draft(
    to: list[str] | None = None,
    subject: str = "",
    body: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_is_html: bool = False,
    open_in_outlook: bool = False,
) -> dict:
    """Legt eine neue E-Mail als Entwurf im Ordner Entwürfe an, mit der
    Standardsignatur aus Outlook. Die Mail wird nicht gesendet.

    to / cc / bcc: E-Mail-Adressen der Empfängerinnen und Empfänger.
    subject: Betreff.
    body: Text der Mail, Zeilenumbrüche werden übernommen. Steht über
        der Signatur.
    body_is_html: body ist ein HTML-Fragment (z. B. mit <p>, <b>, <ul>).
    open_in_outlook: Entwurf zusätzlich in einem Outlook-Fenster öffnen.
    """
    return await _run(_create_draft_impl, to, cc, bcc, subject, body,
                      body_is_html, open_in_outlook)


@mcp.tool()
async def create_reply_draft(
    entry_id: str,
    body: str = "",
    store_id: str = "",
    reply_all: bool = False,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_is_html: bool = False,
    open_in_outlook: bool = False,
) -> dict:
    """Legt eine Antwort auf eine E-Mail als Entwurf an, wie "Antworten"
    bzw. "Allen antworten" in Outlook: Empfänger und Betreff werden
    übernommen, der Text steht über Signatur und zitierter Mail. Die
    Antwort wird nicht gesendet, die Originalmail nicht verändert.

    entry_id / store_id: Werte aus search_mails.
    body: Text der Antwort, Zeilenumbrüche werden übernommen.
    reply_all: An alle Empfängerinnen und Empfänger antworten.
    cc / bcc: Zusätzliche E-Mail-Adressen.
    body_is_html: body ist ein HTML-Fragment.
    open_in_outlook: Entwurf zusätzlich in einem Outlook-Fenster öffnen.
    """
    return await _run(_create_reply_draft_impl, entry_id, store_id, body,
                      body_is_html, reply_all, cc, bcc, open_in_outlook)


@mcp.tool()
async def create_forward_draft(
    entry_id: str,
    to: list[str] | None = None,
    body: str = "",
    store_id: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_is_html: bool = False,
    open_in_outlook: bool = False,
) -> dict:
    """Legt eine Weiterleitung einer E-Mail als Entwurf an, mit allen
    Anhängen der Originalmail. Der Text steht über Signatur und
    weitergeleiteter Mail. Die Weiterleitung wird nicht gesendet.

    entry_id / store_id: Werte aus search_mails.
    to / cc / bcc: E-Mail-Adressen der Empfängerinnen und Empfänger.
    body: Begleittext, Zeilenumbrüche werden übernommen.
    body_is_html: body ist ein HTML-Fragment.
    open_in_outlook: Entwurf zusätzlich in einem Outlook-Fenster öffnen.
    """
    return await _run(_create_forward_draft_impl, entry_id, store_id, to, cc,
                      bcc, body, body_is_html, open_in_outlook)


@mcp.tool()
async def open_meeting_draft(
    start: str,
    subject: str = "",
    end: str = "",
    duration_minutes: int = 30,
    location: str = "",
    required: list[str] | None = None,
    optional: list[str] | None = None,
    rooms: list[str] | None = None,
    body: str = "",
) -> dict:
    """Öffnet eine vorbefüllte Besprechungseinladung in einem Outlook-
    Fenster. Sie wird weder gespeichert noch gesendet und steht erst im
    Kalender, wenn die Benutzerin bzw. der Benutzer sie in Outlook
    speichert oder sendet. Danach kann sie mit diesem Server nicht mehr
    geändert werden. Ohne Teilnehmende entsteht ein einfacher Termin.
    Eine Teams-Besprechung wird im Fenster über die Schaltfläche
    hinzugefügt. Freie Zeiten vorher mit find_meeting_times suchen.

    start: Beginn (JJJJ-MM-TT HH:MM), Ortszeit.
    subject: Betreff.
    end: Ende (JJJJ-MM-TT HH:MM). Leer = start + duration_minutes.
    duration_minutes: Dauer in Minuten, wenn end leer ist (Standard 30).
    location: Ort.
    required / optional: E-Mail-Adressen der erforderlichen bzw.
        optionalen Teilnehmenden.
    rooms: E-Mail-Adressen von Räumen oder Ressourcen.
    body: Text der Einladung (reiner Text, Zeilenumbrüche werden übernommen).
    """
    return await _run(_open_meeting_draft_impl, subject, start, end,
                      duration_minutes, location, required, optional, rooms,
                      body)


def main() -> None:
    try:
        locale.setlocale(locale.LC_TIME, "")
    except locale.Error:
        log.warning("Systemgebietsschema konnte nicht gesetzt werden")
    log.info("Outlook MCP Server %s startet (lesend, Entwürfe ohne Senden)",
             __version__)
    mcp.run()


if __name__ == "__main__":
    main()
