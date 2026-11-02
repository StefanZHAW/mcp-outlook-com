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
"openpyxl" und "pillow".
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
import sys
import tempfile
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
    return {
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


def _search_mails_impl(query, folder, include_subfolders, sender, date_from,
                       date_to, unread_only, has_attachments, search_body,
                       limit) -> dict:
    limit = _limit(limit)
    dt_from = _parse_date(date_from, "date_from")
    dt_to = _parse_date(date_to, "date_to")
    if dt_to and not _has_time(date_to):
        dt_to += timedelta(days=1)  # Enddatum inklusive

    root = _resolve_folder(folder, FOLDER_INBOX)
    folders = (_collect_folders(root, ITEM_TYPE_MAIL, 10)
               if include_subfolders else [root])

    conditions = []
    if query and query.strip():
        q = _dasl_text(query.strip())
        parts = [f"\"urn:schemas:httpmail:subject\" LIKE '%{q}%'"]
        if search_body:
            parts.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{q}%'")
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
    dasl = "@SQL=" + " AND ".join(conditions) if conditions else None

    hits: list[dict] = []
    scanned = 0
    truncated = False
    for current in folders:
        try:
            items = current.Items
            if dasl:
                items = items.Restrict(dasl)
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
            if not (received and dt_to and received >= dt_to):
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
    return result


def _get_mail_impl(entry_id, store_id, max_body_chars) -> dict:
    item = _open_item(entry_id, store_id)
    attachments = _attachment_list(item)
    limit = max(500, min(int(max_body_chars or MAX_BODY_CHARS), MAX_BODY_CHARS))
    return {
        "entry_id": _safe(item, "EntryID", ""),
        "betreff": _safe(item, "Subject", ""),
        "von": _safe(item, "SenderName", ""),
        "von_email": _smtp_of_sender(item),
        "an": _safe(item, "To", ""),
        "cc": _safe(item, "CC", ""),
        "empfangen": _fmt(_dt(_safe(item, "ReceivedTime"))),
        "gesendet": _fmt(_dt(_safe(item, "SentOn"))),
        "wichtigkeit": IMPORTANCE.get(_safe(item, "Importance"), "normal"),
        "kategorien": _safe(item, "Categories", ""),
        "konversation": _safe(item, "ConversationTopic", ""),
        "ordner": _folder_path(_safe(item, "Parent")),
        "anhaenge": attachments,
        "text": _truncate(_safe(item, "Body", ""), limit),
    }


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
    }


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

    # Reihenfolge ist wichtig: Sort -> IncludeRecurrences -> Restrict,
    # sonst werden Serientermine nicht korrekt aufgelöst.
    items = folder.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    flt = (f"[Start] < '{_calendar_filter_date(end)}' AND "
           f"[End] > '{_calendar_filter_date(start)}'")
    restricted = items.Restrict(flt)

    needle = (query or "").strip().lower()
    events = []
    scanned = 0
    item = restricted.GetFirst()
    while item is not None and len(events) < limit and scanned < MAX_SCAN:
        scanned += 1
        s = _dt(_safe(item, "Start"))
        e = _dt(_safe(item, "End"))
        if s and s >= end:
            break  # nach Start sortiert: keine weiteren Treffer
        # Zusätzliche Prüfung in Python, unabhängig vom Datumsformat.
        if s and e and s < end and e > start:
            summary = _event_summary(item, folder)
            haystack = " ".join(str(summary.get(k) or "") for k in
                                ("betreff", "ort", "organisator", "kategorien")).lower()
            if not needle or needle in haystack:
                events.append(summary)
        item = restricted.GetNext()

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
            email = ""
            try:
                email = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            except Exception:
                email = _safe(recipient, "Address", "")
            attendees.append({
                "name": _safe(recipient, "Name", ""),
                "email": email,
                "typ": RECIPIENT_TYPE.get(_safe(recipient, "Type"), ""),
                "antwort": RESPONSE_STATUS.get(_safe(recipient, "MeetingResponseStatus"), ""),
            })
    except Exception:
        pass

    if data["serie"]:
        try:
            pattern = item.GetRecurrencePattern()
            data["serienmuster"] = {
                "beginn": _fmt(_dt(_safe(pattern, "PatternStartDate"))),
                "ende": _fmt(_dt(_safe(pattern, "PatternEndDate"))),
                "ohne_ende": bool(_safe(pattern, "NoEndDate", False)),
            }
        except Exception:
            pass

    limit = max(500, min(int(max_body_chars or MAX_BODY_CHARS), MAX_BODY_CHARS))
    data.update({
        "erforderliche_teilnehmende": _safe(item, "RequiredAttendees", ""),
        "optionale_teilnehmende": _safe(item, "OptionalAttendees", ""),
        "teilnehmende": attendees,
        "kalender": _folder_path(folder),
        "anhaenge": _attachment_list(item),
        "text": _truncate(_safe(item, "Body", ""), limit),
    })
    return data


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
SUPPORTED_HINT = ("Unterstützt: PDF, DOCX, XLSX, PPTX, MSG, EML, ZIP "
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
        "suchen, dann Details mit get_mail bzw. get_event abrufen. Anhänge "
        "mit get_attachment lesen (Nummer 'nr' aus get_mail). Für die "
        "Verfügbarkeit anderer Personen get_free_busy, für gemeinsame freie "
        "Zeiten find_meeting_times verwenden. Entwürfe mit create_draft, "
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
    """
    return await _run(_search_mails_impl, query, folder, include_subfolders,
                      sender, date_from, date_to, unread_only,
                      has_attachments, search_body, limit)


@mcp.tool()
async def get_mail(entry_id: str, store_id: str = "",
                   max_body_chars: int = MAX_BODY_CHARS) -> dict:
    """Liest eine einzelne E-Mail vollständig (Kopfdaten, Text, Liste der
    Anhänge mit Nummer 'nr'). Die Mail wird dabei nicht als gelesen
    markiert. Den Inhalt eines Anhangs liefert get_attachment.

    entry_id / store_id: Werte aus search_mails.
    max_body_chars: Maximale Textlänge.
    """
    return await _run(_get_mail_impl, entry_id, store_id, max_body_chars)


@mcp.tool()
async def get_attachment(entry_id: str, store_id: str = "", index: int = 0,
                         name: str = "", offset: int = 0,
                         max_chars: int = MAX_ATTACHMENT_CHARS):
    """Liest den Inhalt eines Anhangs einer E-Mail oder eines Termins.
    Text wird aus PDF, DOCX, XLSX, PPTX, angehängten Mails (MSG, EML),
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
    """Liest einen Termin mit Teilnehmenden, Antworten und Beschreibung
    (z. B. Teams-Link). Bei Serienterminen wird der Serienmaster mit
    Serienmuster geliefert.

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
