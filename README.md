# mcp-outlook-com

MCP-Server des Teams Enterprise Applications für **Claude Desktop**: lesender Zugriff auf E-Mails, Kalender und Anhänge im **klassischen Outlook** über das COM-Objektmodell, dazu das Anlegen von E-Mail-Entwürfen.

Schreibend werden nur Entwürfe angelegt (neue Mail, Antwort, Weiterleitung) und ungesendete Entwürfe im Ordner Entwürfe geändert. Es wird nichts gesendet, verschoben oder gelöscht; dafür gibt es bewusst kein Werkzeug. Mails werden beim Lesen nicht als gelesen markiert. Anhänge werden zum Lesen kurz in ein temporäres Verzeichnis gespeichert und danach wieder gelöscht.

## Installation

Installiert wird der Server mit dem Skript `Install-ClaudeMcp.ps1` aus [mcp-install-for-claude-desktop](https://gitlab.zhaw.ch/enterprise-applications/mcp-install-for-claude-desktop) (Komponente `outlook`). Das Skript holt mit Git den neuesten Stand von `main` aus diesem Repository, installiert die Abhängigkeiten aus `requirements.txt` und legt den Server unter `%USERPROFILE%\.mcp\outlook\outlook_mcp_server.py` ab.

**Jeder Push auf `main` gelangt bei der nächsten Ausführung des Installationsskripts auf alle Rechner.** Änderungen deshalb vorher testen.

## Werkzeuge

| Werkzeug | Zweck |
|---|---|
| `list_folders` | E-Mail-Ordner aller Postfächer mit Anzahl (ungelesener) Elemente |
| `search_mails` | Mails suchen (Begriff, Ordner, Absender, Zeitraum, ungelesen, mit Anhang) |
| `get_mail` | Eine Mail vollständig lesen, mit Liste der Anhänge |
| `get_attachment` | Inhalt eines Anhangs: Text aus PDF, DOCX, XLSX, PPTX, MSG, EML, HTML und Textdateien, Inhaltsverzeichnis von ZIP, Bilder direkt (bei Bedarf verkleinert). Lange Texte in Teilen (`offset`). |
| `list_calendars` | Kalenderordner aller Postfächer |
| `list_events` | Termine in einem Zeitraum, Serientermine aufgelöst, auch freigegebene Kalender (mit Leserechten) |
| `get_event` | Einen Termin mit Teilnehmenden, Antworten und Beschreibung lesen |
| `get_free_busy` | Belegte Zeiten von Personen oder Räumen (beschäftigt, mit Vorbehalt, abwesend, anderswo tätig), ohne Betreff und Ort |
| `find_meeting_times` | Zeitfenster, in denen alle Teilnehmenden frei sind, mit Mindestdauer und Tageszeitfenster |
| `create_draft` | Neue Mail als Entwurf, mit Standardsignatur |
| `create_reply_draft` | Antwort bzw. Allen antworten als Entwurf, Text über Signatur und Zitat |
| `create_forward_draft` | Weiterleitung als Entwurf, mit den Anhängen der Originalmail |
| `update_draft` | Betreff, Empfänger und Text eines ungesendeten Entwurfs im Ordner Entwürfe ändern |

### Entwürfe

Entwürfe werden mit `Save` im Ordner Entwürfe gespeichert, nie gesendet. Gesendet wird nur von Hand in Outlook. Mit `open_in_outlook` öffnet sich der Entwurf zusätzlich in einem Outlook-Fenster zum Prüfen.

- Der Text (`body`) wird als HTML eingefügt, Zeilenumbrüche bleiben erhalten. Mit `body_is_html` kann ein HTML-Fragment übergeben werden.
- Die Schrift (Standard Aptos 12 pt) wird direkt an jedem Absatz gesetzt. Sonst gälte die Formatvorlage der Mail, die bei Antworten von der Originalmail stammt und je nach Outlook-Version abweicht. Eigene Schriftangaben in einem HTML-Fragment haben Vorrang.
- Die Standardsignatur fügt Outlook selbst ein, wie beim Verfassen einer Mail (Signatur für neue Mails bzw. für Antworten und Weiterleitungen).
- `update_draft` bearbeitet nur ungesendete Mails im Standardordner Entwürfe eines Postfachs. Mit `body_mode` = `voranstellen` (Standard) kommt der Text über den bisherigen Inhalt, mit `ersetzen` wird der ganze Inhalt inklusive Signatur und Zitat ersetzt.
- Die Originalmail wird beim Antworten oder Weiterleiten nicht verändert.

### Freigegebene Kalender und Frei/Gebucht

`list_events` mit `shared_mailbox` liest die Termine aus dem Kalenderordner der anderen Person. Dafür braucht es eine Freigabe mit Leserechten («Alle Details» bzw. «Prüfer»). Bei der Standardfreigabe «Frei/Gebucht», auch mit Titel und Ort, liefert Outlook dort keine Termine, obwohl die Kalenderansicht in Outlook belegte Zeiten zeigt: Diese Ansicht stammt aus dem Verfügbarkeitsdienst von Exchange, nicht aus dem Kalenderordner.

`get_free_busy` und `find_meeting_times` nutzen deshalb `Recipient.FreeBusy` wie die Terminplanung in Outlook und funktionieren mit jeder Frei/Gebucht-Freigabe innerhalb der Organisation. Externe Adressen werden nicht unterstützt.

Zu beachten:
- Die Auflösung ist das gewählte Raster (`interval_minutes`, Standard 15 Minuten). Ein Termin von 08:15 bis 08:55 belegt den Block 08:15 bis 09:00.
- Ein ganztägiger Eintrag «Anderswo tätig» überdeckt in Outlook alle anderen Termine des Tages. `find_meeting_times` wertet ihn als belegt und weist darauf hin.
- Wie weit in die Zukunft Daten verfügbar sind, bestimmt die Veröffentlichung von Frei/Gebucht in Exchange.

## Konfiguration

Über Umgebungsvariablen im Abschnitt `env` des Eintrags in `claude_desktop_config.json`:

| Variable | Standard | Bedeutung |
|---|---|---|
| `OUTLOOK_MCP_MAX_RESULTS` | 50 | Höchstzahl Treffer pro Suche |
| `OUTLOOK_MCP_MAX_BODY_CHARS` | 8000 | Textlänge bei `get_mail` und `get_event` |
| `OUTLOOK_MCP_MAX_SCAN` | 5000 | Höchstzahl durchsuchter Elemente |
| `OUTLOOK_MCP_MAX_CALENDAR_DAYS` | 366 | Längster Zeitraum bei `list_events` |
| `OUTLOOK_MCP_MAX_FREEBUSY_DAYS` | 62 | Längster Zeitraum bei `get_free_busy` und `find_meeting_times` |
| `OUTLOOK_MCP_PREVIEW_CHARS` | 200 | Länge der Vorschau in Suchtreffern |
| `OUTLOOK_MCP_MAX_ATTACHMENT_BYTES` | 30000000 | Grösster lesbarer Anhang |
| `OUTLOOK_MCP_MAX_ATTACHMENT_CHARS` | 30000 | Textmenge pro Aufruf von `get_attachment` |
| `OUTLOOK_MCP_MAX_IMAGE_BYTES` | 3500000 | Grössere Bilder werden verkleinert |
| `OUTLOOK_MCP_MAX_IMAGE_EDGE` | 1568 | Längste Kante verkleinerter Bilder (Pixel) |
| `OUTLOOK_MCP_DRAFT_FONT` | Aptos | Schriftart für den Text von Entwürfen, leer = Formatvorlage der Mail |
| `OUTLOOK_MCP_DRAFT_FONT_SIZE` | 12 | Schriftgrösse (pt) für den Text von Entwürfen |
| `OUTLOOK_MCP_DATE_FORMAT` | leer | Datumsformat für Kalenderfilter, leer = kurzes Datumsformat von Windows. Bei Problemen z. B. `%d.%m.%Y %H:%M` |

Beispiel und Hinweise zum Eintragen: siehe README von [mcp-install-for-claude-desktop](https://gitlab.zhaw.ch/enterprise-applications/mcp-install-for-claude-desktop).

## Voraussetzungen

- Windows mit dem **klassischen** Outlook (das neue Outlook bietet keinen COM-Zugriff)
- Python 3.12 mit den Paketen aus `requirements.txt`
- Ausführung als normaler Benutzer, nicht mit Administratorrechten

## Entwicklung

Lokal testen, ohne Claude Desktop:

```powershell
uv venv --python 3.12
uv pip install -r requirements.txt
.venv\Scripts\python.exe -m py_compile outlook_mcp_server.py
```

Abhängigkeiten nur in `requirements.txt` ändern; das Installationsskript übernimmt sie von dort. `mcp` bleibt auf 1.x begrenzt, da Version 2 die API ändern kann.
