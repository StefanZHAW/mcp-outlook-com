# mcp-outlook-com

MCP-Server des Teams Enterprise Applications für **Claude Desktop**: lesender Zugriff auf E-Mails, Kalender und Anhänge im **klassischen Outlook** über das COM-Objektmodell.

Es werden keine Elemente verändert, verschoben, gesendet oder gelöscht. Mails werden beim Lesen nicht als gelesen markiert. Anhänge werden zum Lesen kurz in ein temporäres Verzeichnis gespeichert und danach wieder gelöscht.

## Installation

Installiert wird der Server mit dem Skript `Install-ClaudeMcp.ps1` aus [claude-desktop-install-mcp](https://gitlab.zhaw.ch/enterprise-applications/claude-desktop-install-mcp) (Komponente `outlook`). Das Skript holt mit Git den neuesten Stand von `main` aus diesem Repository, installiert die Abhängigkeiten aus `requirements.txt` und legt den Server unter `%USERPROFILE%\.mcp\outlook\outlook_mcp_server.py` ab.

**Jeder Push auf `main` gelangt bei der nächsten Ausführung des Installationsskripts auf alle Rechner.** Änderungen deshalb vorher testen.

## Werkzeuge

| Werkzeug | Zweck |
|---|---|
| `list_folders` | E-Mail-Ordner aller Postfächer mit Anzahl (ungelesener) Elemente |
| `search_mails` | Mails suchen (Begriff, Ordner, Absender, Zeitraum, ungelesen, mit Anhang) |
| `get_mail` | Eine Mail vollständig lesen, mit Liste der Anhänge |
| `get_attachment` | Inhalt eines Anhangs: Text aus PDF, DOCX, XLSX, PPTX, MSG, EML, HTML und Textdateien, Inhaltsverzeichnis von ZIP, Bilder direkt (bei Bedarf verkleinert). Lange Texte in Teilen (`offset`). |
| `list_calendars` | Kalenderordner aller Postfächer |
| `list_events` | Termine in einem Zeitraum, Serientermine aufgelöst, auch freigegebene Kalender |
| `get_event` | Einen Termin mit Teilnehmenden, Antworten und Beschreibung lesen |

## Konfiguration

Über Umgebungsvariablen im Abschnitt `env` des Eintrags in `claude_desktop_config.json`:

| Variable | Standard | Bedeutung |
|---|---|---|
| `OUTLOOK_MCP_MAX_RESULTS` | 50 | Höchstzahl Treffer pro Suche |
| `OUTLOOK_MCP_MAX_BODY_CHARS` | 8000 | Textlänge bei `get_mail` und `get_event` |
| `OUTLOOK_MCP_MAX_SCAN` | 5000 | Höchstzahl durchsuchter Elemente |
| `OUTLOOK_MCP_MAX_CALENDAR_DAYS` | 366 | Längster Zeitraum bei `list_events` |
| `OUTLOOK_MCP_PREVIEW_CHARS` | 200 | Länge der Vorschau in Suchtreffern |
| `OUTLOOK_MCP_MAX_ATTACHMENT_BYTES` | 30000000 | Grösster lesbarer Anhang |
| `OUTLOOK_MCP_MAX_ATTACHMENT_CHARS` | 30000 | Textmenge pro Aufruf von `get_attachment` |
| `OUTLOOK_MCP_MAX_IMAGE_BYTES` | 3500000 | Grössere Bilder werden verkleinert |
| `OUTLOOK_MCP_MAX_IMAGE_EDGE` | 1568 | Längste Kante verkleinerter Bilder (Pixel) |
| `OUTLOOK_MCP_DATE_FORMAT` | leer | Datumsformat für Kalenderfilter, leer = kurzes Datumsformat von Windows. Bei Problemen z. B. `%d.%m.%Y %H:%M` |

Beispiel und Hinweise zum Eintragen: siehe README von [claude-desktop-install-mcp](https://gitlab.zhaw.ch/enterprise-applications/claude-desktop-install-mcp).

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
