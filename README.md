# mcp-outlook-com

MCP-Server des Teams Enterprise Applications für **Claude Desktop**: lesender Zugriff auf E-Mails, Kalender und Anhänge im **klassischen Outlook** über das COM-Objektmodell, dazu das Anlegen von E-Mail-Entwürfen und das Vorbereiten von Besprechungseinladungen.

Schreibend werden nur Entwürfe angelegt (neue Mail, Antwort, Weiterleitung); bestehende Entwürfe werden nicht geändert. Besprechungseinladungen werden vorbefüllt in einem Outlook-Fenster geöffnet, aber nicht gespeichert. Es wird nichts gesendet, verschoben oder gelöscht; dafür gibt es bewusst kein Werkzeug. Mails werden beim Lesen nicht als gelesen markiert. Anhänge werden zum Lesen kurz in ein temporäres Verzeichnis gespeichert und danach wieder gelöscht.

## Installation

Installiert wird der Server mit dem Skript `Install-ClaudeMcp.ps1` aus [mcp-install-for-claude-desktop](https://gitlab.zhaw.ch/enterprise-applications/mcp-install-for-claude-desktop) (Komponente `outlook`). Das Skript lädt das Paket (Wheel) eines [Releases](https://github.com/StefanZHAW/mcp-outlook-com/releases) in der dort fest eingetragenen Version herunter, prüft die Prüfsumme (SHA256) und installiert es mit seinen Abhängigkeiten in `%USERPROFILE%\.mcp\outlook\.venv`. Git wird dafür nicht benötigt.

Ein Push auf `main` ändert an installierten Rechnern nichts. Neue Stände kommen erst mit einem Release und der angepassten Version im Installationsskript dorthin.

Von Hand, z. B. zum Testen:

```powershell
pip install https://github.com/StefanZHAW/mcp-outlook-com/releases/download/v<Version>/mcp_outlook_com-<Version>-py3-none-any.whl
```

Gestartet wird der Server mit `mcp-outlook-com.exe` aus der virtuellen Umgebung oder mit `python -m mcp_outlook_com`.

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
| `open_meeting_draft` | Besprechungseinladung vorbefüllt in einem Outlook-Fenster öffnen, ohne zu speichern |

### Entwürfe

Entwürfe werden mit `Save` im Ordner Entwürfe gespeichert, nie gesendet. Gesendet wird nur von Hand in Outlook. Mit `open_in_outlook` öffnet sich der Entwurf zusätzlich in einem Outlook-Fenster zum Prüfen.

- Der Text (`body`) wird als HTML eingefügt, Zeilenumbrüche bleiben erhalten. Mit `body_is_html` kann ein HTML-Fragment übergeben werden.
- Die Schrift (Standard Aptos 12 pt) wird direkt an jedem Absatz gesetzt. Sonst gälte die Formatvorlage der Mail, die bei Antworten von der Originalmail stammt und je nach Outlook-Version abweicht. Eigene Schriftangaben in einem HTML-Fragment haben Vorrang.
- Die Standardsignatur fügt Outlook selbst ein, wie beim Verfassen einer Mail (Signatur für neue Mails bzw. für Antworten und Weiterleitungen).
- Einmal angelegte Entwürfe ändert der Server nicht mehr. Korrekturen in Outlook vornehmen oder einen neuen Entwurf anlegen.
- Die Originalmail wird beim Antworten oder Weiterleiten nicht verändert.

### Besprechungseinladungen

`open_meeting_draft` öffnet eine vorbefüllte Einladung (Betreff, Beginn, Ende bzw. Dauer, Ort, erforderliche und optionale Teilnehmende, Räume, Text) in einem Outlook-Fenster. Sie wird weder gespeichert noch gesendet: Gesendet oder gespeichert wird von Hand im Fenster, wird es ohne Speichern geschlossen, bleibt nichts zurück. Bis dahin blockiert sie keine Zeit im Kalender.

- Ohne Teilnehmende und Räume entsteht ein einfacher Termin.
- Eine Teams-Besprechung wird im Fenster über die Schaltfläche hinzugefügt (oder automatisch, wenn das in Teams für alle Besprechungen eingestellt ist).
- Der Text wird als reiner Text übernommen.
- Eine geöffnete Einladung kann der Server nicht mehr ändern, da sie keine `EntryID` hat. Korrekturen im Fenster vornehmen oder eine neue Einladung öffnen.

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
- Python 3.12 (die Abhängigkeiten installiert das Paket mit)
- Ausführung als normaler Benutzer, nicht mit Administratorrechten

## Entwicklung

Lokal testen, ohne Claude Desktop:

```powershell
uv venv --python 3.12
uv pip install -e .
.venv\Scripts\python.exe -c "import mcp_outlook_com.server"
```

Der Code liegt in `mcp_outlook_com/server.py`. Abhängigkeiten stehen in `pyproject.toml` (`dependencies`); das Paket bringt sie bei der Installation mit. `mcp` bleibt auf 1.x begrenzt, da Version 2 die API ändern kann.

Bei jedem Push auf `main` und jedem Pull-Request baut der Workflow `CI` das Wheel und prüft unter Windows, dass es sich installieren und importieren lässt.

## Releases

1. Version in `pyproject.toml` erhöhen (z. B. `1.0.0` → `1.1.0`), committen und auf `main` pushen.
2. Tag setzen und pushen:

   ```powershell
   git tag v1.1.0
   git push origin v1.1.0
   ```

3. Der Workflow `Release` prüft, dass Tag und Version übereinstimmen, baut das Wheel und erstellt den [Release](https://github.com/StefanZHAW/mcp-outlook-com/releases) mit dem Wheel, der Prüfsumme (SHA256) und automatisch erzeugten Release Notes. Die Release Notes lassen sich danach auf GitHub ergänzen.
4. In `Install-ClaudeMcp.ps1` (mcp-install-for-claude-desktop) `$OutlookVersion` und `$OutlookWheelSha256` auf die Werte aus dem Release setzen. Erst damit kommt die neue Version auf die Rechner.
