# mcp-outlook-com

MCP server for **Claude Desktop**: read access to emails, calendars and attachments in **classic Outlook** via the COM object model, plus creating email drafts and preparing meeting invitations.

The only write operations are creating drafts (new mail, reply, forward); existing drafts are never modified. Meeting invitations are opened pre-filled in an Outlook window but not saved. Nothing is sent, moved or deleted; there is deliberately no tool for that. Reading a mail does not mark it as read. Attachments are saved briefly to a temporary directory for reading and deleted afterwards.

## Installation

Install the package (wheel) of a [release](https://github.com/StefanZHAW/mcp-outlook-com/releases) into its own virtual environment, e.g. `%USERPROFILE%\.mcp\outlook\.venv`. The dependencies are installed along with it; Git is not required. The SHA256 checksum of the wheel is listed in the release notes.

```powershell
py -3.12 -m venv "$env:USERPROFILE\.mcp\outlook\.venv"
& "$env:USERPROFILE\.mcp\outlook\.venv\Scripts\python.exe" -m pip install https://github.com/StefanZHAW/mcp-outlook-com/releases/download/v<version>/mcp_outlook_com-<version>-py3-none-any.whl
```

Alternatively, [uv](https://docs.astral.sh/uv/) can start the server with `uvx` directly from the wheel, without a dedicated virtual environment (see the example under [Configuration](#configuration)). uv creates the environment on first start and takes it from its cache afterwards.

A push to `main` does not change anything on installed machines. New versions only arrive there with a release and a reinstallation.

Start the server with `mcp-outlook-com.exe` from the virtual environment or with `python -m mcp_outlook_com`.

## Tools

| Tool | Purpose |
|---|---|
| `list_folders` | Mail folders of all mailboxes with the number of (unread) items |
| `search_mails` | Search mails (term, folder, sender, date range, unread, with attachments) |
| `get_mail` | Read a mail in full, with a list of its attachments |
| `get_attachment` | Content of an attachment: text from PDF, DOCX, XLSX, PPTX, MSG, EML, HTML and text files, table of contents of ZIP files, images directly (downscaled if needed). Long texts in chunks (`offset`). |
| `list_calendars` | Calendar folders of all mailboxes |
| `list_events` | Events in a date range, recurring events expanded, including shared calendars (with read permission) |
| `get_event` | Read an event with attendees, responses and description |
| `get_free_busy` | Busy times of people or rooms (busy, tentative, out of office, working elsewhere), without subject and location |
| `find_meeting_times` | Time slots in which all attendees are free, with minimum duration and time-of-day window |
| `create_draft` | New mail as a draft, with the default signature |
| `create_reply_draft` | Reply or reply all as a draft, text above signature and quoted mail |
| `create_forward_draft` | Forward as a draft, with the attachments of the original mail |
| `open_meeting_draft` | Open a pre-filled meeting invitation in an Outlook window without saving it |

### Drafts

Drafts are stored in the Drafts folder with `Save` and never sent. Sending is done manually in Outlook. With `open_in_outlook`, the draft is also opened in an Outlook window for review.

- The text (`body`) is inserted as HTML; line breaks are preserved. With `body_is_html`, an HTML fragment can be passed.
- The font (default Aptos 12 pt) is set directly on each paragraph. Otherwise the mail's style sheet would apply, which for replies comes from the original mail and varies between Outlook versions. Font settings in an HTML fragment take precedence.
- Outlook inserts the default signature itself, just as when composing a mail (signature for new messages or for replies and forwards).
- Once created, drafts are no longer modified by the server. Make corrections in Outlook or create a new draft.
- The original mail is not modified when replying or forwarding.

### Meeting invitations

`open_meeting_draft` opens a pre-filled invitation (subject, start, end or duration, location, required and optional attendees, rooms, text) in an Outlook window. It is neither saved nor sent: sending or saving is done manually in the window; if the window is closed without saving, nothing remains. Until then, it does not block any time in the calendar.

- Without attendees and rooms, a simple appointment is created.
- A Teams meeting is added in the window via the button (or automatically if Teams is set up to do so for all meetings).
- The text is inserted as plain text.
- The server cannot modify an invitation once it is open, since it has no `EntryID`. Make corrections in the window or open a new invitation.

### Shared calendars and free/busy

`list_events` with `shared_mailbox` reads the events from the other person's calendar folder. This requires a share with read permission ("Can view all details" or "Reviewer"). With the default "Can view when I'm busy" share, even with titles and locations, Outlook returns no events there, although the Outlook calendar view shows busy times: that view comes from the Exchange availability service, not from the calendar folder.

`get_free_busy` and `find_meeting_times` therefore use `Recipient.FreeBusy`, like the Outlook Scheduling Assistant, and work with any free/busy share within the organization. External addresses are not supported.

Note:
- The resolution is the chosen grid (`interval_minutes`, default 15 minutes). An event from 08:15 to 08:55 occupies the block from 08:15 to 09:00.
- In Outlook, an all-day "Working elsewhere" entry covers all other events of the day. `find_meeting_times` treats it as busy and points this out.
- How far into the future data is available depends on the free/busy publishing settings in Exchange.

## Configuration

Via environment variables in the `env` section of the entry in `claude_desktop_config.json`:

| Variable | Default | Meaning |
|---|---|---|
| `OUTLOOK_MCP_MAX_RESULTS` | 50 | Maximum number of results per search |
| `OUTLOOK_MCP_MAX_BODY_CHARS` | 8000 | Text length for `get_mail` and `get_event` |
| `OUTLOOK_MCP_MAX_SCAN` | 5000 | Maximum number of items scanned |
| `OUTLOOK_MCP_MAX_CALENDAR_DAYS` | 366 | Longest date range for `list_events` |
| `OUTLOOK_MCP_MAX_FREEBUSY_DAYS` | 62 | Longest date range for `get_free_busy` and `find_meeting_times` |
| `OUTLOOK_MCP_PREVIEW_CHARS` | 200 | Length of the preview in search results |
| `OUTLOOK_MCP_MAX_ATTACHMENT_BYTES` | 30000000 | Largest readable attachment |
| `OUTLOOK_MCP_MAX_ATTACHMENT_CHARS` | 30000 | Amount of text per `get_attachment` call |
| `OUTLOOK_MCP_MAX_IMAGE_BYTES` | 3500000 | Larger images are downscaled |
| `OUTLOOK_MCP_MAX_IMAGE_EDGE` | 1568 | Longest edge of downscaled images (pixels) |
| `OUTLOOK_MCP_DRAFT_FONT` | Aptos | Font for the text of drafts, empty = the mail's style sheet |
| `OUTLOOK_MCP_DRAFT_FONT_SIZE` | 12 | Font size (pt) for the text of drafts |
| `OUTLOOK_MCP_DATE_FORMAT` | empty | Date format for calendar filters, empty = Windows short date format. If there are problems, e.g. `%d.%m.%Y %H:%M` |

Example entry in `claude_desktop_config.json` (adjust the path to your installation):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\<user>\\.mcp\\outlook\\.venv\\Scripts\\mcp-outlook-com.exe",
      "env": {
        "PYTHONUTF8": "1",
        "OUTLOOK_MCP_MAX_RESULTS": "50",
        "OUTLOOK_MCP_MAX_BODY_CHARS": "20000"
      }
    }
  }
}
```

With `uvx` instead of a dedicated installation (`uvx` must be on the `PATH`, otherwise specify the full path; the version is part of the wheel URL):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": [
        "--python", "3.12",
        "--from", "https://github.com/StefanZHAW/mcp-outlook-com/releases/download/v<version>/mcp_outlook_com-<version>-py3-none-any.whl",
        "mcp-outlook-com"
      ]
    }
  }
}
```

After a change, quit Claude Desktop completely and restart it.

## Requirements

- Windows with **classic** Outlook (the new Outlook does not offer COM access)
- Python 3.12 (the package installs its dependencies)
- Run as a normal user, not with administrator rights

## Development

Test locally without Claude Desktop:

```powershell
uv venv --python 3.12
uv pip install -e .
.venv\Scripts\python.exe -c "import mcp_outlook_com.server"
```

The code is in `mcp_outlook_com/server.py`. Dependencies are listed in `pyproject.toml` (`dependencies`); the package brings them along on installation. `mcp` stays limited to 1.x, since version 2 may change the API.

On every push to `main` and every pull request, the `CI` workflow builds the wheel and checks on Windows that it can be installed and imported.

## Releases

1. Bump the version in `pyproject.toml` (e.g. `1.0.0` → `1.1.0`), commit and push to `main`.
2. Create and push a tag:

   ```powershell
   git tag v1.1.0
   git push origin v1.1.0
   ```

3. The `Release` workflow checks that tag and version match, builds the wheel and creates the [release](https://github.com/StefanZHAW/mcp-outlook-com/releases) with the wheel, the SHA256 checksum and automatically generated release notes. The release notes can be extended on GitHub afterwards.

## References

- [Model Context Protocol](https://modelcontextprotocol.io/)
- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk)
- [Outlook object model (VBA reference)](https://learn.microsoft.com/office/vba/api/overview/outlook/object-model)
- [pywin32](https://github.com/mhammond/pywin32)

## License

[MIT](LICENSE)
