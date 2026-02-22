"""
Outlook MCP Server

Exposes Outlook emails and calendar via FastMCP so an AI can search and read
emails and calendar events directly.
Uses win32com.client COM automation to connect to a running Outlook instance.

Usage:
    python server.py                       # run as MCP server (stdio)
    fastmcp dev server.py                  # interactive dev mode
"""

import re
import hashlib
import html as html_mod
import base64
import os
import pythoncom
import win32com.client
from collections import OrderedDict
from datetime import datetime, timedelta
from fastmcp import FastMCP
from mcp.types import Icon

def _load_icon(filename: str) -> Icon:
    icons_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images")
    with open(os.path.join(icons_dir, filename), "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return Icon(src=f"data:image/png;base64,{b64}", mimeType="image/png")

_icon_server = _load_icon("mcp-server.png")
_icon_list_folders = _load_icon("folder-library.png")
_icon_search_emails = _load_icon("mail.png")
_icon_search_calendar = _load_icon("calendar.png")
_icon_read_item = _load_icon("mail-read.png")

mcp = FastMCP("Outlook", icons=[_icon_server], instructions=(
    "Search and read Outlook emails and calendar events. "
    "Use list_folders to discover available mail stores/folders. "
    "Use search_emails to find emails by date, sender, recipient, or keyword. "
    "Use search_calendar to find calendar events in a date range. "
    "Use read_item with the short id from search results to get full content "
    "(works for both emails and calendar events). "
    "Recent emails are in the primary mailbox; older emails are in the Online Archive. "
    "Use include_archive=true to also search archived emails."
))

# ---------------------------------------------------------------------------
# Outlook folder-type constants (OlDefaultFolders enumeration)
# ---------------------------------------------------------------------------
OL_FOLDER_DELETED = 3
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
OL_FOLDER_DRAFTS = 16
OL_FOLDER_JUNK = 23

FOLDER_MAP = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT,
    "drafts": OL_FOLDER_DRAFTS,
    "deleted": OL_FOLDER_DELETED,
    "junk": OL_FOLDER_JUNK,
}


# ---------------------------------------------------------------------------
# COM helpers
# ---------------------------------------------------------------------------

def _get_namespace():
    """Return a fresh MAPI namespace (must be called after CoInitialize)."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def _get_default_folder(namespace, folder_name: str):
    """Resolve a friendly folder name to a COM folder object."""
    folder_id = FOLDER_MAP.get(folder_name.lower())
    if folder_id is None:
        raise ValueError(f"Unknown folder '{folder_name}'. Use: inbox, sent, drafts, deleted, junk")
    return namespace.GetDefaultFolder(folder_id)


def _find_archive_folder(namespace, folder_name: str):
    """Find a folder inside the Online Archive store."""
    archive_folder_names = {
        "inbox": "Inbox",
        "sent": "Sent Items",
        "drafts": "Drafts",
        "deleted": "Deleted Items",
        "junk": "Junk Email",
    }
    target = archive_folder_names.get(folder_name.lower(), folder_name)

    for i in range(1, namespace.Stores.Count + 1):
        store = namespace.Stores.Item(i)
        if "Online Archive" in store.DisplayName:
            root = store.GetRootFolder()
            for j in range(1, root.Folders.Count + 1):
                folder = root.Folders.Item(j)
                if folder.Name == target:
                    return folder
    return None


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def strip_html(html: str) -> str:
    """Basic HTML to plain text conversion."""
    if not html:
        return ""
    text = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</(p|div|tr|li|h[1-6])>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = html_mod.unescape(text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


MAX_BODY_LENGTH = 50_000  # ~50k chars ≈ ~12k tokens


def _get_body(item, truncate: bool = True) -> str:
    """Get plain-text body, falling back to HTMLBody with strip_html."""
    body = getattr(item, 'Body', '') or ''
    if not body.strip():
        html_body = getattr(item, 'HTMLBody', '') or ''
        if html_body:
            body = strip_html(html_body)
    body = body.strip()
    if truncate and len(body) > MAX_BODY_LENGTH:
        body = body[:MAX_BODY_LENGTH] + "\n\n[body truncated — use read_item with full_body=true to get the complete text]"
    return body


# ---------------------------------------------------------------------------
# Sender cleanup (used by read_item for full Exchange DN resolution)
# ---------------------------------------------------------------------------

def _clean_sender(item) -> str:
    """Return a clean 'Name <email>' sender string, resolving Exchange DNs."""
    name = getattr(item, 'SenderName', '') or ''
    email = getattr(item, 'SenderEmailAddress', '') or ''

    if email.upper().startswith('/O=') or '/CN=' in email.upper():
        try:
            smtp = item.Sender.GetExchangeUser().PrimarySmtpAddress
            if smtp:
                email = smtp
            else:
                return name
        except Exception:
            return name

    if email and email != name:
        return f"{name} <{email}>"
    return name


# ---------------------------------------------------------------------------
# Short ID cache
# ---------------------------------------------------------------------------

_id_cache: OrderedDict[str, str] = OrderedDict()  # short_id -> real entry_id
MAX_CACHE_SIZE = 500


def _hash_id(entry_id: str) -> str:
    """Convert an entry_id to a 4-char base36 hash."""
    digest = hashlib.sha256(entry_id.encode()).digest()
    num = int.from_bytes(digest[:4], 'big')
    chars = '0123456789abcdefghijklmnopqrstuvwxyz'
    result = []
    for _ in range(4):
        num, rem = divmod(num, 36)
        result.append(chars[rem])
    return ''.join(result)


def _assign_short_id(entry_id: str) -> str:
    """Assign a deterministic 4-char base36 ID and cache the mapping."""
    short = _hash_id(entry_id)

    existing = _id_cache.get(short)
    if existing and existing != entry_id:
        for suffix in range(1, 100):
            candidate = f"{short}{suffix}"
            existing = _id_cache.get(candidate)
            if not existing or existing == entry_id:
                short = candidate
                break

    if short not in _id_cache and len(_id_cache) >= MAX_CACHE_SIZE:
        for _ in range(len(_id_cache) // 2):
            _id_cache.popitem(last=False)

    _id_cache[short] = entry_id
    _id_cache.move_to_end(short)
    return short


def _resolve_id(id_str: str) -> str:
    """Resolve a short ID to a real entry_id, or pass through if already a full ID."""
    real_id = _id_cache.get(id_str)
    if real_id is not None:
        _id_cache.move_to_end(id_str)
        return real_id
    return id_str


# ---------------------------------------------------------------------------
# DASL filter builder
# ---------------------------------------------------------------------------

PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"


def _build_dasl_filter(query: str, date_from: str, date_to: str,
                       sender: str, to: str, is_read: bool | None = None) -> str:
    """Build a unified DASL filter string for Folder.GetTable().

    All conditions use DASL syntax so they combine in a single filter.
    Returns an empty string if no filters are needed.
    """
    parts = []

    if date_from:
        dt = datetime.strptime(date_from, "%Y-%m-%d")
        parts.append(
            f"\"urn:schemas:httpmail:date\" >= '{dt.strftime('%m/%d/%Y 00:00')}'")
    if date_to:
        dt = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
        parts.append(
            f"\"urn:schemas:httpmail:date\" < '{dt.strftime('%m/%d/%Y 00:00')}'")

    if query:
        safe = query.replace("'", "''")
        parts.append(
            f"(\"urn:schemas:httpmail:subject\" ci_phrasematch '{safe}' "
            f"OR \"urn:schemas:httpmail:textdescription\" ci_phrasematch '{safe}')")

    if sender:
        words = sender.replace("'", "''").split()
        if len(words) == 1:
            w = words[0]
            parts.append(
                f"(\"urn:schemas:httpmail:sendername\" LIKE '%{w}%' "
                f"OR \"urn:schemas:httpmail:fromemail\" LIKE '%{w}%')")
        else:
            # Require each word to appear in sendername independently
            word_parts = [f"\"urn:schemas:httpmail:sendername\" LIKE '%{w}%'"
                          for w in words]
            parts.append(f"({' AND '.join(word_parts)})")

    if to:
        words = to.replace("'", "''").split()
        word_parts = [f"\"urn:schemas:httpmail:displayto\" LIKE '%{w}%'"
                      for w in words]
        parts.append(f"({' AND '.join(word_parts)})")

    if is_read is not None:
        parts.append(f"\"urn:schemas:httpmail:read\" = {1 if is_read else 0}")

    if not parts:
        return ""
    return "@SQL=" + " AND ".join(parts)


# ---------------------------------------------------------------------------
# Table-based search
# ---------------------------------------------------------------------------

def _search_folder(folder, filter_str: str, max_results: int,
                    earliest_first: bool = False) -> list[dict]:
    """Search a folder using GetTable() and return lightweight summary dicts.

    GetTable avoids loading full COM MailItem objects — it fetches only the
    requested columns directly from the store, which is significantly faster
    for listing/browsing than Items.Restrict + per-item property access.
    """
    if max_results <= 0:
        return []

    table = folder.GetTable(filter_str or "", 0)
    table.Columns.RemoveAll()
    table.Columns.Add("EntryID")
    table.Columns.Add("Subject")
    table.Columns.Add("SentOn")
    table.Columns.Add("SenderName")
    table.Columns.Add(PR_SENDER_EMAIL)
    table.Columns.Add("To")
    table.Columns.Add("CC")
    table.Columns.Add("MessageClass")
    table.Sort("SentOn", not earliest_first)

    results = []
    while not table.EndOfTable and len(results) < max_results:
        try:
            row = table.GetNextRow()

            msg_class = row("MessageClass") or ""
            if msg_class and not msg_class.startswith("IPM.Note"):
                continue

            entry_id = row("EntryID") or ""
            sent_on = row("SentOn")
            sender_name = row("SenderName") or ""
            sender_email = row(PR_SENDER_EMAIL) or ""

            # Fast sender formatting — skips Exchange DN resolution
            if (sender_email
                    and not sender_email.upper().startswith('/O=')
                    and sender_email != sender_name):
                sender = f"{sender_name} <{sender_email}>"
            else:
                sender = sender_name

            result = {
                "id": _assign_short_id(entry_id),
                "date": sent_on.strftime('%Y-%m-%d %H:%M') if sent_on else "unknown",
                "subject": row("Subject") or "(no subject)",
                "sender": sender,
                "to": row("To") or "",
            }

            cc = row("CC") or ""
            if cc:
                result["cc"] = cc

            results.append(result)
        except Exception:
            continue

    return results


# ---------------------------------------------------------------------------
# Full item extraction (for read_item)
# ---------------------------------------------------------------------------

RESPONSE_STATUS_MAP = {
    0: "none",
    1: "organized",
    2: "tentative",
    3: "accepted",
    4: "declined",
    5: "not_responded",
}

BUSY_STATUS_MAP = {
    0: "free",
    1: "tentative",
    2: "busy",
    3: "out_of_office",
    4: "working_elsewhere",
}

_RESPONSE_LOOKUP = {v: k for k, v in RESPONSE_STATUS_MAP.items()}


def _extract_mail(item, truncate: bool = True) -> dict:
    """Extract full details from a COM mail item."""
    sent_on = getattr(item, 'SentOn', None)
    body = _get_body(item, truncate=truncate)

    result = {
        "date": sent_on.strftime('%Y-%m-%d %H:%M') if sent_on else 'unknown',
        "subject": getattr(item, 'Subject', '(no subject)') or '(no subject)',
        "sender": _clean_sender(item),
        "to": getattr(item, 'To', '') or '',
        "body": body,
    }

    cc = getattr(item, 'CC', '') or ''
    if cc:
        result["cc"] = cc

    importance = getattr(item, 'Importance', 1)  # 0=Low, 1=Normal, 2=High
    if importance != 1:
        result["importance"] = {0: "Low", 2: "High"}.get(importance, str(importance))

    categories = getattr(item, 'Categories', '') or ''
    if categories:
        result["categories"] = categories

    attachments = []
    try:
        for i in range(1, item.Attachments.Count + 1):
            att = item.Attachments.Item(i)
            attachments.append(att.FileName)
    except Exception:
        pass
    if attachments:
        result["attachments"] = attachments

    return result


def _extract_calendar(item, truncate: bool = True) -> dict:
    """Extract full details from a COM AppointmentItem."""
    start = getattr(item, 'Start', None)
    end = getattr(item, 'End', None)
    body = _get_body(item, truncate=truncate)

    result = {
        "subject": getattr(item, 'Subject', '(no subject)') or '(no subject)',
        "start": start.strftime('%Y-%m-%d %H:%M') if start else 'unknown',
        "end": end.strftime('%Y-%m-%d %H:%M') if end else 'unknown',
        "duration": getattr(item, 'Duration', 0),
        "location": getattr(item, 'Location', '') or '',
        "organizer": getattr(item, 'Organizer', '') or '',
    }

    required = getattr(item, 'RequiredAttendees', '') or ''
    optional = getattr(item, 'OptionalAttendees', '') or ''
    if required:
        result["required_attendees"] = required
    if optional:
        result["optional_attendees"] = optional

    response = getattr(item, 'ResponseStatus', 0)
    result["response"] = RESPONSE_STATUS_MAP.get(response, str(response))

    busy = getattr(item, 'BusyStatus', 2)
    result["busy_status"] = BUSY_STATUS_MAP.get(busy, str(busy))

    if getattr(item, 'IsRecurring', False):
        result["is_recurring"] = True

    if body:
        result["body"] = body

    categories = getattr(item, 'Categories', '') or ''
    if categories:
        result["categories"] = categories

    attachments = []
    try:
        for i in range(1, item.Attachments.Count + 1):
            att = item.Attachments.Item(i)
            attachments.append(att.FileName)
    except Exception:
        pass
    if attachments:
        result["attachments"] = attachments

    return result


# ---------------------------------------------------------------------------
# MCP Tools
# ---------------------------------------------------------------------------

@mcp.tool(icons=[_icon_list_folders])
def list_folders() -> list[dict]:
    """List all Outlook stores and their top-level folders with item counts."""
    pythoncom.CoInitialize()
    try:
        namespace = _get_namespace()
        result = []
        for i in range(1, namespace.Stores.Count + 1):
            store = namespace.Stores.Item(i)
            store_info = {"store_name": store.DisplayName, "folders": []}
            try:
                root = store.GetRootFolder()
                for j in range(1, root.Folders.Count + 1):
                    folder = root.Folders.Item(j)
                    try:
                        count = folder.Items.Count
                    except Exception:
                        count = -1
                    if count != 0:
                        store_info["folders"].append({
                            "name": folder.Name,
                            "count": count,
                        })
            except Exception as e:
                store_info["error"] = str(e)
            result.append(store_info)
        return result
    finally:
        pythoncom.CoUninitialize()


@mcp.tool(icons=[_icon_search_emails])
def search_emails(
    query: str = "",
    folder: str = "inbox",
    sender: str = "",
    to: str = "",
    date_from: str = "",
    date_to: str = "",
    include_archive: bool = False,
    is_read: bool | None = None,
    earliest_first: bool = False,
    max_results: int = 20,
) -> dict:
    """Search Outlook emails with filters. Returns summaries with short id for read_item.

    Sorted newest-first by default. If count equals max_results, there may be more results.
    Results do not include body — use read_item for full content.

    Args:
        query: Phrase match in subject/body (words must appear together as a phrase).
        folder: "inbox", "sent", "drafts", "deleted", or "junk".
        sender: Filter by sender display name (partial match).
        to: Filter by recipient display name (partial match).
        date_from: Start date YYYY-MM-DD (inclusive).
        date_to: End date YYYY-MM-DD (inclusive).
        include_archive: Also search the Online Archive.
        is_read: Filter by read status. True = read only, False = unread only.
        earliest_first: Sort earliest-first instead of latest-first.
        max_results: Max results to return (default 20).
    """
    if date_to and not date_from:
        raise ValueError("date_from is required when date_to is specified.")

    pythoncom.CoInitialize()
    try:
        namespace = _get_namespace()
        fname = folder.lower()
        filter_str = _build_dasl_filter(query, date_from, date_to, sender, to, is_read)
        results = []

        # Search archive first when earliest_first, since it has older emails
        archive_first = include_archive and earliest_first

        if archive_first:
            try:
                archive_folder = _find_archive_folder(namespace, fname)
                if archive_folder:
                    results.extend(
                        _search_folder(archive_folder, filter_str,
                                       max_results, earliest_first))
            except Exception:
                pass

        # Primary mailbox
        if len(results) < max_results:
            try:
                primary_folder = _get_default_folder(namespace, fname)
                results.extend(_search_folder(primary_folder, filter_str,
                                              max_results - len(results), earliest_first))
            except Exception:
                pass

        # Online Archive (when newest-first)
        if include_archive and not archive_first and len(results) < max_results:
            try:
                archive_folder = _find_archive_folder(namespace, fname)
                if archive_folder:
                    results.extend(
                        _search_folder(archive_folder, filter_str,
                                       max_results - len(results), earliest_first)
                    )
            except Exception:
                pass

        trimmed = results[:max_results]
        return {"count": len(trimmed), "max_results": max_results, "results": trimmed}
    finally:
        pythoncom.CoUninitialize()


@mcp.tool(icons=[_icon_search_calendar])
def search_calendar(
    date_from: str = "",
    date_to: str = "",
    query: str = "",
    response: str = "",
    earliest_first: bool = True,
    max_results: int = 20,
) -> dict:
    """Search Outlook calendar events in a date range. Returns summaries with short id for read_item.

    Args:
        date_from: Start date YYYY-MM-DD (inclusive). Defaults to today.
        date_to: End date YYYY-MM-DD (inclusive). Defaults to date_from (single day).
        query: Filter by subject (partial match).
        response: Filter by response (accepted, tentative, declined, organized, none, not_responded).
        earliest_first: Sort earliest-first (default true, showing soonest events first).
        max_results: Max results to return (default 20).
    """
    if date_to and not date_from:
        raise ValueError("date_from is required when date_to is specified.")
    if response and response.lower() not in _RESPONSE_LOOKUP:
        raise ValueError(f"Unknown response '{response}'. Use: {', '.join(_RESPONSE_LOOKUP)}")

    pythoncom.CoInitialize()
    try:
        namespace = _get_namespace()
        folder = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)

        # Default to today if no date_from
        if not date_from:
            date_from = datetime.now().strftime("%Y-%m-%d")
        if not date_to:
            date_to = date_from

        start_dt = datetime.strptime(date_from, "%Y-%m-%d")
        end_dt = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)

        # Must Sort ascending THEN IncludeRecurrences THEN Restrict — order
        # matters for recurring event expansion (descending breaks it).
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        restrict_str = (
            f"[Start] >= '{start_dt.strftime('%m/%d/%Y 00:00')}'"
            f" AND [Start] < '{end_dt.strftime('%m/%d/%Y 00:00')}'"
        )
        restrict_str += " AND [MeetingStatus] <> 5 AND [MeetingStatus] <> 7"
        items = items.Restrict(restrict_str)

        results = []
        item = items.GetFirst()
        while item and (not earliest_first or len(results) < max_results):
            try:
                resp = getattr(item, 'ResponseStatus', 0)
                if response and RESPONSE_STATUS_MAP.get(resp) != response.lower():
                    item = items.GetNext()
                    continue

                # Filter by subject if query given
                subject = getattr(item, 'Subject', '') or '(no subject)'
                if query and query.lower() not in subject.lower():
                    item = items.GetNext()
                    continue

                start = getattr(item, 'Start', None)
                end = getattr(item, 'End', None)
                entry_id = getattr(item, 'EntryID', '') or ''

                result = {
                    "id": _assign_short_id(entry_id),
                    "date": start.strftime('%Y-%m-%d') if start else 'unknown',
                    "start": start.strftime('%H:%M') if start else 'unknown',
                    "end": end.strftime('%H:%M') if end else 'unknown',
                    "subject": subject,
                    "location": getattr(item, 'Location', '') or '',
                    "organizer": getattr(item, 'Organizer', '') or '',
                    "response": RESPONSE_STATUS_MAP.get(resp, str(resp)),
                }

                busy = getattr(item, 'BusyStatus', 2)
                if busy != 2:  # Only include if not the default "busy"
                    result["busy_status"] = BUSY_STATUS_MAP.get(busy, str(busy))

                if getattr(item, 'IsRecurring', False):
                    result["is_recurring"] = True

                results.append(result)
            except Exception:
                pass
            item = items.GetNext()

        if not earliest_first:
            results.reverse()
        results = results[:max_results]
        return {"count": len(results), "max_results": max_results, "results": results}
    finally:
        pythoncom.CoUninitialize()


@mcp.tool(icons=[_icon_read_item])
def read_item(entry_id: str, full_body: bool = False) -> dict:
    """Read the full content of an email or calendar event by its ID.

    Args:
        entry_id: The short id from search_emails/search_calendar results, or a full EntryID.
        full_body: Return the complete body without truncation (default false).
    """
    pythoncom.CoInitialize()
    try:
        namespace = _get_namespace()
        real_id = _resolve_id(entry_id)
        item = namespace.GetItemFromID(real_id)
        msg_class = getattr(item, 'MessageClass', '') or ''
        if msg_class.startswith('IPM.Appointment'):
            return _extract_calendar(item, truncate=not full_body)
        return _extract_mail(item, truncate=not full_body)
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# MCP Prompts
# ---------------------------------------------------------------------------

@mcp.prompt()
def weekly_summary() -> str:
    """Summarize what I did each day this past work week based on my emails and calendar."""
    return (
        "Use search_calendar and search_emails (both inbox and sent) for the past "
        "Monday through Friday. For each day, summarize my meetings and notable "
        "email activity. End with any themes or highlights for the week."
    )


@mcp.prompt()
def agenda() -> str:
    """Show my agenda for today with relevant email context."""
    return (
        "Search my calendar for today. For each meeting, search for recent emails "
        "involving the organizer or related to the meeting subject. Present my "
        "schedule in chronological order with any relevant email context that "
        "would help me prepare."
    )


@mcp.prompt()
def next_meeting() -> str:
    """Prep me for my next upcoming meeting."""
    return (
        "Search my calendar for today to find my next upcoming meeting. Then "
        "search for recent emails (past 7 days, both inbox and sent) involving "
        "the attendees or related to the meeting subject. Give me a briefing: "
        "who's attending, what the meeting is about, and any relevant email "
        "threads I should be aware of."
    )


@mcp.prompt()
def unanswered_emails() -> str:
    """Find emails that look like they need a reply from me."""
    return (
        "Search my inbox for the past 5 days, then search my sent folder for "
        "the same period. Compare them to identify inbox emails that appear to "
        "ask me a question or request action, where I haven't sent a reply to "
        "that thread. List them with the sender, date, subject, and a brief "
        "note on what seems to be needed."
    )


@mcp.prompt()
def annual_review() -> str:
    """Analyze the past year of emails for accomplishments, contributions, and praise to support an annual review."""
    return (
        "Help me prepare for my annual performance review. Search my sent "
        "folder and inbox over the past 12 months. First identify my most "
        "frequent contacts to understand key relationships. Then search for "
        "evidence of accomplishments, completed work, praise from others, and "
        "examples of helping or unblocking teammates. Read promising results "
        "for detail. Compile a summary with key accomplishments, recognition "
        "received, and examples of collaboration — include direct quotes where "
        "they strengthen the evidence."
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
