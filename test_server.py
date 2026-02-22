"""
Pytest test suite for Outlook MCP Server tools.

Integration tests that call real MCP tools against a running Outlook instance.
No mocking — Outlook must be available on the machine.

Run all tests:
    pytest test_server.py -v

Run a single test:
    pytest test_server.py::test_list_folders -v
"""

import re
import pytest
from server import (
    list_folders, search_emails, search_calendar, read_item,
    _shorten_urls, _id_cache,
)


# ============================================================================
# list_folders
# ============================================================================

def test_list_folders():
    """list_folders returns a list of stores, each with store_name and folders."""
    result = list_folders()
    assert isinstance(result, list)
    assert len(result) > 0
    for store in result:
        assert "store_name" in store
        assert "folders" in store


# ============================================================================
# search_emails
# ============================================================================

def test_search_emails():
    """Default inbox search returns count, max_results, and results with expected keys."""
    result = search_emails()
    assert "count" in result
    assert "max_results" in result
    assert "results" in result
    assert result["count"] > 0
    email = result["results"][0]
    for key in ("id", "date", "subject", "sender", "to"):
        assert key in email, f"Missing key '{key}' in email result"


def test_search_emails_folder():
    """folder='sent' partial match finds Sent Items."""
    result = search_emails(folder="sent")
    assert result["count"] > 0
    email = result["results"][0]
    for key in ("id", "date", "subject", "sender", "to"):
        assert key in email


def test_search_emails_store():
    """store='archive' searches the Online Archive store."""
    result = search_emails(store="archive", folder="inbox")
    assert "count" in result
    assert "results" in result
    # Archive may have 0 results, but the call should succeed without error


def test_search_emails_is_read():
    """is_read=True returns only read emails."""
    result = search_emails(is_read=True, max_results=5)
    assert result["count"] > 0
    # We can't directly verify read status from the summary, but the filter
    # should have been applied — just verify the call succeeds with results.


def test_search_emails_date_range():
    """date_from and date_to narrow results to a date range."""
    result = search_emails(
        date_from="2025-01-01",
        date_to="2025-12-31",
        max_results=5,
    )
    assert "count" in result
    assert "results" in result
    # Verify dates are within range
    for email in result["results"]:
        assert email["date"] >= "2025-01-01"
        assert email["date"] < "2026-01-01"


def test_search_emails_date_validation():
    """date_to without date_from raises an error."""
    with pytest.raises(Exception):
        search_emails(date_to="2025-12-31")


def test_search_emails_earliest_first():
    """earliest_first=True reverses the default sort order."""
    newest_first = search_emails(max_results=5)
    earliest_first = search_emails(max_results=5, earliest_first=True)

    if newest_first["count"] > 1 and earliest_first["count"] > 1:
        # The first result of each should be different (opposite ends)
        assert newest_first["results"][0]["date"] != earliest_first["results"][0]["date"]


# ============================================================================
# search_calendar
# ============================================================================

def test_search_calendar():
    """Calendar search returns events with expected keys."""
    # Search a wide range to ensure we find events
    result = search_calendar(
        date_from="2025-01-01",
        date_to="2026-12-31",
    )
    assert "count" in result
    assert "max_results" in result
    assert "results" in result
    assert result["count"] > 0
    event = result["results"][0]
    for key in ("id", "date", "start", "end", "subject"):
        assert key in event, f"Missing key '{key}' in calendar result"


def test_search_calendar_earliest_first():
    """earliest_first=False reverses the default ascending order."""
    asc = search_calendar(
        date_from="2025-01-01",
        date_to="2026-12-31",
        earliest_first=True,
    )
    desc = search_calendar(
        date_from="2025-01-01",
        date_to="2026-12-31",
        earliest_first=False,
    )
    if asc["count"] > 1 and desc["count"] > 1:
        assert asc["results"][0]["date"] != desc["results"][-1]["date"] or \
               asc["results"][0]["start"] != desc["results"][-1]["start"]


def test_search_calendar_response_filter():
    """response='accepted' filters to only accepted events."""
    result = search_calendar(
        date_from="2025-01-01",
        date_to="2026-12-31",
        response="accepted",
    )
    assert "count" in result
    for event in result["results"]:
        assert event["response"] == "accepted"


def test_search_calendar_bad_response():
    """Invalid response value raises an error."""
    with pytest.raises(Exception):
        search_calendar(response="bogus_status")


# ============================================================================
# read_item
# ============================================================================

def test_read_item():
    """Search for an email, then read its full content by ID."""
    search = search_emails(max_results=1)
    assert search["count"] > 0
    email_id = search["results"][0]["id"]

    item = read_item(entry_id=email_id)
    # read_item returns full email details
    assert "subject" in item
    assert "body" in item
    assert "sender" in item


def test_read_item_url():
    """read_item with a url: prefix resolves to the cached URL."""
    # Seed a long URL into the cache via _shorten_urls
    long_url = "https://example.com/" + "x" * 100
    text = f"Check this link: {long_url} for details."
    shortened = _shorten_urls(text)

    # Extract the [url:XXXX] placeholder
    match = re.search(r'\[url:(\w+)\]', shortened)
    assert match, f"Expected [url:ID] placeholder in: {shortened}"
    url_id = match.group(1)

    # Verify it's in the cache
    assert url_id in _id_cache
    assert _id_cache[url_id] == long_url

    # Call read_item with url: prefix — should return the URL directly
    result = read_item(entry_id=f"url:{url_id}")
    assert "url" in result
    assert result["url"] == long_url


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s", "--tb=short"])
