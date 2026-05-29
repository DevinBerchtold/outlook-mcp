# Outlook MCP Server

FastMCP server exposing Outlook email and calendar via COM automation (`win32com.client`).

## Files

- `outlook_mcp/server.py` — MCP server (tools: `list_folders`, `search_emails`, `search_calendar`, `read_item`)
- `test_server.py` — integration tests (14 tests, ~35s)

## Development

Install in editable mode with the test dependencies:

```
pip install -e ".[dev]"
```

## Architecture

Tools use a `_com_session()` context manager for COM threading. Search uses DASL filters via `Folder.GetTable()` for performance. Calendar uses `Items.Restrict` with `IncludeRecurrences`. A short-ID cache (`_id_cache`) maps 4-char base36 hashes to full EntryIDs. Long URLs in bodies are replaced with `[url:ID]` placeholders via `_shorten_urls`. Prompts provide reusable workflows (weekly summary, agenda, meeting prep, etc.).

## Testing

Integration tests against live Outlook — no mocking. Requires Outlook running on the machine.

```
pytest test_server.py -v -s
```

Tests call tool functions directly (FastMCP 3 keeps `@mcp.tool` functions callable as regular Python). Error-case tests use `pytest.raises`.
