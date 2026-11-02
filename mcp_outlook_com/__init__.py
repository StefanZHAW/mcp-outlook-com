"""Outlook MCP Server für Claude Desktop (klassisches Outlook über COM)."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("mcp-outlook-com")
except PackageNotFoundError:
    # Direkt aus dem Quellordner gestartet, ohne Installation.
    __version__ = "0+unbekannt"
