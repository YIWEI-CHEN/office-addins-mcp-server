"""
Office Add‑ins MCP Server
=========================

This module defines a simple Model Context Protocol (MCP) server using the
FastMCP framework.  The server exposes a single tool that fetches details of a
Microsoft Office add‑in identified by its asset ID.  MCP clients (for example,
LLM‑powered applications) can call this tool to retrieve up‑to‑date metadata
about a given add‑in.  By leveraging FastMCP's decorators and type hints, the
tool schema is automatically generated and validated for the client.

The server uses SSE (Server‑Sent Events) transport, making it suitable for
deployment as a web service.  When run directly, it listens on all network
interfaces on port 8000.
"""

from __future__ import annotations

# Import FastMCP from the official MCP SDK.  FastMCP is located in the
# mcp.server.fastmcp module.
from mcp.server.fastmcp import FastMCP

from .tools import get_addin_details


def create_server() -> FastMCP:
    """Instantiate and configure the MCP server.

    Returns
    -------
    fastmcp.FastMCP
        An instance of a configured FastMCP server ready to register tools.
    """
    return FastMCP(
        name="OfficeAddinsMCPServer",
        instructions=(
            "This MCP server provides tools to fetch Office add‑ins details from the "
            "Microsoft Office Add‑ins API. Use the available tools to retrieve "
            "information about a specific add‑in by its asset ID."
        ),
    )


# Create the MCP server instance
mcp = create_server()


@mcp.tool(
    name="get_addin_details",
    description="Fetch details of a Microsoft Office add‑in by its asset ID.",
    tags={"office", "addins", "api"},
)
async def get_addin_details_tool(asset_id: str) -> dict:
    """MCP tool wrapper for get_addin_details."""
    return await get_addin_details(asset_id)


def main() -> None:
    """Entry point for running the MCP server.

    When executed as a script, this function starts the FastMCP server using
    Server‑Sent Events (SSE) transport.  The server listens on all
    network interfaces on port 8000.  Adjust the host and port as needed
    for your deployment environment.
    """
    # Start the server with SSE transport.  This runs an ASGI application
    # under the hood using Uvicorn.  Pass `transport="stdio"` instead
    # for local tools (e.g. CLI clients)【410474369011793†L221-L227】.
    mcp.run(transport="sse", host="0.0.0.0", port=8000)


if __name__ == "__main__":
    main()