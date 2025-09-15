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

import httpx
# Import FastMCP from the official MCP SDK.  FastMCP is located in the
# mcp.server.fastmcp module.
from mcp.server.fastmcp import FastMCP


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
async def get_addin_details(asset_id: str) -> dict:
    """Retrieve metadata for an Office add‑in.

    This asynchronous tool issues an HTTP GET request to the Office Add‑ins API
    endpoint and returns the JSON response as a Python dictionary.  Clients
    should provide the unique asset ID for the add‑in they want to inspect.

    Parameters
    ----------
    asset_id: str
        The unique identifier of the Office add‑in to look up.  This ID is
        typically a GUID assigned by the Office Store.

    Returns
    -------
    dict
        A dictionary containing the add‑in details returned by the API.

    Raises
    ------
    httpx.HTTPStatusError
        If the API response status is not 200 OK.
    httpx.RequestError
        If there is a network failure while communicating with the API.
    """
    # Construct the query URL.  The API expects the asset ID as a query
    # parameter named "assetid".  No authentication is required for this
    # endpoint at the time of writing.
    url = (
        "https://api.addins.omex.office.net/api/addins/details"
        f"?assetid={asset_id}"
    )

    # Use an asynchronous HTTP client to avoid blocking the event loop.  A
    # context manager ensures that network resources are cleaned up properly.
    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.get(url)
        # Raise an exception if the response status indicates an error.  FastMCP
        # automatically converts exceptions into MCP error responses for the
        # client.  See documentation for more details【410474369011793†L400-L447】.
        response.raise_for_status()
        return response.json()


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