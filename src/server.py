"""
A Model Context Protocol (MCP) server for discovering and managing Microsoft
Office Add‑ins across Word, Excel, PowerPoint, Outlook, and Teams.

Features:
- Multiple transport support: STDIO (default), SSE, and HTTP
- Environment-based configuration via .env files
- Async HTTP client for Office Add-ins API calls
"""

from __future__ import annotations

import os
from pathlib import Path
import sys

from dotenv import load_dotenv

# Import FastMCP from the official MCP SDK.  FastMCP is located in the
# mcp.server.fastmcp module.
from mcp.server.fastmcp import FastMCP

from .tools import get_addin_details


def load_server_config() -> dict:
    """Load server configuration from .env file.

    Uses python-dotenv to load server configuration from a .env file in the project root.
    Falls back to defaults if variables are not set or the file doesn't exist.

    Returns
    -------
    dict
        Server configuration with keys: transport, host, port, path, sse_path
    """
    # Look for .env file in the project root and load it
    env_file = Path(__file__).parent.parent / ".env"
    load_dotenv(dotenv_path=env_file)

    # Get transport from environment variable, default to "stdio"
    transport = os.getenv("TRANSPORT", "stdio").lower()

    # Validate transport type
    if transport not in ["sse", "stdio", "http"]:
        print(f"Warning: Invalid transport '{transport}' in .env file. Using 'stdio' instead.")
        transport = "stdio"

    # Get host, port, and path configuration
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    path = os.getenv("PATH_PREFIX", "/")
    sse_path = os.getenv("SSE_PATH", "/sse")

    return {
        "transport": transport,
        "host": host,
        "port": port,
        "path": path,
        "sse_path": sse_path
    }


# Create the MCP server instance
mcp = FastMCP("Office Add‑ins MCP Server")


def register_tools() -> None:
    """Register all tools with the MCP server.

    Parameters
    ----------
    mcp : FastMCP
        The FastMCP server instance to register tools with.
    """
    @mcp.tool(
        name="get_addin_details",
        description="Fetch details of a Microsoft Office add‑in by its asset ID.",
    )
    async def get_addin_details_tool(asset_id: str) -> dict:
        """MCP tool wrapper for get_addin_details."""
        return await get_addin_details(asset_id)


def run_server() -> None:
    """Run the MCP server with configuration from .env file.

    Loads server configuration from .env file and starts the FastMCP server.
    Configuration includes transport type, host, port, and path prefix.
    """
    try:
        # Load server configuration from .env file
        config = load_server_config()

        # Register all tools with the server
        register_tools()

        # Start the server with the specified transport and configuration
        if config["transport"] == "stdio":
            # STDIO transport for local CLI clients
            mcp.run(transport="stdio")
        elif config["transport"] == "sse":
            # SSE transport for web service deployment
            mcp.run(transport="sse", host=config["host"], port=config["port"], path=config["sse_path"])
        elif config["transport"] == "http":
            # HTTP transport for streamable HTTP requests
            mcp.run(transport="streamable-http", host=config["host"], port=config["port"], path=config["path"])

    except KeyboardInterrupt:
        print("\nServer shutdown requested by user")
    except Exception as e:
        print(f"Error starting MCP server: {e}")
        sys.exit(1)


def main() -> None:
    """Entry point for running the MCP server.

    When executed as a script, this function loads transport configuration from
    .env file, registers tools and starts the FastMCP server.
    """
    run_server()


if __name__ == "__main__":
    main()