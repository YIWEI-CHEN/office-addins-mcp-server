"""
A Model Context Protocol (MCP) server for discovering and managing Microsoft
Office Add‑ins across Word, Excel, PowerPoint, Outlook, and Teams.

Features:
- Multiple transport support: STDIO (default), SSE, and HTTP
- Async HTTP client for Office Add-ins API calls
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime

import click

# Import FastMCP from the official MCP SDK.  FastMCP is located in the
# mcp.server.fastmcp module.
from mcp.server.fastmcp import FastMCP

from office_addins_mcp_server.tools import get_addin_details, search_addins


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("office-addins-mcp")


def register_tools(mcp: FastMCP) -> None:
    """Register all tools with the MCP server.

    Parameters
    ----------
    mcp : FastMCP
        The FastMCP server instance to register tools with.
    """
    logger.info("Registering MCP tools...")
    
    @mcp.tool(
        name="get_addin_details",
        description="Fetch details of a Microsoft Office add‑in by its asset ID.",
    )
    async def get_addin_details_tool(asset_id: str) -> dict:
        """MCP tool wrapper for get_addin_details."""
        logger.debug(f"Fetching add-in details for asset ID: {asset_id}")
        return await get_addin_details(asset_id)
    
    @mcp.tool(
        name="search_addins",
        description="Search for Office Add-ins using comprehensive filtering, sorting, and pagination options.",
    )
    async def search_addins_tool(
        query: str | None = None,
        category: list[str] | None = None,
        free: bool | None = None,
        clients: list[str] | None = None,
        productgroup: list[str] | None = None,
        productids: list[str] | None = None,
        assetids: list[str] | None = None,
        providertype: str | None = None,
        orderfield: str | None = None,
        orderby: str | None = None,
        top: int | None = None,
        skiptoitem: int | None = None,
        date: str | None = None,
        getMetaOSApps: bool | None = None,
    ) -> dict:
        """MCP tool wrapper for search_addins."""
        logger.debug(f"Searching add-ins with query: {query}, filters: {locals()}")
        return await search_addins(
            query=query,
            category=category,
            free=free,
            clients=clients,
            productgroup=productgroup,
            productids=productids,
            assetids=assetids,
            providertype=providertype,
            orderfield=orderfield,
            orderby=orderby,
            top=top,
            skiptoitem=skiptoitem,
            date=date,
            getMetaOSApps=getMetaOSApps,
        )
    
    logger.info("Successfully registered 2 tools: get_addin_details, search_addins")


def create_mcp_server() -> FastMCP:
    """Create and configure the MCP server instance.

    Returns
    -------
    FastMCP
        Configured MCP server instance with tools registered
    """
    logger.info("Creating MCP server instance...")
    mcp = FastMCP("Office Add‑ins MCP Server")
    
    # Register all tools with the server
    register_tools(mcp)
    
    return mcp


def run_server(transport: str | None = None) -> FastMCP:
    """Run the MCP server with the specified transport.

    Creates and starts the FastMCP server with the given transport type.
    Defaults to stdio transport if no transport is specified.
    
    Parameters
    ----------
    transport : str | None
        Transport type (stdio, sse, or http). Defaults to stdio.
    
    Returns
    -------
    FastMCP
        The configured MCP server instance
    """
    try:
        logger.info("=" * 60)
        logger.info("🚀 Starting Office Add-ins MCP Server")
        logger.info("=" * 60)
        logger.info(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Create the MCP server instance
        mcp = create_mcp_server()

        # Start the server with the specified transport (default to stdio)
        transport = transport or "stdio"
        
        if transport == "stdio":
            logger.info("🔌 Starting server with STDIO transport for local CLI clients")
            mcp.run(transport="stdio")
        elif transport == "sse":
            logger.info("🌐 Starting server with SSE transport")
            mcp.run(transport="sse")
        elif transport == "http":
            logger.info("🌐 Starting server with HTTP transport")
            mcp.run(transport="streamable-http")

        return mcp

    except KeyboardInterrupt:
        logger.info("\n⏹️  Server shutdown requested by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"❌ Error starting MCP server: {e}")
        sys.exit(1)


@click.command()
@click.option(
    "--transport", "-t",
    type=click.Choice(["stdio", "sse", "http"], case_sensitive=False),
    help="Transport protocol to use (overrides .env file). stdio: for local CLI clients, sse: for web clients, http: for REST API clients"
)
def main(transport: str | None = None) -> None:
    """Office Add-ins MCP Server
    
    A Model Context Protocol (MCP) server for discovering and managing Microsoft
    Office Add‑ins across Word, Excel, PowerPoint, Outlook, and Teams.
    
    Examples:
    
    \b
    uv run office-addins-mcp-server --transport stdio
    uv run office-addins-mcp-server --transport sse  
    uv run office-addins-mcp-server --transport http
    """
    logger.info("Initializing Office Add-ins MCP Server...")
    
    # Run server with transport argument from click
    run_server(transport)


if __name__ == "__main__":
    main()