"""
A Model Context Protocol (MCP) server for discovering and managing Microsoft
Office Add‑ins across Word, Excel, PowerPoint, Outlook, and Teams.

Features:
- Multiple transport support: STDIO (default), SSE, and HTTP
- Environment-based configuration via .env files
- Async HTTP client for Office Add-ins API calls
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
import sys
from datetime import datetime

from dotenv import load_dotenv

# Import FastMCP from the official MCP SDK.  FastMCP is located in the
# mcp.server.fastmcp module.
from mcp.server.fastmcp import FastMCP

from office_addins_mcp_server.tools import get_addin_details


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
    logger.info(f"Loading configuration from: {env_file}")
    if env_file.exists():
        load_dotenv(dotenv_path=env_file)
        logger.info("Configuration loaded from .env file")
    else:
        logger.info("No .env file found, using default configuration")

    # Get transport from environment variable, default to "stdio"
    transport = os.getenv("TRANSPORT", "stdio").lower()

    # Validate transport type
    if transport not in ["sse", "stdio", "http"]:
        logger.warning(f"Invalid transport '{transport}' in configuration. Using 'stdio' instead.")
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


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("office-addins-mcp")

# Create the MCP server instance
mcp = FastMCP("Office Add‑ins MCP Server")


def register_tools() -> None:
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
    
    logger.info("Successfully registered 1 tool: get_addin_details")


def run_server() -> None:
    """Run the MCP server with configuration from .env file.

    Loads server configuration from .env file and starts the FastMCP server.
    Configuration includes transport type, host, port, and path prefix.
    """
    try:
        logger.info("=" * 60)
        logger.info("🚀 Starting Office Add-ins MCP Server")
        logger.info("=" * 60)
        logger.info(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Load server configuration from .env file
        config = load_server_config()
        logger.info(f"Server configuration: {config}")

        # Register all tools with the server
        register_tools()

        # Start the server with the specified transport and configuration
        if config["transport"] == "stdio":
            logger.info("🔌 Starting server with STDIO transport for local CLI clients")
            mcp.run(transport="stdio")
        elif config["transport"] == "sse":
            logger.info(f"🌐 Starting server with SSE transport at {config['host']}:{config['port']}{config['sse_path']}")
            mcp.run(transport="sse")
        elif config["transport"] == "http":
            logger.info(f"🌐 Starting server with HTTP transport at {config['host']}:{config['port']}{config['path']}")
            mcp.run(transport="streamable-http")

    except KeyboardInterrupt:
        logger.info("\n⏹️  Server shutdown requested by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"❌ Error starting MCP server: {e}")
        sys.exit(1)


def main() -> None:
    """Entry point for running the MCP server.

    When executed as a script, this function loads transport configuration from
    .env file, registers tools and starts the FastMCP server.
    """
    logger.info("Initializing Office Add-ins MCP Server...")
    run_server()


if __name__ == "__main__":
    main()