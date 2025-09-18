"""
Azure App Service deployment for Office Add-ins MCP Server using Starlette.

This module creates a Starlette ASGI application that reuses the MCP server
from office_addins_mcp_server/server.py for Azure App Service deployment.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

import contextlib
from contextlib import asynccontextmanager

from dotenv import load_dotenv
from starlette.applications import Starlette
from starlette.middleware.cors import CORSMiddleware
from starlette.routing import Mount

# Import the MCP server creation function
from office_addins_mcp_server.server import create_mcp_server


# Lifespan context manager to start/stop the MCP session manager with the FastAPI app
@asynccontextmanager
async def mcp_lifespan(app):
    async with contextlib.AsyncExitStack() as stack:
        await stack.enter_async_context(mcp.session_manager.run())
        yield


def load_app_config() -> dict:
    """Load application configuration from .env file.

    Uses python-dotenv to load configuration from a .env file in the project root.
    Falls back to defaults if variables are not set or the file doesn't exist.

    Returns
    -------
    dict
        Application configuration with keys: host, port, debug
    """
    # Look for .env file in the project root and load it
    env_file = Path(__file__).parent / ".env"
    logger.info(f"Loading configuration from: {env_file}")
    if env_file.exists():
        load_dotenv(dotenv_path=env_file)
        logger.info("Configuration loaded from .env file")
    else:
        logger.info("No .env file found, using default configuration")

    return {
        "host": os.getenv("HOST", "0.0.0.0"),
        "port": int(os.getenv("PORT", "8000")),
        "debug": os.getenv("DEBUG", "false").lower() == "true"
    }


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("office-addins-mcp-server-app")


# Create the MCP server instance
mcp = create_mcp_server()


# Create the Starlette application
app = Starlette(
    # debug=config.get("debug", False),
    routes=[
        # Mount the MCP server at /addins path
        Mount("/addins", app=mcp.streamable_http_app()),
    ],
    lifespan=mcp_lifespan
)

# Wrap ASGI application with CORS middleware to expose Mcp-Session-Id header
# for browser-based clients (ensures 500 errors get proper CORS headers)
app = CORSMiddleware(
    app,
    allow_origins=["*"],  # Allow all origins - adjust as needed for production
    allow_methods=["GET", "POST", "DELETE"],  # MCP streamable HTTP methods
    expose_headers=["Mcp-Session-Id"],
)


if __name__ == "__main__":
    import uvicorn
    
    config = load_app_config()
    logger.info(f"🌐 Starting development server at {config['host']}:{config['port']}")
    
    uvicorn.run(
        app,
        host=config["host"],
        port=config["port"],
        reload=config["debug"]
    )