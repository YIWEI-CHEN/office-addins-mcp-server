# Office Add‑ins MCP Server
[中文版本](./README_zh_tw.md)

## Overview

A Model Context Protocol (MCP) server for discovering and managing Microsoft Office Add‑ins across Word, Excel, PowerPoint, Outlook, and Teams. This server enables AI agents to search add-ins and retrieve detailed metadata, install or uninstall add-ins, handle submission, validation, and publishing of custom add-ins.

This repository provides a comprehensive server implementation based on the
**Model Context Protocol (MCP)** using the official Python SDK. MCP
standardizes how large language models (LLMs) communicate with external data
sources and tools.  The `FastMCP` class encapsulates the complexity of MCP
and lets developers expose ordinary Python functions as MCP tools or
resources with minimal boilerplate.

Currently, the server provides basic add-in detail retrieval functionality, with comprehensive add-in management features planned for future releases (see Roadmap section).

## Installation and Setup (Local Server)

This project uses [uv](https://docs.astral.sh/uv/) to manage Python dependencies and virtual environments and includes a `pyproject.toml` configuration file and a `uv.lock` file to ensure reproducible builds across different environments.

<!-- ### Prerequisites

First, install uv if you haven't already:

```bash
# Install uv (cross-platform)
curl -LsSf https://astral.sh/uv/install.sh | sh

# Or on macOS with Homebrew
brew install uv

# Or on Windows with PowerShell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
``` -->

<!-- ### Project Setup -->

1. **Clone the repository**:

   ```bash
   git clone <repository-url> office-addins-mcp-server
   cd office-addins-mcp-server
   ```

2. **Install dependencies using uv**:

   ```bash
   # This will create a virtual environment and install all dependencies
   # based on pyproject.toml and uv.lock
   uv sync
   ```
<!-- 
3. **Activate the virtual environment** (optional):

   ```bash
   # Activate the virtual environment manually
   source .venv/bin/activate  # Linux/macOS
   # or
   .venv\Scripts\activate     # Windows
   ```

   Alternatively, you can use `uv run` to execute commands in the virtual environment without activating it. -->

### Running the Server

You can run the server in several ways:

```bash
# Option 1: Using the installed script (recommended)
uv run office-addins-mcp-server

# Option 2: With specific transport
uv run office-addins-mcp-server --transport stdio
uv run office-addins-mcp-server --transport sse
uv run office-addins-mcp-server --transport http

# Option 3: Using uv run with direct path
uv run python office_addins_mcp_server/server.py --transport stdio

# Option 4: After activating the virtual environment
source .venv/bin/activate
python office_addins_mcp_server/server.py --transport stdio
```

**Transport Types:**
- **`stdio`** (default): Standard input/output transport, perfect for local testing and CLI integration
- **`sse`**: Server-Sent Events transport, ideal for web service deployment
- **`http`**: Streamable HTTP transport, suitable for HTTP-based integrations

## 🧪 Experimental Remote Server

> **⚠️ EXPERIMENTAL**: A remote instance of this MCP server is available for testing purposes only. This is not intended for production use and may have limited uptime, rate limits, or be discontinued without notice.

**Remote MCP Endpoint**: `https://app-api-gmqmpcvoduxtc.azurewebsites.net/addins/mcp`

**Usage**: You can test the MCP protocol with this endpoint, but please deploy your own instance for any serious work.

## Testing the Server

To verify that the server works, connect with an MCP‑compatible client and
invoke the `get_addin_details` tool.  The official SDK provides CLI tools
via `uv run mcp`.  For example:

```bash
# Start the server in development mode
uv run mcp dev office_addins_mcp_server/server.py

# Or install it into Claude Desktop
uv run mcp install office_addins_mcp_server/server.py

# Or use the installed script
uv run office-addins-mcp-server
```

Once running, you can connect using Claude Desktop or the MCP Inspector to
test the tool.  Refer to the official documentation for writing custom
clients.


## Azure App Service Deployment (Quick Start)
Deploy the MCP server as a web service on Azure App Service using Azure Developer CLI (azd). This provides a production-ready HTTP endpoint with automatic scaling and monitoring.

```bash
# Install Azure Developer CLI
curl -fsSL https://aka.ms/install-azd.sh | bash

# Authenticate with Azure
azd auth login

# Deploy to Azure (first time)
azd up

# Redeploy after changes
azd deploy
```

📖 **[Complete Azure App Service Deployment Guide →](./AZURE_APP_SERVICE.md)**

## Announcements

🎉 2025-09-13: The Office Add-ins MCP server is created.

## Roadmap

The following features are planned for future development:

1. **Add-in Search** - Implement comprehensive search functionality for Office add-ins
2. **Add Prompts to Search Add-in Features** - Enhance search with intelligent prompts and filtering
3. **OAuth2 Authentication** - Secure authentication for Microsoft Graph API access
4. **Show Installed Add-ins** - Display user's currently installed add-ins across Office applications
5. **Install/Uninstall Add-ins** - Programmatic installation and removal of add-ins
6. **Submit Custom Add-ins** - Enable submission of custom add-ins to the Office Store
7. **Validate Custom Add-ins** - Automated validation and compliance checking for custom add-ins
8. **Publish Custom Add-ins** - Streamlined publishing workflow for custom add-ins
9. **M365 Admin Push Add-ins** - Allow Microsoft 365 administrators to centrally deploy add-ins

Contributions are welcome—feel free to submit issues or pull requests.

<!-- ## Project Structure

```
office-addins-mcp-server/
├── office_addins_mcp_server/    # Standalone MCP server implementation
│   ├── server.py               # Main FastMCP server
│   └── tools/                  # MCP tools implementation
│       └── addin_tools.py      # Office add-ins API tools
├── azure-functions/            # Azure Functions serverless implementation
│   ├── function_app.py         # Azure Functions app
│   ├── host.json              # Azure Functions configuration
│   ├── pyproject.toml         # Azure Functions project config
│   ├── requirements.txt       # Azure Functions dependencies
│   ├── deploy.sh              # Automated deployment script
│   ├── dev.sh                 # Local development script
│   └── README.md              # Azure Functions documentation
├── requirements.txt            # Standalone server dependencies
├── pyproject.toml             # Standalone server project config
└── README.md                  # This file
``` -->

<!-- **Choose your deployment:**
- **Local/Standalone**: Use `uv run office-addins-mcp-server` in the root directory
- **Azure App Service**: Use `azd up` in the root directory -->