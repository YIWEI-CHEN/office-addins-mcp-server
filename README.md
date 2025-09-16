# Office Add‑ins MCP Server
[中文版本](./README_zh_tw.md)

## News

**2024-09-13** - Launched the MCP server development

## Overview

A Model Context Protocol (MCP) server for discovering and managing Microsoft Office Add‑ins across Word, Excel, PowerPoint, Outlook, and Teams. This server enables AI agents to search add-ins and retrieve detailed metadata, install or uninstall add-ins, handle submission, validation, and publishing of custom add-ins.

This repository provides a comprehensive server implementation based on the
**Model Context Protocol (MCP)** using the official Python SDK. MCP
standardizes how large language models (LLMs) communicate with external data
sources and tools.  The `FastMCP` class encapsulates the complexity of MCP
and lets developers expose ordinary Python functions as MCP tools or
resources with minimal boilerplate.

Currently, the server provides basic add-in detail retrieval functionality, with comprehensive add-in management features planned for future releases (see Roadmap section).

## Features

* **Standardized interface:** MCP servers act like APIs for LLMs, allowing
  secure access to internal tools or data.
* **FastMCP convenience:** The SDK generates tool schemas from type hints
  and docstrings, minimizing boilerplate while supporting both synchronous
  and asynchronous functions.
* **SSE transport:** This example uses Server‑Sent Events for remote clients,
  but you can switch to STDIO for local testing.
* **Async HTTP:** The tool uses `httpx.AsyncClient` to call the Office
  Add‑ins API without blocking the event loop.

## Installation and Setup

This project recommends using [uv](https://docs.astral.sh/uv/) to manage
Python dependencies.  uv handles dependency resolution and virtual
environments via `pyproject.toml` and `uv.lock`.

1. **Clone the repository**:

   ```bash
   git clone <repository-url> office-addins-mcp-server
   cd office-addins-mcp-server
   ```

2. **Initialize the uv project**:

   ```bash
   # Generate pyproject.toml and uv.lock in the project root
   uv init
   ```

3. **Add dependencies**:

   ```bash
   # Add the MCP SDK and httpx to the project
   uv add "mcp[cli]"
   uv add httpx
   ```

4. **Install dependencies** (if a lock file already exists):

   ```bash
   uv install
   ```

After these steps, uv manages all dependencies.  You can run the server
with:

```bash
uv run python src/server.py
```

The dependencies used are:

* `mcp[cli]` – the official Model Context Protocol SDK and CLI tools.
* `httpx` – an asynchronous HTTP client for calling the Office API.

## Running the Server

Start the server by running:

```bash
python src/server.py
```

By default the server uses SSE transport and listens on `0.0.0.0:8000`.
To change the port or use STDIO transport, edit the call to `mcp.run()` in
`src/server.py`.

## Testing the Server

To verify that the server works, connect with an MCP‑compatible client and
invoke the `get_addin_details` tool.  The official SDK provides CLI tools
via `uv run mcp`.  For example:

```bash
# Start the server in development mode
uv run mcp dev src/server.py

# Or install it into Claude Desktop
uv run mcp install src/server.py
```

Once running, you can connect using Claude Desktop or the MCP Inspector to
test the tool.  Refer to the official documentation for writing custom
clients.

## Project Structure

```
office-addins-mcp-server/
├── docs/
│   └── execution_plan.md  # Implementation plan in English
├── src/
│   ├── __init__.py       # Marks src as a package
│   └── server.py         # MCP server implementation
├── requirements.txt       # Python dependencies
├── README.md              # English project description
└── README_zh_tw.md        # Traditional Chinese description
```

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

## Future Improvements

This project currently implements a single tool.  Potential enhancements
include:

* **Caching:** Cache frequent lookups to reduce calls to the Office API.
* **Improved error handling:** Provide richer error messages and recovery
  guidance.
* **Additional tools:** Expand the server with other Office API functions,
  such as searching add‑ins or listing a user's installed add‑ins.

Contributions are welcome—feel free to submit issues or pull requests.
