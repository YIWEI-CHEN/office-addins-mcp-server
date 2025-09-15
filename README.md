# Office Add‑ins MCP Server
[中文版本](.README_zh_tw.md)

## Overview

This repository provides a simple server implementation based on the
**Model Context Protocol (MCP)** using the official Python SDK. MCP
standardizes how large language models (LLMs) communicate with external data
sources and tools.  The `FastMCP` class encapsulates the complexity of MCP
and lets developers expose ordinary Python functions as MCP tools or
resources with minimal boilerplate.

This server exposes a single tool that retrieves the details of a Microsoft
Office add‑in by its **asset ID**.  An MCP‑enabled client (for example,
Claude Desktop) can call this tool to fetch metadata from the Office
Add‑ins API.

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

## Future Improvements

This project currently implements a single tool.  Potential enhancements
include:

* **Caching:** Cache frequent lookups to reduce calls to the Office API.
* **Improved error handling:** Provide richer error messages and recovery
  guidance.
* **Additional tools:** Expand the server with other Office API functions,
  such as searching add‑ins or listing a user’s installed add‑ins.

Contributions are welcome—feel free to submit issues or pull requests.
