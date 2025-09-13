# Office Add‑ins MCP Server

This repository contains a **Model Context Protocol (MCP) server** implemented in C# using the official `ModelContextProtocol` SDK. The server exposes two tools: **Add‑in search** and **Add‑in details**. These tools call existing Office Add‑ins APIs (provided by you) to search for available add‑ins and retrieve detailed information about a specific add‑in. The server is designed for GitHub Copilot and other MCP‑enabled clients, allowing agents to discover and use Office add‑ins within an AI workflow.

## Features

- 🚀 **MCP server via .NET** – Built on top of the official MCP C# SDK for easy integration and extensibility:contentReference[oaicite:0]{index=0}.  
- 🔎 **Search Office add‑ins** – Provide a keyword to find matching add‑ins via the configured Add‑in Search API.  
- 📄 **Get add‑in details** – Retrieve detailed metadata by add‑in ID via the configured Add‑in Details API.  
- 🌐 **Configurable endpoints** – API URIs and query patterns are read from configuration or environment variables, so you can point the tools at your own services.  
- 🧩 **Extensible design** – Add more tools simply by creating new methods annotated with `McpServerTool`:contentReference[oaicite:1]{index=1}.


## Quick start

1. Install [.NET 8.0 SDK](https://dotnet.microsoft.com/en-us/download) or newer.
2. Clone the repository and restore dependencies:

   ```bash
   git clone https://github.com/your-org/office-addins-mcp-server.git
   cd office-addins-mcp-server/src/OfficeAddinsMcpServer
   dotnet restore
   ```
3. Configure your API endpoints via environment variables or by editing Constants.cs:
- OFFICE_ADDIN_SEARCH_URI
- OFFICE_ADDIN_SEARCH_QUERY
- OFFICE_ADDIN_DETAILS_URI
- OFFICE_ADDIN_DETAILS_QUERY

4. Run the MCP server:
```bash
   dotnet run
```

# License
MIT