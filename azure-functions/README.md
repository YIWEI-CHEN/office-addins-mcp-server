# Office Add-ins MCP Server for Azure Functions ⚠️

An Azure Functions-based Model Context Protocol (MCP) server for discovering and managing Microsoft Office Add-ins across Word, Excel, PowerPoint, Outlook, and Teams.

> **⚠️ EXPERIMENTAL**: This Azure Functions implementation is currently experimental and has not been fully tested. Use at your own risk and consider using the standalone server implementation for production workloads.

## Overview

This server exposes Office add-ins management tools via the Model Context Protocol (MCP) for integration with AI clients like GitHub Copilot and other language models. It's built using Azure Functions with MCP extension support.

## Features

- **get_addin_details**: Fetch details of a Microsoft Office add-in by its asset ID
- **health_check**: Health check endpoint to verify the MCP server is running
- **Azure Functions runtime**: Serverless deployment with automatic scaling
- **MCP protocol support**: Compatible with GitHub Copilot and other MCP clients

## Prerequisites

- Azure Functions Core Tools v4.0.7030 or later
- Python 3.10 or later
- Azure subscription (for deployment)

## Local Development

1. **Install Azure Functions Core Tools**:
   ```bash
   npm install -g azure-functions-core-tools@4 --unsafe-perm true
   ```

2. **Navigate to the Azure Functions directory**:
   ```bash
   cd azure-functions
   ```

3. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run locally**:
   ```bash
   func start
   ```

   The server will be available at:
   - **Streamable HTTP**: `http://localhost:7071/runtime/webhooks/mcp`
   - **Server-Sent Events**: `http://localhost:7071/runtime/webhooks/mcp/sse`

## Deployment to Azure

1. **Navigate to the Azure Functions directory**:
   ```bash
   cd azure-functions
   ```

2. **Create a Function App**:
   ```bash
   az functionapp create --resource-group myResourceGroup --consumption-plan-location westeurope --runtime python --runtime-version 3.11 --functions-version 4 --name myFunctionApp --storage-account mystorageaccount
   ```

3. **Deploy**:
   ```bash
   func azure functionapp publish myFunctionApp
   ```

3. **Get the MCP extension system key**:
   ```bash
   az functionapp keys list --resource-group myResourceGroup --name myFunctionApp --query systemKeys.mcp_extension --output tsv
   ```

## MCP Client Configuration

### For GitHub Copilot in VS Code

Create or update your `mcp.json` configuration:

```json
{
    "inputs": [
        {
            "type": "promptString",
            "id": "functions-mcp-extension-system-key",
            "description": "Azure Functions MCP Extension System Key",
            "password": true
        },
        {
            "type": "promptString",
            "id": "functionapp-host",
            "description": "The host domain of the function app (e.g., myapp.azurewebsites.net)"
        }
    ],
    "servers": {
        "local-office-addins-mcp": {
            "type": "http",
            "url": "http://localhost:7071/runtime/webhooks/mcp"
        },
        "azure-office-addins-mcp": {
            "type": "http",
            "url": "https://${input:functionapp-host}/runtime/webhooks/mcp",
            "headers": {
                "x-functions-key": "${input:functions-mcp-extension-system-key}"
            }
        }
    }
}
```

## Available Tools

### get_addin_details

Fetches details of a Microsoft Office add-in by its asset ID.

**Parameters:**
- `asset_id` (string, required): The unique identifier of the Office add-in (typically a GUID)

**Example usage:**
```
@azure-office-addins-mcp get_addin_details asset_id="12345678-1234-1234-1234-123456789abc"
```

### health_check

Returns the health status of the MCP server.

**Parameters:** None

**Example usage:**
```
@azure-office-addins-mcp health_check
```

## Configuration

The server is configured via `host.json`:

```json
{
  "extensions": {
    "mcp": {
      "instructions": "Microsoft Office Add-ins MCP Server - Use this server to discover and manage Office add-ins...",
      "serverName": "Office Add-ins MCP Server",
      "serverVersion": "1.0.0"
    }
  }
}
```

## Development vs Production

### Local Development
- Uses Azure Functions Core Tools
- No authentication required
- Available at `http://localhost:7071/runtime/webhooks/mcp`

### Azure Deployment
- Requires system key authentication (`mcp_extension`)
- Available at `https://yourapp.azurewebsites.net/runtime/webhooks/mcp`
- Automatic scaling and high availability

## Troubleshooting

1. **Import errors during local development**: Make sure all dependencies are installed via `pip install -r requirements.txt`

2. **Authentication errors in Azure**: Ensure you're providing the correct `mcp_extension` system key in the `x-functions-key` header

3. **Timeout errors**: The function timeout is set to 5 minutes in `host.json`. Adjust if needed for your use case.

## Migration from Standalone MCP Server

This project has been converted from a standalone FastMCP server to Azure Functions. Key changes:

- **Transport**: Now uses Azure Functions MCP extension instead of FastMCP transports
- **Deployment**: Serverless Azure Functions instead of self-hosted server
- **Configuration**: `host.json` and `local.settings.json` instead of `.env` files
- **Authentication**: Azure Functions system keys instead of custom auth

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test locally with `func start`
5. Submit a pull request

For issues or questions, please open an issue on GitHub.