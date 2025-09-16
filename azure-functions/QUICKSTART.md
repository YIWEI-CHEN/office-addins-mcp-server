# Quick Start Guide ⚠️

> **⚠️ EXPERIMENTAL**: This Azure Functions implementation is experimental and not fully verified. Proceed with caution.

This directory contains the Azure Functions implementation of the Office Add-ins MCP Server.

## 🚀 Quick Local Development

```bash
# Navigate to this directory
cd azure-functions

# Start local development (installs dependencies and starts server)
./dev.sh
```

The server will be available at:
- **Streamable HTTP**: `http://localhost:7071/runtime/webhooks/mcp`
- **Server-Sent Events**: `http://localhost:7071/runtime/webhooks/mcp/sse`

## ☁️ Quick Azure Deployment

```bash
# Navigate to this directory
cd azure-functions

# Deploy to Azure (creates all resources and deploys)
./deploy.sh
```

The deployment script will:
1. Create a resource group
2. Create a storage account
3. Create an Azure Function App
4. Deploy your code
5. Display the connection details

## 📁 Directory Structure

```
azure-functions/
├── function_app.py          # Main Azure Functions app
├── host.json               # Azure Functions runtime configuration
├── local.settings.json     # Local development settings
├── requirements.txt        # Python dependencies for Azure Functions
├── pyproject.toml          # Azure Functions project configuration
├── .funcignore            # Deployment exclusions
├── deploy.sh              # Automated deployment script
├── dev.sh                 # Local development script
├── mcp.json.template      # MCP client configuration template
├── README.md              # Detailed documentation
└── .vscode/               # VS Code configuration
    └── settings.json      # Azure Functions settings
```

## 🔧 Manual Setup

If you prefer manual setup instead of using the scripts:

### Local Development
1. `pip install -r requirements.txt`
2. `func start`

### Azure Deployment
1. `az group create --name myResourceGroup --location eastus`
2. `az storage account create --name mystorageaccount --resource-group myResourceGroup`
3. `az functionapp create --resource-group myResourceGroup --name myFunctionApp --storage-account mystorageaccount --runtime python`
4. `func azure functionapp publish myFunctionApp`

## 📖 More Information

See [README.md](./README.md) for detailed documentation, configuration options, and troubleshooting.