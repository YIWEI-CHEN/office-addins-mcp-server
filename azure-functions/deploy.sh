#!/bin/bash
# deploy.sh - Deployment script for Azure Functions MCP Server

set -e

# Configuration
RESOURCE_GROUP="${RESOURCE_GROUP:-office-addins-mcp-rg}"
LOCATION="${LOCATION:-eastus}"
FUNCTION_APP_NAME="${FUNCTION_APP_NAME:-office-addins-mcp-$(date +%s)}"
STORAGE_ACCOUNT_NAME="${STORAGE_ACCOUNT_NAME:-mcpstorage$(date +%s)}"

echo "🚀 Deploying Office Add-ins MCP Server to Azure Functions"
echo "=================================================="
echo "Resource Group: $RESOURCE_GROUP"
echo "Location: $LOCATION"
echo "Function App: $FUNCTION_APP_NAME"
echo "Storage Account: $STORAGE_ACCOUNT_NAME"
echo "=================================================="

# Check if Azure CLI is installed
if ! command -v az &> /dev/null; then
    echo "❌ Azure CLI is not installed. Please install it first."
    echo "   Visit: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli"
    exit 1
fi

# Check if Functions Core Tools is installed
if ! command -v func &> /dev/null; then
    echo "❌ Azure Functions Core Tools is not installed. Please install it first."
    echo "   Run: npm install -g azure-functions-core-tools@4 --unsafe-perm true"
    exit 1
fi

# Login check
echo "🔐 Checking Azure login status..."
if ! az account show &> /dev/null; then
    echo "❌ Not logged in to Azure. Please run 'az login' first."
    exit 1
fi

# Create resource group
echo "📦 Creating resource group..."
az group create --name "$RESOURCE_GROUP" --location "$LOCATION"

# Create storage account
echo "💾 Creating storage account..."
az storage account create \
    --name "$STORAGE_ACCOUNT_NAME" \
    --resource-group "$RESOURCE_GROUP" \
    --location "$LOCATION" \
    --sku Standard_LRS

# Create function app
echo "⚡ Creating Function App..."
az functionapp create \
    --resource-group "$RESOURCE_GROUP" \
    --consumption-plan-location "$LOCATION" \
    --runtime python \
    --runtime-version 3.11 \
    --functions-version 4 \
    --name "$FUNCTION_APP_NAME" \
    --storage-account "$STORAGE_ACCOUNT_NAME" \
    --os-type Linux

# Deploy the function
echo "🚀 Deploying function code..."
func azure functionapp publish "$FUNCTION_APP_NAME"

# Get the MCP extension system key
echo "🔑 Retrieving MCP extension system key..."
MCP_KEY=$(az functionapp keys list --resource-group "$RESOURCE_GROUP" --name "$FUNCTION_APP_NAME" --query systemKeys.mcp_extension --output tsv)

# Display results
echo ""
echo "✅ Deployment completed successfully!"
echo "=================================================="
echo "Function App URL: https://$FUNCTION_APP_NAME.azurewebsites.net"
echo "MCP Endpoint (Streamable HTTP): https://$FUNCTION_APP_NAME.azurewebsites.net/runtime/webhooks/mcp"
echo "MCP Endpoint (SSE): https://$FUNCTION_APP_NAME.azurewebsites.net/runtime/webhooks/mcp/sse"
echo "MCP Extension System Key: $MCP_KEY"
echo "=================================================="
echo ""
echo "🔧 MCP Client Configuration:"
echo "Add this to your mcp.json file:"
echo ""
echo '{'
echo '  "servers": {'
echo "    \"azure-office-addins-mcp\": {"
echo "      \"type\": \"http\","
echo "      \"url\": \"https://$FUNCTION_APP_NAME.azurewebsites.net/runtime/webhooks/mcp\","
echo "      \"headers\": {"
echo "        \"x-functions-key\": \"$MCP_KEY\""
echo "      }"
echo "    }"
echo "  }"
echo "}"