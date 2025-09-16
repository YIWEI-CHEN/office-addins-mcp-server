#!/bin/bash
# dev.sh - Local development script for Azure Functions MCP Server

set -e

echo "🛠️  Starting local development environment for Office Add-ins MCP Server"
echo "======================================================================"

# Check if Functions Core Tools is installed
if ! command -v func &> /dev/null; then
    echo "❌ Azure Functions Core Tools is not installed."
    echo "   Run: npm install -g azure-functions-core-tools@4 --unsafe-perm true"
    exit 1
fi

# Check if we're in the right directory
if [ ! -f "function_app.py" ]; then
    echo "❌ function_app.py not found. Please run this script from the azure-functions directory."
    exit 1
fi

# Install dependencies
echo "📦 Installing Python dependencies..."
pip install -r requirements.txt

# Start the function app
echo "🚀 Starting Azure Functions host..."
echo "   Local endpoints:"
echo "   - Streamable HTTP: http://localhost:7071/runtime/webhooks/mcp"
echo "   - Server-Sent Events: http://localhost:7071/runtime/webhooks/mcp/sse"
echo ""
echo "Press Ctrl+C to stop the server"
echo "======================================================================"

func start