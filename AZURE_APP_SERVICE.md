# Deploy Office Add-ins MCP Server to Azure App Service

This guide provides comprehensive instructions for deploying the Office Add-ins MCP Server to Azure App Service using Azure Developer CLI (azd). This deployment option provides a production-ready HTTP endpoint with automatic scaling, monitoring, and HTTPS support.

## Overview

Azure App Service deployment offers several advantages:
- **Production-Ready**: Automatic HTTPS with Azure-managed certificates
- **Scalability**: Built-in auto-scaling capabilities
- **Monitoring**: Integrated logging and application insights
- **Reliability**: High availability with SLA guarantees
- **Cost-Effective**: Pay only for what you use with multiple pricing tiers

## Prerequisites

### 1. Azure Developer CLI (azd)

Install the Azure Developer CLI on your system:

**Linux/macOS:**
```bash
curl -fsSL https://aka.ms/install-azd.sh | bash
```

**Windows (PowerShell):**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://aka.ms/install-azd.ps1 | iex"
```

**Verification:**
```bash
azd version
```

### 2. Azure Subscription

- Ensure you have an active Azure subscription
- You'll need appropriate permissions to create resources in your subscription
- If you don't have a subscription, you can create a [free Azure account](https://azure.microsoft.com/free/)

### 3. Git Repository

- Clone or fork this repository to your local machine
- Ensure you're working from the `main` branch

## Deployment Process

### Step 1: Prepare Your Environment

1. **Clone the repository**:
   ```bash
   git clone https://github.com/YIWEI-CHEN/office-addins-mcp-server.git
   cd office-addins-mcp-server
   ```

2. **Ensure you're on the main branch**:
   ```bash
   git checkout main
   ```

### Step 2: Authenticate with Azure

1. **Login to Azure**:
   ```bash
   azd auth login
   ```
   
   This will open a browser window for Azure authentication. Sign in with your Azure account credentials.

2. **Verify authentication**:
   ```bash
   azd auth login --check-status
   ```

### Step 3: Initialize the Project

**First-time deployment only:**

```bash
azd init
```

This command will:
- Scan your project structure
- Detect the Python application
- Set up deployment configuration

### Step 4: Deploy to Azure

```bash
azd up
```

During deployment, you'll be prompted for:

1. **Azure Subscription**: Select from your available subscriptions
2. **Azure Region**: Choose a region close to your users (e.g., `East US`, `West US 2`, `West Europe`)
3. **Environment Name**: Enter a name for your deployment environment (e.g., `office-addins-prod`)

**Example interaction:**
```
? Select an Azure Subscription to use: 
  1. My Production Subscription (xxx-xxx-xxx)
> 2. My Development Subscription (yyy-yyy-yyy)

? Select an Azure location to use: 
> 1. (US) East US (eastus)
  2. (US) West US 2 (westus2)
  3. (Europe) West Europe (westeurope)

? Enter a new environment name: office-addins-prod
```

## Deployment Output

After successful deployment (typically 5-10 minutes), you'll receive:

### Infrastructure Created
- **Resource Group**: `rg-{environment-name}`
- **App Service Plan**: `plan-{environment-name}` (Linux, Basic B1)
- **App Service**: `app-api-{unique-id}`
- **Application Insights**: For monitoring and logging

### Endpoints
- **App Service URL**: `https://app-api-{unique-id}.azurewebsites.net`
- **MCP Endpoint**: `https://app-api-{unique-id}.azurewebsites.net/addins/mcp`

### Example Output
```
SUCCESS: Your application was provisioned and deployed to Azure in 8 minutes 32 seconds.
You can view the resources created under the resource group rg-office-addins-prod in the Azure portal:
https://portal.azure.com/#@/resource/subscriptions/{subscription-id}/resourceGroups/rg-office-addins-prod

The application is available at: https://app-api-gmqmpcvoduxtc.azurewebsites.net
```

## Testing Your Deployment

### 1. Basic Health Check

Test if the service is running:

```bash
curl https://your-app-url.azurewebsites.net/health
```

### 2. MCP Endpoint Test

Test the MCP server endpoint:

```bash
curl -H "Accept: text/event-stream" \
     https://your-app-url.azurewebsites.net/addins/mcp
```

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "method": "notifications/initialized",
  "params": {
    "serverInfo": {
      "name": "office-addins-mcp-server",
      "version": "0.1.0"
    }
  }
}
```

### 3. Integration with MCP Clients

Configure your MCP client (like Claude Desktop) to use your deployed endpoint:

```json
{
  "servers": {
    "office-addins": {
      "command": "curl",
      "args": ["-H", "Accept: text/event-stream", "https://your-app-url.azurewebsites.net/addins/mcp"]
    }
  }
}
```

## Configuration and Management

### Environment Variables

The deployment automatically configures these environment variables:

| Variable | Value | Description |
|----------|-------|-------------|
| `HOST` | `0.0.0.0` | Listen on all interfaces |
| `PORT` | `8000` | Application port |
| `TRANSPORT` | `sse` | Server-Sent Events transport |
| `DEBUG` | `false` | Production mode |

### Scaling Configuration

**Current Setup:**
- **App Service Plan**: Basic B1 (1 vCPU, 1.75 GB RAM)
- **Auto-scaling**: Manual scaling (can be upgraded to automatic)

**To upgrade for production workloads:**

1. **Scale up** (increase instance size):
   ```bash
   az appservice plan update --name plan-{environment-name} \
     --resource-group rg-{environment-name} \
     --sku S1  # Standard tier with auto-scaling
   ```

2. **Scale out** (increase instance count):
   ```bash
   az webapp scale --name app-api-{unique-id} \
     --resource-group rg-{environment-name} \
     --instance-count 3
   ```

### Monitoring and Logging

#### View Live Logs

```bash
# Using Azure CLI
az webapp log tail --name app-api-{unique-id} \
  --resource-group rg-{environment-name}

# Using azd
azd monitor --live
```

#### Application Insights

Access detailed monitoring at:
```
https://portal.azure.com → Your Resource Group → Application Insights
```

Available metrics:
- Request rates and response times
- Error rates and exceptions  
- CPU and memory usage
- Custom telemetry from the MCP server

## Management Commands

### Redeploy Application Only

After making code changes:

```bash
azd deploy
```

This redeploys only the application code without recreating infrastructure.

### View Deployment Status

```bash
azd show
```

### Environment Management

```bash
# List environments
azd env list

# Select different environment
azd env select {environment-name}

# View environment configuration
azd env get-values
```

### Resource Cleanup

**⚠️ Warning**: This will delete all Azure resources created for this deployment.

```bash
azd down
```

You'll be prompted to confirm deletion:
```
? Total resources to delete: 4, are you sure you want to continue? (Y/n)
```

## Troubleshooting

### Common Issues

#### 1. Authentication Errors
```bash
# Clear and re-authenticate
azd auth logout
azd auth login
```

#### 2. Deployment Failures
```bash
# View detailed logs
azd deploy --debug

# Check infrastructure status
azd show --output json
```

#### 3. Application Not Starting
```bash
# Check application logs
az webapp log tail --name app-api-{unique-id} \
  --resource-group rg-{environment-name}
```

#### 4. MCP Endpoint Issues

- Verify the endpoint URL format: `https://{app-url}/addins/mcp`
- Check that the `Accept: text/event-stream` header is included
- Ensure your MCP client supports SSE transport

### Performance Optimization

#### 1. Cold Start Optimization
- Enable "Always On" setting to prevent cold starts
- Consider upgrading to Standard or Premium tier

#### 2. Connection Limits
- Basic tier supports up to 350 concurrent connections
- Upgrade to Standard/Premium for higher limits

#### 3. Geographic Distribution
- Deploy to multiple regions for global users
- Use Azure Front Door for global load balancing

## Cost Optimization

### Pricing Tiers

| Tier | Monthly Cost (USD) | Features |
|------|-------------------|----------|
| **Free** | $0 | 60 minutes/day, 1GB storage |
| **Basic B1** | ~$13 | Custom domains, SSL, manual scaling |
| **Standard S1** | ~$56 | Auto-scaling, staging slots, backups |
| **Premium P1v3** | ~$124 | Advanced scaling, VNet integration |

### Cost-Saving Tips

1. **Use Free Tier** for development/testing
2. **Scale down** during off-hours
3. **Monitor usage** with Azure Cost Management
4. **Set spending alerts** to avoid surprises

## Security Best Practices

### 1. Network Security
- **HTTPS Only**: Enabled by default
- **Custom Domains**: Add your own domain with SSL certificate
- **IP Restrictions**: Limit access to specific IP ranges if needed

### 2. Authentication
- Consider adding Azure AD authentication for production
- Use API keys or tokens for programmatic access

### 3. Monitoring
- Enable Application Insights security monitoring
- Set up alerts for unusual activity patterns

## Next Steps

1. **Custom Domain**: Configure your own domain name
2. **CI/CD Pipeline**: Set up automated deployments from GitHub
3. **Monitoring**: Configure custom alerts and dashboards
4. **Scaling**: Implement auto-scaling rules based on usage
5. **Security**: Add authentication and authorization layers

## Support and Resources

- **Azure App Service Documentation**: https://docs.microsoft.com/azure/app-service/
- **Azure Developer CLI**: https://docs.microsoft.com/azure/developer/azure-developer-cli/
- **MCP Protocol**: https://github.com/modelcontextprotocol/
- **Project Issues**: https://github.com/YIWEI-CHEN/office-addins-mcp-server/issues

---

**Need help?** Create an issue in the [GitHub repository](https://github.com/YIWEI-CHEN/office-addins-mcp-server/issues) with your deployment logs and error messages.