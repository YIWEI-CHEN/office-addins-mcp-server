"""
Azure Functions MCP Server for Office Add-ins Discovery and Management

This Azure Functions app exposes Microsoft Office Add-ins management tools
via the Model Context Protocol (MCP) for integration with AI clients like
GitHub Copilot and other language models.

Features:
- MCP tool endpoint for getting Office add-in details
- Azure Functions runtime with MCP extension support
- Async HTTP client for Office Add-ins API calls
"""

import json
import logging
import azure.functions as func
import httpx

# Initialize the Azure Functions app
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ToolProperty:
    """Helper class to define MCP tool properties."""
    
    def __init__(self, property_name: str, property_type: str, description: str, required: bool = False):
        self.propertyName = property_name
        self.propertyType = property_type
        self.description = description
        self.required = required

    def to_dict(self):
        return {
            "propertyName": self.propertyName,
            "propertyType": self.propertyType,
            "description": self.description,
            "required": self.required
        }


# Define tool properties for get_addin_details
get_addin_details_properties = [
    ToolProperty(
        property_name="asset_id",
        property_type="string", 
        description="The unique identifier of the Office add-in to look up. This ID is typically a GUID assigned by the Office Store.",
        required=True
    )
]

# Convert tool properties to JSON
get_addin_details_properties_json = json.dumps([prop.to_dict() for prop in get_addin_details_properties])


@app.generic_trigger(
    arg_name="context",
    type="mcpToolTrigger",
    toolName="get_addin_details",
    description="Fetch details of a Microsoft Office add-in by its asset ID.",
    toolProperties=get_addin_details_properties_json,
)
async def get_addin_details(context) -> str:
    """
    Retrieve metadata for an Office add-in.

    This function issues an HTTP GET request to the Office Add-ins API
    endpoint and returns the JSON response containing add-in details.

    Args:
        context: The MCP trigger context containing the asset_id argument.

    Returns:
        str: JSON string containing the add-in details or error message.
    """
    try:
        # Parse the context to get arguments
        content = json.loads(context)
        arguments = content.get("arguments", {})
        asset_id = arguments.get("asset_id")
        
        if not asset_id:
            logger.error("No asset_id provided in the request")
            return json.dumps({"error": "asset_id is required"})
        
        logger.info(f"Fetching add-in details for asset ID: {asset_id}")
        
        # Construct the Office Add-ins API URL
        url = (
            "https://api.addins.omex.office.net/api/addins/details"
            f"?assetid={asset_id}"
        )
        
        # Make the HTTP request to the Office Add-ins API
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.get(url)
            response.raise_for_status()
            
            addin_details = response.json()
            logger.info(f"Successfully retrieved add-in details for asset ID: {asset_id}")
            
            return json.dumps(addin_details, indent=2)
            
    except httpx.HTTPStatusError as e:
        error_msg = f"HTTP error {e.response.status_code} when fetching add-in details: {e.response.text}"
        logger.error(error_msg)
        return json.dumps({"error": error_msg})
        
    except httpx.RequestError as e:
        error_msg = f"Network error when fetching add-in details: {str(e)}"
        logger.error(error_msg)
        return json.dumps({"error": error_msg})
        
    except json.JSONDecodeError as e:
        error_msg = f"Invalid JSON in context: {str(e)}"
        logger.error(error_msg)
        return json.dumps({"error": error_msg})
        
    except Exception as e:
        error_msg = f"Unexpected error: {str(e)}"
        logger.error(error_msg)
        return json.dumps({"error": error_msg})


@app.generic_trigger(
    arg_name="context",
    type="mcpToolTrigger",
    toolName="health_check",
    description="Health check endpoint to verify the MCP server is running.",
    toolProperties="[]",
)
def health_check(context) -> str:
    """
    Health check endpoint that returns server status.
    
    Args:
        context: The MCP trigger context (not used).
        
    Returns:
        str: JSON string with health status information.
    """
    from datetime import datetime
    
    health_info = {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "service": "Office Add-ins MCP Server (Azure Functions)",
        "version": "1.0.0"
    }
    
    logger.info("Health check requested")
    return json.dumps(health_info, indent=2)