# Examples and Demonstrations

This folder contains example scripts and test utilities for the Office Add-ins MCP Server.

## Files

### `demo_search.py`
A comprehensive demonstration script that showcases all the search capabilities of the `search_addins` function. This script includes examples for:
- Basic search functionality
- Advanced filtering (by client, category, free status)
- Pagination and sorting
- Specific add-in lookup
- Detailed add-in information retrieval

Run with: `uv run python examples/demo_search.py`

### `test_server.py` 
A test script that verifies the MCP server configuration and tool registration. This script:
- Tests that both tools (`get_addin_details` and `search_addins`) are properly registered
- Validates tool functionality with real API calls
- Ensures the server is ready for production use

Run with: `uv run python examples/test_server.py`

## Usage

These examples are designed to help developers understand how to use the Office Add-ins MCP Server and can serve as templates for building custom integrations.

```bash
# Run the demonstration
uv run python examples/demo_search.py

# Test the server configuration
uv run python examples/test_server.py

# Run unit tests
uv run pytest tests/
```