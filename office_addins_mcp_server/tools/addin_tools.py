"""
Office Add‑ins MCP Tools
========================

This module contains the MCP tools for managing Microsoft Office Add‑ins.
Each tool is implemented as an async function that can be registered with
the FastMCP server.
"""

from __future__ import annotations

import httpx


async def get_addin_details(asset_id: str) -> dict:
    """Retrieve metadata for an Office add‑in.

    This asynchronous tool issues an HTTP GET request to the Office Add‑ins API
    endpoint and returns the JSON response as a Python dictionary.  Clients
    should provide the unique asset ID for the add‑in they want to inspect.

    Parameters
    ----------
    asset_id: str
        The unique identifier of the Office add‑in to look up.  This ID is
        typically a GUID assigned by the Office Store.

    Returns
    -------
    dict
        A dictionary containing the add‑in details returned by the API.

    Raises
    ------
    httpx.HTTPStatusError
        If the API response status is not 200 OK.
    httpx.RequestError
        If there is a network failure while communicating with the API.
    """
    # Construct the query URL.  The API expects the asset ID as a query
    # parameter named "assetid".  No authentication is required for this
    # endpoint at the time of writing.
    url = (
        "https://api.addins.omex.office.net/api/addins/details"
        f"?assetid={asset_id}"
    )

    # Use an asynchronous HTTP client to avoid blocking the event loop.  A
    # context manager ensures that network resources are cleaned up properly.
    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.get(url)
        # Raise an exception if the response status indicates an error.  FastMCP
        # automatically converts exceptions into MCP error responses for the
        # client.  See documentation for more details【410474369011793†L400-L447】.
        response.raise_for_status()
        return response.json()