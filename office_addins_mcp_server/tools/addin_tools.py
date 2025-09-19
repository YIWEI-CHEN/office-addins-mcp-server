"""
Office Add‑ins MCP Tools
========================

This module contains the MCP tools for managing Microsoft Office Add‑ins.
Each tool is implemented as an async function that can be registered with
the FastMCP server.
"""

from __future__ import annotations

import httpx
from typing import Optional, List, Union


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


async def search_addins(
    query: Optional[str] = None,
    category: Optional[List[str]] = None,
    free: Optional[bool] = None,
    clients: Optional[List[str]] = None,
    productgroup: Optional[List[str]] = None,
    productids: Optional[List[str]] = None,
    assetids: Optional[List[str]] = None,
    providertype: Optional[str] = None,
    orderfield: Optional[str] = None,
    orderby: Optional[str] = None,
    top: Optional[int] = None,
    skiptoitem: Optional[int] = None,
    date: Optional[str] = None,
    getMetaOSApps: Optional[bool] = None,
) -> dict:
    """Search for Office Add-ins using the Office Store API.

    This tool provides comprehensive search functionality for Microsoft Office Add-ins
    using the Office Store Search API. It supports filtering, sorting, pagination,
    and various search criteria.

    Parameters
    ----------
    query : str, optional
        Search query/keywords to find specific add-ins (parameter: qu).
    category : List[str], optional
        Filter by categories (comma-separated list).
        Examples: ["Productivity", "Communication", "Education"]
    free : bool, optional
        Filter for free add-ins. Note: May not strictly filter only free add-ins.
    clients : List[str], optional
        Filter by Office client platform-product combinations.
        Format: Platform_Product (e.g., "Win32_Excel", "Mac_Word", "WAC_PowerPoint")
        Platforms: Win32, Mac, WAC (Web), Any (for Power BI)
        Products: Word, PowerPoint, Excel, OneNote, Outlook, PowerBI
    productgroup : List[str], optional
        Filter by product groups (e.g., ["Office"]).
    productids : List[str], optional
        Filter by specific product IDs/GUIDs.
    assetids : List[str], optional
        Filter by specific asset IDs (e.g., ["WA104381441", "WA102957665"]).
    providertype : str, optional
        Filter by provider type (e.g., "gmail").
    orderfield : str, optional
        Field to sort by. Options: "None" (0), "Title" (1), "Date" (2), "Price" (3), "Rating" (4)
    orderby : str, optional
        Sort direction. Options: "Desc", "Asc"
    top : int, optional
        Number of results to return (limit). Default varies by API.
    skiptoitem : int, optional
        Number of results to skip (offset for pagination).
    date : str, optional
        Date override for cache bypass (yyyy-MM-dd format).
    getMetaOSApps : bool, optional
        Include MetaOS applications in results.

    Returns
    -------
    dict
        A dictionary containing search results with the following structure:
        - TotalCount: Total number of matching results
        - Values: Array of add-in objects for the current page

    Raises
    ------
    httpx.HTTPStatusError
        If the API response status is not 200 OK.
    httpx.RequestError
        If there is a network failure while communicating with the API.

    Examples
    --------
    Search for Zoom add-ins:
        await search_addins(query="Zoom")

    Search for free Excel add-ins sorted by rating:
        await search_addins(
            query="productivity", 
            clients=["Win32_Excel"], 
            free=True, 
            orderfield="Rating", 
            orderby="Desc", 
            top=10
        )

    Get specific add-ins by asset ID:
        await search_addins(assetids=["WA104381441", "WA102957665"])
    """
    # Base URL for the search API
    base_url = "https://api.addins.omex.office.net/api/addins/search"
    
    # Build query parameters
    params = {}
    
    if query is not None:
        params["qu"] = query
    
    if category is not None:
        params["category"] = ",".join(category)
    
    if free is not None:
        params["free"] = str(free).lower()
    
    if clients is not None:
        params["clients"] = ",".join(clients)
    
    if productgroup is not None:
        params["productgroup"] = ",".join(productgroup)
    
    if productids is not None:
        params["productids"] = ",".join(productids)
    
    if assetids is not None:
        params["assetids"] = ",".join(assetids)
    
    if providertype is not None:
        params["providertype"] = providertype
    
    if orderfield is not None:
        params["orderfield"] = orderfield
    
    if orderby is not None:
        params["orderby"] = orderby
    
    if top is not None:
        params["top"] = str(top)
    
    if skiptoitem is not None:
        params["skiptoitem"] = str(skiptoitem)
    
    if date is not None:
        params["date"] = date
    
    if getMetaOSApps is not None:
        params["getMetaOSApps"] = str(getMetaOSApps).lower()

    # Use an asynchronous HTTP client to make the request
    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.get(base_url, params=params)
        # Raise an exception if the response status indicates an error
        response.raise_for_status()
        return response.json()