#!/usr/bin/env python3
"""
Demonstration of the Office Add-ins Search API functionality
===========================================================

This script demonstrates various search capabilities of the search_addins function
using examples from the Office-AddIns-Search-API-Guide.md.
"""

import asyncio
import json
from office_addins_mcp_server.tools.addin_tools import search_addins, get_addin_details


async def demo_basic_search():
    """Demonstrate basic search functionality."""
    print("🔍 Basic Search for 'Zoom' add-ins:")
    print("=" * 50)
    
    result = await search_addins(query="Zoom", top=3)
    print(f"Total Results: {result['TotalCount']}")
    print(f"Showing first {len(result['Values'])} results:\n")
    
    for i, addin in enumerate(result['Values'], 1):
        print(f"{i}. {addin['Title']}")
        print(f"   Rating: {addin.get('Rating', 'N/A')} ⭐")
        print(f"   Provider: {addin.get('ProviderName', 'Unknown')}")
        print(f"   Price: {addin.get('Pricing', {}).get('Price', 'Unknown')}")
        print()


async def demo_advanced_filtering():
    """Demonstrate advanced filtering capabilities."""
    print("🎯 Advanced Filtering - Free Excel add-ins sorted by rating:")
    print("=" * 60)
    
    result = await search_addins(
        query="productivity",
        clients=["Win32_Excel", "Mac_Excel", "WAC_Excel"],
        free=True,
        orderfield="Rating",
        orderby="Desc",
        top=5
    )
    
    print(f"Total Results: {result['TotalCount']}")
    print(f"Showing top {len(result['Values'])} results:\n")
    
    for i, addin in enumerate(result['Values'], 1):
        supported_clients = [client['Client'] for client in addin.get('SupportedClients', [])]
        excel_clients = [client for client in supported_clients if 'Excel' in client]
        
        print(f"{i}. {addin['Title']}")
        print(f"   Rating: {addin.get('Rating', 'N/A')} ⭐")
        print(f"   Excel Support: {', '.join(excel_clients) if excel_clients else 'Cross-platform'}")
        print(f"   Categories: {', '.join([cat['Title'] for cat in addin.get('Categories', [])])}")
        print()


async def demo_specific_lookup():
    """Demonstrate specific add-in lookup by asset ID."""
    print("📋 Specific Add-in Lookup:")
    print("=" * 35)
    
    # Search by asset IDs
    result = await search_addins(assetids=["WA104381441", "WA102957665"])
    
    print(f"Found {len(result['Values'])} specific add-ins:\n")
    
    for addin in result['Values']:
        print(f"• {addin['Title']} (ID: {addin['Id']})")
        print(f"  Description: {addin.get('ShortDescription', 'No description')[:100]}...")
        print(f"  Last Updated: {addin.get('LastUpdatedDate', 'Unknown')[:10]}")
        print()


async def demo_pagination():
    """Demonstrate pagination functionality."""
    print("📄 Pagination Demo - Office add-ins:")
    print("=" * 40)
    
    # First page
    page1 = await search_addins(query="office", top=3, skiptoitem=0)
    print(f"Page 1 (Total available: {page1['TotalCount']}):")
    for i, addin in enumerate(page1['Values'], 1):
        print(f"  {i}. {addin['Title']}")
    
    # Second page  
    page2 = await search_addins(query="office", top=3, skiptoitem=3)
    print(f"\nPage 2:")
    for i, addin in enumerate(page2['Values'], 4):
        print(f"  {i}. {addin['Title']}")


async def demo_sorting():
    """Demonstrate sorting capabilities."""
    print("📊 Sorting Demo - Alphabetical vs Rating:")
    print("=" * 45)
    
    # Sort by title
    by_title = await search_addins(orderfield="Title", orderby="Asc", top=3)
    print("Sorted by Title (A-Z):")
    for i, addin in enumerate(by_title['Values'], 1):
        print(f"  {i}. {addin['Title']}")
    
    # Sort by rating
    by_rating = await search_addins(orderfield="Rating", orderby="Desc", top=3)
    print(f"\nSorted by Rating (High to Low):")
    for i, addin in enumerate(by_rating['Values'], 1):
        print(f"  {i}. {addin['Title']} - {addin.get('Rating', 'N/A')} ⭐")


async def demo_addin_details():
    """Demonstrate getting detailed add-in information."""
    print("🔍 Detailed Add-in Information:")
    print("=" * 40)
    
    details = await get_addin_details("WA102957665")
    addin = details["Value"]
    
    print(f"Title: {addin['Title']}")
    print(f"Description: {addin.get('ShortDescription', 'No description')}")
    print(f"Provider: {addin.get('ProviderName', 'Unknown')}")
    print(f"Version: {addin.get('Version', 'Unknown')}")
    print(f"Categories: {', '.join([cat['Title'] for cat in addin.get('Categories', [])])}")
    print(f"Supported Clients: {len(addin.get('SupportedClients', []))} platforms")
    
    if addin.get('SupportedClients'):
        print("\nPlatform Support:")
        for client in addin['SupportedClients'][:3]:  # Show first 3
            print(f"  • {client['DisplayName']}")


async def main():
    """Run all demonstrations."""
    print("🏢 Office Add-ins Search API Demonstration")
    print("=" * 50)
    print("This demo showcases the search_addins function capabilities")
    print("based on the Office-AddIns-Search-API-Guide.md examples.\n")
    
    demos = [
        demo_basic_search,
        demo_advanced_filtering,
        demo_specific_lookup,
        demo_pagination,
        demo_sorting,
        demo_addin_details,
    ]
    
    for demo in demos:
        try:
            await demo()
            print("\n" + "─" * 80 + "\n")
        except Exception as e:
            print(f"❌ Error in {demo.__name__}: {e}\n")
    
    print("✅ Demonstration completed!")


if __name__ == "__main__":
    asyncio.run(main())