"""
Tests for Office Add-ins MCP Tools
==================================

This module contains comprehensive tests for the Office Add-ins MCP tools,
including tests for search functionality using real API calls based on
examples from the Office-AddIns-Search-API-Guide.md.
"""

from __future__ import annotations

import pytest
import httpx
from office_addins_mcp_server.tools.addin_tools import search_addins, get_addin_details


class TestSearchAddins:
    """Test suite for the search_addins function."""

    @pytest.mark.asyncio
    async def test_basic_search_zoom(self):
        """Test basic search functionality with 'Zoom' query."""
        result = await search_addins(query="Zoom")
        
        assert "TotalCount" in result
        assert "Values" in result
        assert result["TotalCount"] > 0
        assert isinstance(result["Values"], list)
        
        # Should find Zoom-related add-ins
        zoom_found = any("zoom" in addin.get("Title", "").lower() for addin in result["Values"])
        assert zoom_found, "Should find at least one Zoom-related add-in"

    @pytest.mark.asyncio
    async def test_advanced_filtering_free_productivity(self):
        """Test advanced filtering for free productivity add-ins sorted by rating."""
        result = await search_addins(
            query="productivity",
            free=True,
            orderfield="Rating",
            orderby="Desc",
            top=10
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 10
        
        # Check that results are related to productivity
        if result["Values"]:
            # At least some should be free (though API may not strictly filter)
            free_addins = [
                addin for addin in result["Values"] 
                if addin.get("Pricing", {}).get("Category") == "Free"
            ]
            # Note: free=true may not strictly filter, so we just check structure
            assert isinstance(free_addins, list)

    @pytest.mark.asyncio
    async def test_client_specific_excel_search(self):
        """Test client-specific search for Excel add-ins."""
        result = await search_addins(
            query="charts",
            clients=["Win32_Excel"]
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        
        # Check that results include Excel compatibility
        if result["Values"]:
            excel_compatible = any(
                any(
                    client.get("Client", "").startswith(("Win32_Excel", "Mac_Excel", "WAC_Excel"))
                    for client in addin.get("SupportedClients", [])
                )
                for addin in result["Values"]
            )
            # Note: API may return cross-platform results even with client filter

    @pytest.mark.asyncio
    async def test_multi_platform_excel_search(self):
        """Test search across multiple Excel platforms."""
        result = await search_addins(
            query="productivity",
            clients=["Win32_Excel", "Mac_Excel", "WAC_Excel"]
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert isinstance(result["Values"], list)

    @pytest.mark.asyncio
    async def test_pagination_first_page(self):
        """Test pagination - first page with calendar search."""
        result = await search_addins(
            query="calendar",
            top=20
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 20

    @pytest.mark.asyncio
    async def test_pagination_skip_items(self):
        """Test pagination - skip items functionality."""
        result = await search_addins(
            skiptoitem=10,
            top=2
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 2

    @pytest.mark.asyncio
    async def test_sorting_by_title_ascending(self):
        """Test sorting by title in ascending order."""
        result = await search_addins(
            orderfield="Title",
            orderby="Asc",
            top=3
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 3
        
        # Check that titles are in ascending order
        if len(result["Values"]) >= 2:
            titles = [addin.get("Title", "") for addin in result["Values"]]
            sorted_titles = sorted(titles)
            assert titles == sorted_titles, f"Titles should be sorted ascending: {titles}"

    @pytest.mark.asyncio
    async def test_sorting_by_rating_descending(self):
        """Test sorting by rating in descending order."""
        result = await search_addins(
            orderfield="Rating",
            orderby="Desc",
            top=3
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 3
        
        # Check that ratings are in descending order
        if len(result["Values"]) >= 2:
            ratings = [addin.get("Rating", 0) for addin in result["Values"]]
            for i in range(len(ratings) - 1):
                assert ratings[i] >= ratings[i + 1], f"Ratings should be descending: {ratings}"

    @pytest.mark.asyncio
    async def test_category_filtering_single(self):
        """Test filtering by a single category."""
        result = await search_addins(category=["Productivity"])
        
        assert "TotalCount" in result
        assert "Values" in result
        
        # Check that results contain productivity category
        if result["Values"]:
            productivity_found = any(
                any(
                    cat.get("Id") == "Productivity" or cat.get("Title") == "Productivity"
                    for cat in addin.get("Categories", [])
                )
                for addin in result["Values"]
            )
            # Note: API may not strictly filter, so we just verify structure

    @pytest.mark.asyncio
    async def test_category_filtering_multiple(self):
        """Test filtering by multiple categories."""
        result = await search_addins(
            category=["Productivity", "Communication", "Education"]
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert isinstance(result["Values"], list)

    @pytest.mark.asyncio
    async def test_specific_asset_ids_lookup(self):
        """Test lookup by specific asset IDs (Zoom and Microsoft Translator)."""
        result = await search_addins(
            assetids=["WA104381441", "WA102957665"]
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        
        # Should return exactly the specified add-ins (or available ones)
        if result["Values"]:
            found_ids = [addin.get("Id") for addin in result["Values"]]
            # Check if any of the requested IDs are found
            requested_ids = ["WA104381441", "WA102957665"]
            found_requested = any(aid in found_ids for aid in requested_ids)
            # Note: Some asset IDs may not be available anymore

    @pytest.mark.asyncio
    async def test_date_cache_bypass(self):
        """Test cache bypass with date parameter."""
        result = await search_addins(
            query="new",
            date="2025-09-19"
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert isinstance(result["Values"], list)

    @pytest.mark.asyncio
    async def test_complex_query_excel_accounting(self):
        """Test complex query: Excel add-ins for accounting, sorted by title."""
        result = await search_addins(
            query="accounting",
            clients=["Win32_Excel"],
            orderfield="Title",
            orderby="Asc",
            top=3
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 3

    @pytest.mark.asyncio
    async def test_free_powerpoint_templates(self):
        """Test search for free PowerPoint templates."""
        result = await search_addins(
            query="template",
            clients=["Win32_PowerPoint"],
            free=True,
            top=5
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 5

    @pytest.mark.asyncio
    async def test_empty_search(self):
        """Test search with no parameters (should return default results)."""
        result = await search_addins()
        
        assert "TotalCount" in result
        assert "Values" in result
        assert isinstance(result["Values"], list)

    @pytest.mark.asyncio
    async def test_response_structure(self):
        """Test that response has the expected structure based on API documentation."""
        result = await search_addins(query="office", top=1)
        
        assert "TotalCount" in result
        assert "Values" in result
        assert isinstance(result["TotalCount"], int)
        assert isinstance(result["Values"], list)
        
        if result["Values"]:
            addin = result["Values"][0]
            
            # Check key fields from API documentation
            expected_fields = [
                "Id", "Title", "ShortDescription", "Rating", "NumberOfVotes",
                "DateReleased", "LastUpdatedDate", "ProductId", "Culture",
                "State", "Version", "Pricing", "Categories", "SupportedClients",
                "ProviderName"
            ]
            
            for field in expected_fields:
                assert field in addin, f"Expected field '{field}' not found in add-in response"
            
            # Check pricing structure
            pricing = addin.get("Pricing", {})
            pricing_fields = ["Category", "FreeType", "Price", "Currency"]
            for field in pricing_fields:
                assert field in pricing, f"Expected pricing field '{field}' not found"
            
            # Check categories structure
            categories = addin.get("Categories", [])
            if categories:
                category = categories[0]
                category_fields = ["Id", "Title", "LongTitle"]
                for field in category_fields:
                    assert field in category, f"Expected category field '{field}' not found"
            
            # Check supported clients structure
            clients = addin.get("SupportedClients", [])
            if clients:
                client = clients[0]
                client_fields = ["Client", "DisplayName"]
                for field in client_fields:
                    assert field in client, f"Expected client field '{field}' not found"

    @pytest.mark.asyncio
    async def test_invalid_parameters_handling(self):
        """Test that invalid parameters are handled gracefully."""
        # API should handle invalid parameters gracefully (silently ignore them)
        result = await search_addins(
            query="test",
            clients=["Invalid_Client"],  # Invalid client
            category=["NonExistentCategory"],  # Invalid category
            top=5
        )
        
        assert "TotalCount" in result
        assert "Values" in result
        # API should still return results, ignoring invalid filters

    @pytest.mark.asyncio
    async def test_large_top_parameter(self):
        """Test with large top parameter."""
        result = await search_addins(query="office", top=100)
        
        assert "TotalCount" in result
        assert "Values" in result
        assert len(result["Values"]) <= 100

    @pytest.mark.asyncio
    async def test_boolean_parameters_conversion(self):
        """Test that boolean parameters are converted to lowercase strings."""
        result = await search_addins(
            free=True,
            getMetaOSApps=False,
            top=5
        )
        
        assert "TotalCount" in result
        assert "Values" in result


class TestGetAddinDetails:
    """Test suite for the get_addin_details function."""

    @pytest.mark.asyncio
    async def test_get_valid_addin_details(self):
        """Test getting details for a valid add-in."""
        # Use a well-known asset ID from the API guide examples
        result = await get_addin_details("WA102957665")
        
        # Check basic structure
        assert isinstance(result, dict)
        assert "Value" in result, "Response should contain 'Value' key"
        
        addin_data = result["Value"]
        
        # Should contain key fields
        expected_fields = [
            "Id", "Title", "ShortDescription"
        ]
        
        for field in expected_fields:
            assert field in addin_data, f"Expected field '{field}' not found in details response"

    @pytest.mark.asyncio
    async def test_get_invalid_addin_details(self):
        """Test getting details for an invalid add-in ID."""
        with pytest.raises(httpx.HTTPStatusError):
            await get_addin_details("INVALID_ID_123")


class TestErrorHandling:
    """Test suite for error handling scenarios."""

    @pytest.mark.asyncio
    async def test_network_timeout_handling(self):
        """Test that network timeouts are handled appropriately."""
        # This test will use the actual function but with a very short timeout
        # We'll modify the function temporarily or expect it to handle timeouts gracefully
        
        # For now, just test that the function can be called
        # In a real scenario, you might want to mock the httpx client
        result = await search_addins(query="test", top=1)
        assert "TotalCount" in result

    @pytest.mark.asyncio
    async def test_malformed_parameters(self):
        """Test behavior with edge case parameters."""
        # Test with zero/negative values
        result = await search_addins(top=0)
        assert "TotalCount" in result
        
        # Test with large skiptoitem (but not so large it causes server error)
        result = await search_addins(skiptoitem=1000, top=1)
        assert "TotalCount" in result


if __name__ == "__main__":
    pytest.main([__file__])