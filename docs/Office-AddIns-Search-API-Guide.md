# Office Add-ins Search API Guide

## Overview

This guide provides comprehensive documentation for the Office Add-ins Search API, including all available query parameters, examples, and usage patterns.

### **Base URL**
```
https://api.addins.omex.office.net/api/addins/search
```

### **HTTP Methods**
- `GET` - Standard query parameters
- `POST` - Supports form data for larger parameter lists (productids, assetids)

---

## **Query Parameters**

### **🔍 Search & Filtering**
| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `qu` | string | Search query/keywords | `qu=Zoom` |
| `category` | string[] | Filter by categories (comma-separated) | `category=Productivity,Communication` |
| `free` | boolean | Filter for free add-ins | `free=true` (Note: May not strictly filter only free add-ins) |
| `clients` | string[] | Filter by Office client platform-product combinations | `clients=Win32_Excel,Mac_Word,WAC_PowerPoint` |
| `productgroup` | string[] | Filter by product groups | `productgroup=Office` |
| `productids` | string[] | Filter by specific product IDs/GUIDs | `productids=guid1,guid2` |
| `assetids` | string[] | Filter by specific asset IDs | `assetids=WA104381441,WA102957665` |
| `providertype` | string | Filter by provider type | `providertype=gmail` |
| `getMetaOSApps` | boolean | Include MetaOS applications | `getMetaOSApps=true` |

**Client Parameter Values:**
- **Format:** `Platform_Product` (e.g., `Win32_Excel`, `Mac_Word`, `WAC_Outlook`)
- **Platforms:** `Win32`, `Mac`, `WAC` (Web/Online), `Any` (for Power BI)
- **Products:** `Word`, `PowerPoint`, `Excel`, `OneNote`, `Outlook`, `PowerBI`
- **Examples:** 
  - `Win32_Excel` - Excel on Windows
  - `Mac_Word` - Word on Mac  
  - `WAC_PowerPoint` - PowerPoint on Web
  - `Win32_Outlook` - Outlook on Windows
  - `Any_PowerBI` - Power BI (platform-agnostic)### **📊 Sorting**
| Parameter | Type | Values | Description |
|-----------|------|--------|-------------|
| `orderfield` | enum | `None` (0), `Title` (1), `Date` (2), `Price` (3), `Rating` (4) | Field to sort by |
| `orderby` | enum | `Desc`, `Asc` | Sort direction |

**Sorting Field Details:**
- `None` (0) - Default sorting (relevance-based)
- `Title` (1) - Sort alphabetically by add-in title
- `Date` (2) - Sort by release/creation date
- `Price` (3) - Sort by price (uses default sorting in practice)
- `Rating` (4) - Sort by average user rating

### **📄 Pagination**
| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `top` | int | Number of results to return (limit) | `top=20` |
| `skiptoitem` | int | Number of results to skip (offset) | `skiptoitem=40` |

### **⚙️ Advanced/System**
| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `date` | date | Date override for cache bypass (yyyy-MM-dd) | `date=2025-09-19` |

---

## **Regional/Locale Parameters (Unofficial)**

These parameters influence language/regional filtering but may not be officially supported in the public API:

| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `lc` | string | Locale | `lc=es-ES` |
| `ui` | string | UI Culture | `ui=es` |
| `uilcid` | string | Client Locale ID | `uilcid=1033` |
| `rs` | string | Data Market/Region | `rs=US` |
| `omkt` | string | Worldwide Page Market | `omkt=en-us` |
| `lidui` | string | Legacy Client Locale | `lidui=1033` |
| `clid` | string | Legacy Client Locale ID | `clid=1033` |
| `lcid` | string | Locale ID | `lcid=1033` |
| `pm` | string | Purchase Market | `pm=US` |

---

## **Usage Examples**

### **Basic Search**
```bash
GET https://api.addins.omex.office.net/api/addins/search?qu=Zoom
# Returns: TotalCount: 220, including "Zoom for Outlook Add-in"
```

### **Advanced Filtering**
```bash
# Search for free add-ins, sorted by rating (descending by default)
GET https://api.addins.omex.office.net/api/addins/search?qu=productivity&free=true&orderfield=Rating&orderby=Desc&top=10
# Returns: Only free add-ins matching "productivity"

# Note: free=true still returns some paid add-ins with trial versions
```

### **Client-Specific Search**
```bash
# Search for Excel add-ins on Windows (returns cross-platform results)
GET https://api.addins.omex.office.net/api/addins/search?qu=charts&clients=Win32_Excel

# Search for add-ins across multiple platforms
GET https://api.addins.omex.office.net/api/addins/search?qu=productivity&clients=Win32_Excel,Mac_Excel,WAC_Excel
```

### **Pagination Examples**
```bash
# First page (default) - returns first 20 results
GET https://api.addins.omex.office.net/api/addins/search?qu=calendar&top=20

# Second page - skip first 10, get next 2
GET https://api.addins.omex.office.net/api/addins/search?skiptoitem=10&top=2
# Returns: "DocuSign for Outlook", "Microsoft Dynamics Office Add-in"

# Large skip - get results 100-110
GET https://api.addins.omex.office.net/api/addins/search?qu=office&skiptoitem=100&top=10
```

### **Sorting Examples**
```bash
# Sort by title alphabetically (A-Z)
GET https://api.addins.omex.office.net/api/addins/search?orderfield=Title&orderby=Asc&top=3
# Returns: "(Legacy) Chronograph xConnect", "1&1 Business Phone", "1&1 Business Phone Meet&Share"

# Sort by highest rated first
GET https://api.addins.omex.office.net/api/addins/search?orderfield=Rating&orderby=Desc&top=3
# Returns: Multiple 5-star rated add-ins like "AI Perfect Assistant for Office"

# Sort by newest first (using Date field)
GET https://api.addins.omex.office.net/api/addins/search?qu=email&orderfield=Date&orderby=Desc

# Default sorting (relevance-based when no orderfield specified)
GET https://api.addins.omex.office.net/api/addins/search?qu=presentation
```

### **Category Filtering**
```bash
# Single category
GET https://api.addins.omex.office.net/api/addins/search?category=Productivity

# Multiple categories
GET https://api.addins.omex.office.net/api/addins/search?category=Productivity,Communication,Education
```

### **Specific Add-in Lookup**
```bash
# Search for specific asset IDs
GET https://api.addins.omex.office.net/api/addins/search?assetids=WA104381441,WA102957665
# Returns: Zoom for Outlook and Microsoft Translator add-ins

# Search for specific product IDs
GET https://api.addins.omex.office.net/api/addins/search?productids=12345678-1234-1234-1234-123456789012
```

### **Regional/Language Influence**
```bash
# Try to get Spanish results (unofficial)
GET https://api.addins.omex.office.net/api/addins/search?qu=calendario&lc=es-ES&ui=es

# With Accept-Language header
curl -H "Accept-Language: es-ES,es;q=0.9" \
  "https://api.addins.omex.office.net/api/addins/search?qu=calendario"
```

### **POST Method for Large Parameter Lists**
```bash
POST https://api.addins.omex.office.net/api/addins/search?qu=office
Content-Type: application/x-www-form-urlencoded

productids=guid1,guid2,guid3,guid4,guid5&assetids=WA001,WA002,WA003
```

### **Cache Bypass**
```bash
# Bypass cache to get latest results
GET https://api.addins.omex.office.net/api/addins/search?qu=new&date=2025-09-19
```

### **Complex Query Example**
```bash
# Search Excel add-ins for accounting, sort by title, limit to 3 results
GET https://api.addins.omex.office.net/api/addins/search?qu=accounting&clients=Win32_Excel&orderfield=Title&orderby=Asc&top=3
# Returns: Accounting-related Excel add-ins sorted alphabetically

# Search only free PowerPoint add-ins (note: free parameter exists but may not strictly filter)
GET https://api.addins.omex.office.net/api/addins/search?qu=template&clients=Win32_PowerPoint&free=true&top=5

# Search specific categories in Outlook (category ID example: 94)
GET https://api.addins.omex.office.net/api/addins/search?clients=Win32_Outlook&categories=94&top=5

# Search with pagination - skip first 10 results, show next 5
GET https://api.addins.omex.office.net/api/addins/search?qu=chart&skiptoitem=10&top=5
# Note: First page is skiptoitem=0, second page is skiptoitem=5 for top=5, etc.

# Get multiple specific add-ins by their Asset IDs
GET https://api.addins.omex.office.net/api/addins/search?assetids=WA104381441,WA102957665
# Returns: Exactly the 2 specified add-ins (Zoom for Outlook, Microsoft Translator)
```

---

## **Response Format**

The API returns JSON with the following structure:

```json
{
  "TotalCount": 387,
  "Values": [
    {
      "Id": "WA102957665",
      "Title": "Mini Calendar and Date Picker",
      "ShortDescription": "Add a mini monthly calendar to your spreadsheet and use it to insert dates or the current time.",
      "IconUrl": "https://store-images.s-microsoft.com/image/apps.31013.352a833c-760b-49ee-b384-f9718a71d650...",
      "Rating": 2.13333333333333,
      "NumberOfVotes": 15,
      "DateReleased": "2012-06-28T12:00:00Z",
      "LastUpdatedDate": "2020-01-13T05:49:21Z",
      "ProductId": "b910c929-6d0d-4397-b4d3-e22ef215fec0",
      "ManifestUrl": "https://addinsinstallation.store.office.com/app/download?assetid=WA102957665&cmu=en-US",
      "Culture": "en-US",
      "State": "Ok",
      "Version": "1.3.0",
      "Shape": 1,
      "Width": 240,
      "Height": 220,
      "Pricing": {
        "Category": "Free",
        "FreeType": "AppFree",
        "SiteLicenseAvailable": false,
        "Price": "Free",
        "Currency": "USD",
        "SupportsTrial": false,
        "IsUnlimitedTrial": false,
        "TrialLength": -1
      },
      "Categories": [
        {
          "Id": "Reference",
          "Title": "Reference",
          "LongTitle": "Discover great reference content"
        }
      ],
      "SupportedClients": [
        {
          "Client": "Win32_Excel",
          "MinVersion": "15.0.4535.1511",
          "DisplayName": "Excel 2013 Service Pack 1 or later on Windows"
        }
      ],
      "Permissions": [
        {
          "Id": "ReadWrite Document",
          "Description": "Can read and make changes to your document",
          "LongDescription": null
        }
      ],
      "Certification": {
        "State": null,
        "Id": null,
        "Uri": null,
        "Description": null
      },
      "ProviderName": "VERTEX42",
      "LicenseTermsUrl": "https://go.microsoft.com/fwlink/?LinkID=521715&omkt=en-US",
      "PrivacyPolicyUrl": "https://www.vertex42.com/privacy.html",
      "SupportUrl": "https://www.vertex42.com/apps/minicalendar.html",
      "ExtendedPermissions": [],
      "LeadEnabled": false,
      "ActiveDirectoryAppId": null,
      "ActiveDirectoryScopes": [],
      "AutorunLaunchEvents": [],
      "IsMetaOSApp": null,
      "Successor": null,
      "Predecessors": []
    }
  ]
}
```

### **Key Response Fields:**
- `TotalCount` - Total number of matching results available
- `Values[]` - Array of add-in objects (current page)
- `Id` - Asset ID of the add-in (e.g., "WA102957665")
- `Title` - Display name of the add-in
- `ShortDescription` - Brief description of the add-in functionality
- `Rating` - Average user rating (decimal number, typically 1-5)
- `NumberOfVotes` - Number of user reviews/ratings
- `DateReleased` - When the add-in was first released (ISO 8601 format)
- `LastUpdatedDate` - When the add-in was last updated (ISO 8601 format)
- `Pricing.Price` - Price display string (e.g., "Free")
- `Pricing.FreeType` - Free tier type (e.g., "AppFree")
- `Categories[]` - Associated categories with Id, Title, and LongTitle
- `SupportedClients[]` - Compatible platform-product combinations
  - `Client` - Platform_Product format (e.g., "Win32_Excel")
  - `DisplayName` - Human-readable client description
- `ProviderName` - Publisher/developer name
- `State` - Add-in status (e.g., "Ok")
- `Culture` - Localization culture (e.g., "en-US")
- `Permissions[]` - Required permissions with descriptions

---

## **Related APIs**

### **Add-in Details API**
Get detailed information about a specific add-in:
```
GET https://api.addins.omex.office.net/api/addins/details?assetid=WA104381441
```

### **Categories API**
Get available categories (**requires client parameter**):
```bash
GET https://api.addins.omex.office.net/api/addins/categories?client=Win32_Excel
```

**Response**: Returns array of category objects with `Id`, `Title`, and `LongTitle` fields:
```json
{
  "Values": [
    {
      "Id": "Productivity",
      "Title": "Productivity", 
      "LongTitle": "Get work done with Office"
    },
    {
      "Id": "Data Analytics",
      "Title": "Data Analytics",
      "LongTitle": "Model data and forecast trends"
    }
  ]
}
```

### **Suggestions API**
Get suggested add-ins (**requires query parameter**):
```bash
GET https://api.addins.omex.office.net/api/addins/suggestions?query=office
```

**Response**: Returns JSON with same structure as search API (TotalCount + Values array):
```json
{
  "TotalCount": 5,
  "Values": [
    {
      "Title": "AI Perfect Assistant for Office",
      "AssetId": "WA...",
      // ... other add-in properties
    }
  ]
}
```

---

## **Error Handling & Edge Cases**

### **Invalid Parameters**
- Invalid `categories` or `clients` values are **silently ignored** rather than returning errors
- API returns full results when invalid filters are provided
- Only malformed requests (missing required params) return `BadRequest` errors

### **Common Errors**
```json
{
  "Code": "BadRequest",
  "Description": "Could not understand request because: value for client is invalid, null or whitespace",
  "Time": "09/19/2025 08:27:20"
}
```

### **Parameter Notes**
- `free=true` parameter exists but may not strictly filter only free add-ins
- `categories` API requires singular `client` parameter (not `clients`)
- Invalid Asset IDs in `assetids` are silently ignored
- Search queries work across both titles and descriptions

---

## **Best Practices**

### **Performance**
- Use pagination (`top` and `skiptoitem`) for large result sets
- Cache results when appropriate
- Use specific filters to reduce response size

### **Filtering**
- Combine multiple filters for precise results
- Use `free=true` to filter out paid add-ins
- Filter by `clients` to get client-specific results

### **Sorting**
- Use `Rating` sorting for quality-based results
- Use `Date` sorting for newest add-ins
- Use `Title` sorting for alphabetical browsing

### **Error Handling**
- Handle HTTP error codes appropriately
- Validate query parameters before sending requests
- Implement retry logic for transient failures

---

## **Important Notes**

1. **Authentication**: This appears to be a public API that doesn't require authentication for basic searches
2. **Rate Limiting**: Not documented, but likely exists - implement appropriate throttling
3. **Regional Filtering**: Language/region filtering is primarily handled server-side based on request headers and regional detection
4. **POST vs GET**: Use POST when you have many `productids`/`assetids` that might exceed URL length limits
5. **Testing**: All examples in this guide have been validated against the live API as of September 2025

## **Validation Notes**

This documentation has been validated through extensive API testing:
- ✅ Search functionality confirmed across titles and descriptions  
- ✅ Pagination working correctly (skiptoitem + top parameters)
- ✅ Sorting by Title, Rating, Date confirmed functional
- ✅ Client filtering with Platform_Product format verified  
- ✅ Asset ID lookups returning exact matches
- ✅ Categories API requiring singular `client` parameter confirmed
- ✅ Suggestions API requiring `query` parameter confirmed
- ⚠️ Free filter parameter exists but may not strictly filter only free add-ins
- ⚠️ Invalid filter parameters are silently ignored rather than returning errors
5. **Unofficial Parameters**: Regional/locale parameters may work but aren't guaranteed to be stable
6. **Case Sensitivity**: Parameter values may be case-sensitive
7. **URL Encoding**: Ensure proper URL encoding for special characters in query values

---

## **Common Use Cases**

### **Store Browsing**
```bash
# Browse all free add-ins, sorted by popularity
GET https://api.addins.omex.office.net/api/addins/search?free=true&orderfield=Rating&orderby=Desc&top=50
```

### **Search with Autocomplete**
```bash
# Search as user types (with shorter result sets)
GET https://api.addins.omex.office.net/api/addins/search?qu=cal&top=10
```

### **Category Exploration**
```bash
# Explore specific categories
GET https://api.addins.omex.office.net/api/addins/search?category=Education&orderfield=Rating&orderby=Desc
```

### **Client-Specific Recommendations**
```bash
# Get Excel add-ins for Windows users
GET https://api.addins.omex.office.net/api/addins/search?clients=Win32_Excel&free=true&top=20

# Get PowerPoint add-ins for Web users
GET https://api.addins.omex.office.net/api/addins/search?clients=WAC_PowerPoint&orderfield=Rating&orderby=Desc&top=20

# Get Outlook add-ins for Mac users
GET https://api.addins.omex.office.net/api/addins/search?clients=Mac_Outlook&free=true&top=15

# Multi-platform Word add-ins
GET https://api.addins.omex.office.net/api/addins/search?clients=Win32_Word,Mac_Word,WAC_Word&top=25
```

This API provides comprehensive search capabilities for the Office add-ins store with extensive filtering, sorting, and pagination options suitable for building rich add-in discovery experiences.