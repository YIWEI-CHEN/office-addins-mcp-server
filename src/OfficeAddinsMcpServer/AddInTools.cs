using System.ComponentModel;
using System.Net.Http;
using System.Text.Encodings.Web;
using ModelContextProtocol.Server;

namespace OfficeAddinsMcpServer;

[McpServerToolType]
public static class AddinTools
{
    [McpServerTool(Name = "addin.search"), Description("Searches Office add‑ins by keyword and returns matching results.")]
    public static async Task<string> SearchAddins(HttpClient httpClient,
        [Description("Keyword for search")] string keyword)
    {
        string encodedKeyword = UrlEncoder.Default.Encode(keyword);
        string query         = Constants.SearchQuery.Replace("{keyword}", encodedKeyword);
        string url           = Constants.SearchUri + query;
        using HttpResponseMessage response = await httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }

    [McpServerTool(Name = "addin.details"), Description("Gets details for a specific Office add‑in by ID.")]
    public static async Task<string> GetAddinDetails(HttpClient httpClient,
        [Description("The identifier of the add‑in")] string addinId)
    {
        string encodedId = UrlEncoder.Default.Encode(addinId);
        string query     = Constants.DetailsQuery.Replace("{id}", encodedId);
        string url       = Constants.DetailsUri + query;
        using HttpResponseMessage response = await httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }
}
