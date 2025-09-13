namespace OfficeAddinsMcpServer;

internal static class Constants
{
    // Environment variable names
    private const string SearchUriEnv    = "OFFICE_ADDIN_SEARCH_URI";
    private const string SearchQueryEnv  = "OFFICE_ADDIN_SEARCH_QUERY";
    private const string DetailsUriEnv   = "OFFICE_ADDIN_DETAILS_URI";
    private const string DetailsQueryEnv = "OFFICE_ADDIN_DETAILS_QUERY";

    // Default values (replace with your actual API endpoints and patterns)
    private const string DefaultSearchUri    = "https://api.example.com/addins/search";
    private const string DefaultSearchQuery  = "?keyword={keyword}&page=1&size=10";
    private const string DefaultDetailsUri   = "https://api.example.com/addins/details";
    private const string DefaultDetailsQuery = "?id={id}";

    public static string SearchUri   => Environment.GetEnvironmentVariable(SearchUriEnv)   ?? DefaultSearchUri;
    public static string SearchQuery => Environment.GetEnvironmentVariable(SearchQueryEnv) ?? DefaultSearchQuery;
    public static string DetailsUri  => Environment.GetEnvironmentVariable(DetailsUriEnv)  ?? DefaultDetailsUri;
    public static string DetailsQuery=> Environment.GetEnvironmentVariable(DetailsQueryEnv)?? DefaultDetailsQuery;
}
