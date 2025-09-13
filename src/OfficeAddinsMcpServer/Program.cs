using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;

namespace OfficeAddinsMcpServer;

internal class Program
{
    private static async Task Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);
        builder.Logging.AddConsole(options =>
        {
            // Redirect logs to stderr to avoid interfering with stdio protocol output
            options.LogToStandardErrorThreshold = LogLevel.Trace;
        });

        builder.Services.AddHttpClient();

        builder.Services
            .AddMcpServer(opts =>
            {
                opts.ServerInfo.Name    = "OfficeAddinsMcpServer";
                opts.ServerInfo.Version = "1.0.0";
            })
            .WithStdioServerTransport()  // local development
            .WithToolsFromAssembly();    // scan for tool definitions

        await builder.Build().RunAsync();
    }
}