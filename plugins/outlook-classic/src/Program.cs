using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using System.Text.Json;
using System.Reflection;

namespace Colossus.OutlookClassic;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args is not ["--stdio"] and not ["--probe"] and not ["--version"])
        {
            Console.Error.WriteLine("Usage: outlook-classic-mcp --stdio | --probe | --version");
            return 2;
        }
        if (args[0] == "--version")
        {
            var version = typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion;
            Console.WriteLine($"outlook-classic-mcp {version}");
            return 0;
        }

        using var dispatcher = new StaDispatcher();
        var outlook = new OutlookReader();
        if (args[0] == "--probe")
        {
            try
            {
                var result = await dispatcher.InvokeAsync(outlook.GetStatus, CancellationToken.None);
                Console.WriteLine(JsonSerializer.Serialize(result));
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(JsonSerializer.Serialize(new { connected = false, error = SafeError.Describe(ex) }));
                return 1;
            }
        }

        // Do not load ambient appsettings or command-line configuration into this plugin.
        var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
        {
            DisableDefaults = true,
            ContentRootPath = AppContext.BaseDirectory
        });
        builder.Logging.ClearProviders(); // Mail content must never enter SDK debug logs.
        builder.Services.AddSingleton(dispatcher);
        builder.Services.AddSingleton(outlook);
        builder.Services.AddMcpServer().WithStdioServerTransport().WithTools<MailTools>();
        await builder.Build().RunAsync();
        return 0;
    }
}
