using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.AspNetCore;
using ModelContextProtocol.Server;

namespace Colossus.OutlookClassic;

/// <summary>
/// A parent-owned, loopback-only MCP transport. Its bearer token arrives on stdin;
/// closing that inherited channel revokes the endpoint and shuts down the helper.
/// </summary>
internal static class HttpCompanion
{
    internal static async Task<int> RunAsync(StaDispatcher dispatcher, OutlookReader outlook)
    {
        var token = await Console.In.ReadLineAsync();
        if (token is null || !ValidToken(token))
        {
            Console.Error.WriteLine("Companion bootstrap token is missing or invalid.");
            return 2;
        }

        var builder = WebApplication.CreateSlimBuilder(new WebApplicationOptions
        {
            Args = [],
            ContentRootPath = AppContext.BaseDirectory
        });
        builder.Configuration.Sources.Clear();
        builder.Configuration.AddInMemoryCollection();
        builder.Logging.ClearProviders();
        builder.WebHost.UseUrls("http://127.0.0.1:0");
        builder.Services.AddSingleton(dispatcher);
        builder.Services.AddSingleton(outlook);
        builder.Services.AddMcpServer()
            .WithHttpTransport(options => options.SessionMode = HttpServerSessionMode.Stateless)
            .WithTools<MailTools>();

        await using var app = builder.Build();
        var boundPort = 0;
        app.Use(async (context, next) =>
        {
            if (!IPAddress.IsLoopback(context.Connection.RemoteIpAddress ?? IPAddress.None)
                || context.Request.Host.Host != "127.0.0.1"
                || context.Request.Host.Port != boundPort
                || context.Request.Headers.ContainsKey("Origin"))
            {
                context.Response.StatusCode = StatusCodes.Status403Forbidden;
                return;
            }
            if (!Authorized(context.Request, token))
            {
                context.Response.StatusCode = StatusCodes.Status401Unauthorized;
                return;
            }
            await next(context);
        });
        app.MapMcp("/mcp");
        await app.StartAsync();
        var server = app.Services.GetRequiredService<IServer>();
        var addresses = server.Features.Get<IServerAddressesFeature>()?.Addresses;
        if (addresses is null || addresses.Count != 1 ||
            !Uri.TryCreate(addresses.Single(), UriKind.Absolute, out var address) ||
            address.Host != "127.0.0.1" || address.Port == 0)
        {
            await app.StopAsync();
            Console.Error.WriteLine("Companion did not bind one IPv4 loopback endpoint.");
            return 1;
        }
        boundPort = address.Port;
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            endpoint = $"http://127.0.0.1:{boundPort}/mcp",
            processId = Environment.ProcessId
        }));
        await Console.Out.FlushAsync();

        // The parent keeps stdin open for the lifetime of this exact process.
        // EOF is a narrow and reliable revocation signal, including parent crashes.
        await Console.In.ReadToEndAsync();
        await app.StopAsync();
        return 0;
    }

    private static bool ValidToken(string? token) =>
        token is { Length: 43 } && token.All(character =>
            character is >= 'A' and <= 'Z' or >= 'a' and <= 'z' or >= '0' and <= '9' or '-' or '_');

    private static bool Authorized(HttpRequest request, string token)
    {
        if (request.Headers.Authorization.Count != 1) return false;
        var header = request.Headers.Authorization.ToString();
        if (!header.StartsWith("Bearer ", StringComparison.Ordinal)) return false;
        var presented = header.AsSpan(7);
        if (presented.Length != token.Length) return false;
        var presentedBytes = Encoding.ASCII.GetBytes(presented.ToString());
        var expectedBytes = Encoding.ASCII.GetBytes(token);
        return CryptographicOperations.FixedTimeEquals(presentedBytes, expectedBytes);
    }
}
