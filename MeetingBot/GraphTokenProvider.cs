using System.Threading;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph.Communications.Client.Authentication;

namespace MeetingBot;

/// <summary>
/// Supplies OAuth access tokens to the Graph Communications calling platform using
/// the bot's app registration (client-credentials). Wired via
/// ICommunicationsClientBuilder.SetAuthentication so the SDK also handles inbound
/// notification validation for us.
/// </summary>
public sealed class GraphTokenProvider : ITokenProvider
{
    private static readonly string[] GraphScope = new[] { "https://graph.microsoft.com/.default" };

    private readonly ClientSecretCredential _credential;

    public GraphTokenProvider(BotOptions options)
    {
        _credential = new ClientSecretCredential(options.TenantId, options.AppId, options.AppSecret);
    }

    public async Task<string> AcquireTokenAsync(string tenant)
    {
        AccessToken token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScope), CancellationToken.None).ConfigureAwait(false);
        return token.Token;
    }
}
