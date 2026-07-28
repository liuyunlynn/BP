namespace MeetingBot;

/// <summary>
/// Strongly-typed configuration for the calling bot. Bind from the "Bot" section
/// of appsettings.json / user-secrets / environment variables.
/// </summary>
public sealed class BotOptions
{
    /// <summary>Azure AD tenant (directory) ID.</summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>The bot's Microsoft App (client) ID from the Azure Bot registration.</summary>
    public string AppId { get; set; } = string.Empty;

    /// <summary>The bot's client secret.</summary>
    public string AppSecret { get; set; } = string.Empty;

    /// <summary>
    /// Object ID (or UPN) of the user who will organize the scheduled meeting.
    /// Application-permission onlineMeeting creation requires an application
    /// access policy granting the app rights over this user.
    /// </summary>
    public string OrganizerUserId { get; set; } = string.Empty;

    /// <summary>
    /// Public HTTPS base URL where Microsoft Graph can reach this bot's signaling
    /// callback (e.g. https://your-tunnel.devtunnels.ms). Must terminate at /api/calling.
    /// </summary>
    public string BotBaseUrl { get; set; } = string.Empty;

    /// <summary>Friendly application name reported to the calling platform.</summary>
    public string AppName { get; set; } = "MeetingBotPoc";

    /// <summary>Relative path of the signaling callback controller.</summary>
    public string CallbackPath { get; set; } = "/api/calling";

    /// <summary>Full callback URI derived from <see cref="BotBaseUrl"/> and <see cref="CallbackPath"/>.</summary>
    public Uri CallbackUri => new Uri(new Uri(BotBaseUrl), CallbackPath);
}
