using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Azure.Core;
using Azure.Identity;

namespace MeetingBot;

/// <summary>
/// Schedules a Teams online meeting using the Microsoft Graph REST API
/// (POST /users/{id}/onlineMeetings). Uses raw HTTP + a client-credential token so
/// it stays independent of any Graph SDK model version used elsewhere in the app.
/// </summary>
public sealed class MeetingScheduler
{
    private static readonly string[] GraphScope = new[] { "https://graph.microsoft.com/.default" };

    private readonly BotOptions _options;
    private readonly HttpClient _httpClient;
    private readonly ClientSecretCredential _credential;
    private readonly ILogger<MeetingScheduler> _logger;

    public MeetingScheduler(BotOptions options, IHttpClientFactory httpClientFactory, ILogger<MeetingScheduler> logger)
    {
        _options = options;
        _httpClient = httpClientFactory.CreateClient(nameof(MeetingScheduler));
        _logger = logger;
        _credential = new ClientSecretCredential(options.TenantId, options.AppId, options.AppSecret);
    }

    /// <summary>
    /// Creates an online meeting owned by <see cref="BotOptions.OrganizerUserId"/> and
    /// returns its details (including the join URL the bot will use to join).
    /// </summary>
    public async Task<ScheduledMeeting> ScheduleMeetingAsync(string subject, DateTimeOffset startTime, DateTimeOffset endTime, CancellationToken cancellationToken)
    {
        AccessToken token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScope), cancellationToken).ConfigureAwait(false);

        string requestUri = $"https://graph.microsoft.com/v1.0/users/{_options.OrganizerUserId}/onlineMeetings";
        object payload = new
        {
            subject,
            startDateTime = startTime.ToUniversalTime().ToString("o"),
            endDateTime = endTime.ToUniversalTime().ToString("o"),
        };

        using HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUri);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        string body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError("Failed to create online meeting. Status '{StatusCode}'. Body '{Body}'.", (int)response.StatusCode, body);
            throw new InvalidOperationException($"Graph onlineMeetings create failed with status {(int)response.StatusCode}.");
        }

        ScheduledMeeting? meeting = JsonSerializer.Deserialize<ScheduledMeeting>(body);
        if (meeting is null || string.IsNullOrEmpty(meeting.JoinWebUrl))
        {
            throw new InvalidOperationException("Graph onlineMeetings response did not contain a joinWebUrl.");
        }

        _logger.LogInformation("Created online meeting '{MeetingId}' with join URL '{JoinUrl}'.", meeting.Id, meeting.JoinWebUrl);
        return meeting;
    }
}

/// <summary>Minimal projection of the Graph onlineMeeting resource.</summary>
public sealed class ScheduledMeeting
{
    [JsonPropertyName("id")]
    public string? Id { get; set; }

    [JsonPropertyName("subject")]
    public string? Subject { get; set; }

    [JsonPropertyName("joinWebUrl")]
    public string? JoinWebUrl { get; set; }
}
