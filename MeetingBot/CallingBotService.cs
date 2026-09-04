using System.Collections.Concurrent;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Client;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Graph.Models;

namespace MeetingBot;

/// <summary>
/// Wraps the Microsoft Graph Communications calling client. Responsible for:
/// building the stateful client, joining meetings as a service-hosted-media
/// bot, and processing call notifications.
/// </summary>
public sealed class CallingBotService
{
    private static readonly TimeSpan VerificationMeetingLifetime = TimeSpan.FromHours(24);
    private readonly BotOptions _options;
    private readonly ILogger<CallingBotService> _logger;
    private readonly ICommunicationsClient _client;
    private readonly ConcurrentDictionary<string, VerificationMeetingRecord> _verificationMeetings =
        new(StringComparer.OrdinalIgnoreCase);

    public CallingBotService(BotOptions options, ILogger<CallingBotService> logger)
    {
        _options = options;
        _logger = logger;

        IGraphLogger graphLogger = new GraphLogger(options.AppName);
        ICommunicationsClientBuilder builder = new CommunicationsClientBuilder(options.AppName, options.AppId, graphLogger);
        builder.SetAuthentication(options.AppId, new GraphTokenProvider(options));
        builder.SetNotificationUrl(options.CallbackUri);
        builder.SetServiceBaseUrl(new Uri("https://graph.microsoft.com/v1.0"));
        _client = builder.Build();
    }

    /// <summary>The underlying client, used by the notification controller.</summary>
    public ICommunicationsClient Client => _client;

    public void StoreVerificationMeeting(string appId, string joinWebUrl)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(appId);
        ArgumentException.ThrowIfNullOrWhiteSpace(joinWebUrl);

        RemoveExpiredVerificationMeetings();
        string threadId = JoinUrlParser.Parse(joinWebUrl).ChatInfo.ThreadId
            ?? throw new FormatException("The meeting join URL did not include a thread ID.");
        _verificationMeetings[appId] = new VerificationMeetingRecord(
            threadId,
            joinWebUrl,
            DateTimeOffset.UtcNow,
            MeetingOpened: false);
    }

    public bool TryGetVerificationMeeting(string appId, out VerificationMeetingRecord meeting)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(appId);

        RemoveExpiredVerificationMeetings();
        return _verificationMeetings.TryGetValue(appId, out meeting!);
    }

    public void RemoveVerificationMeeting(string appId, VerificationMeetingRecord meeting)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(appId);
        ArgumentNullException.ThrowIfNull(meeting);

        ((ICollection<KeyValuePair<string, VerificationMeetingRecord>>)_verificationMeetings)
            .Remove(new KeyValuePair<string, VerificationMeetingRecord>(appId, meeting));
    }

    public bool MarkVerificationMeetingOpened(string appId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(appId);

        RemoveExpiredVerificationMeetings();
        while (_verificationMeetings.TryGetValue(appId, out VerificationMeetingRecord? meeting))
        {
            if (meeting.MeetingOpened)
            {
                return true;
            }

            if (_verificationMeetings.TryUpdate(appId, meeting with { MeetingOpened = true }, meeting))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Joins the meeting identified by <paramref name="joinWebUrl"/> as a
    /// service-hosted-media bot (no audio/video processing).
    /// Returns the id of the newly created call.
    /// </summary>
    public async Task<string> JoinMeetingAsync(string joinWebUrl, CancellationToken cancellationToken)
    {
        (ChatInfo chatInfo, MeetingInfo meetingInfo) = JoinUrlParser.Parse(joinWebUrl);

        JoinMeetingParameters joinParameters = new JoinMeetingParameters(
            chatInfo,
            meetingInfo,
            Array.Empty<Modality>(),
            prefetchPrompts: null,
            isInteractiveRosterEnabled: true,
            optIntoDeltaRoster: null,
            isParticipantInfoUpdatesEnabled: true)
        {
            TenantId = _options.TenantId,
        };

        ICall call = await _client.Calls().AddAsync(joinParameters, Guid.NewGuid(), cancellationToken).ConfigureAwait(false);
        string callId = call.Id;

        _logger.LogInformation("Join request sent. Call id '{CallId}'.", callId);
        return callId;
    }

    private void RemoveExpiredVerificationMeetings()
    {
        DateTimeOffset expirationThreshold = DateTimeOffset.UtcNow - VerificationMeetingLifetime;
        foreach ((string appId, VerificationMeetingRecord meeting) in _verificationMeetings)
        {
            if (meeting.CreatedAt <= expirationThreshold)
            {
                ((ICollection<KeyValuePair<string, VerificationMeetingRecord>>)_verificationMeetings)
                    .Remove(new KeyValuePair<string, VerificationMeetingRecord>(appId, meeting));
            }
        }
    }
}

public sealed record VerificationMeetingRecord(
    string ThreadId,
    string JoinWebUrl,
    DateTimeOffset CreatedAt,
    bool MeetingOpened);
