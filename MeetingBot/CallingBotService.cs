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
/// (roster-only) bot, and exposing the live participant roster.
///
/// The bot can be in <b>multiple meetings simultaneously</b>. Every active call
/// is tracked independently in <see cref="_calls"/>, keyed by its call id, so
/// roster reads and leaves always target one specific call.
/// </summary>
public sealed class CallingBotService
{
    private readonly BotOptions _options;
    private readonly ILogger<CallingBotService> _logger;
    private readonly ICommunicationsClient _client;

    /// <summary>All calls the bot is currently joined to, keyed by call id.</summary>
    private readonly ConcurrentDictionary<string, ICall> _calls = new(StringComparer.OrdinalIgnoreCase);

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

    /// <summary>
    /// Joins the meeting identified by <paramref name="joinWebUrl"/> as a
    /// service-hosted-media bot (no audio/video processing — roster only).
    /// Returns the id of the newly created call, used to scope later roster
    /// reads and the leave request.
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

        _calls[callId] = call;

        // Roster changes for THIS call print that call's roster only.
        call.Participants.OnUpdated += (sender, args) => PrintRoster(callId);

        // Auto-clean the entry when the call terminates so ended calls don't leak.
        call.OnUpdated += (sender, args) =>
        {
            if (sender.Resource?.State == CallState.Terminated)
            {
                _calls.TryRemove(callId, out _);
                _logger.LogInformation("Call '{CallId}' terminated; removed from active calls.", callId);
            }
        };

        _logger.LogInformation("Join request sent. Call id '{CallId}'. Active calls: {Count}.", callId, _calls.Count);
        return callId;
    }

    /// <summary>Leaves the specified meeting, if the bot is in it.</summary>
    /// <returns><c>true</c> if a matching active call was found and left.</returns>
    public async Task<bool> LeaveMeetingAsync(string callId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(callId) || !_calls.TryRemove(callId, out ICall? call))
        {
            return false;
        }

        await call.DeleteAsync().ConfigureAwait(false);
        _logger.LogInformation("Left meeting. Call id '{CallId}'. Active calls: {Count}.", callId, _calls.Count);
        return true;
    }

    /// <summary>The ids of all calls the bot is currently joined to.</summary>
    public IReadOnlyList<string> GetActiveCallIds() => _calls.Keys.ToArray();

    /// <summary>
    /// Returns a snapshot of the participants in the specified call only,
    /// including the bot itself (which appears as an application participant).
    /// Returns <c>null</c> if the bot is not in a call with that id.
    /// </summary>
    public IReadOnlyList<ParticipantSnapshot>? ListParticipants(string callId)
    {
        if (string.IsNullOrWhiteSpace(callId) || !_calls.TryGetValue(callId, out ICall? call))
        {
            return null;
        }

        List<ParticipantSnapshot> snapshots = new List<ParticipantSnapshot>();
        foreach (IParticipant participant in call.Participants)
        {
            snapshots.Add(ParticipantSnapshot.From(participant.Resource));
        }

        return snapshots;
    }

    /// <summary>Prints the roster of the specified call to the logger.</summary>
    public void PrintRoster(string callId)
    {
        IReadOnlyList<ParticipantSnapshot>? participants = ListParticipants(callId);
        if (participants is null)
        {
            _logger.LogWarning("PrintRoster called for unknown call id '{CallId}'.", callId);
            return;
        }

        _logger.LogInformation("Call '{CallId}' roster: {Count} participant(s).", callId, participants.Count);
        foreach (ParticipantSnapshot participant in participants)
        {
            _logger.LogInformation(
                "  - '{DisplayName}' (id '{Id}', kind '{Kind}', muted '{Muted}', inLobby '{InLobby}').",
                participant.DisplayName,
                participant.Id,
                participant.Kind,
                participant.IsMuted,
                participant.IsInLobby);
        }
    }
}

/// <summary>A flattened, printable view of a meeting participant.</summary>
public sealed record ParticipantSnapshot(string DisplayName, string Id, string Kind, bool IsMuted, bool IsInLobby)
{
    public static ParticipantSnapshot From(Participant? resource)
    {
        IdentitySet? identity = resource?.Info?.Identity;
        Identity? user = identity?.User;
        Identity? application = identity?.Application;
        Identity? device = identity?.Device;

        bool isApplication = application is not null
            || (identity?.AdditionalData?.ContainsKey("applicationInstance") ?? false);

        Identity? effective = user ?? application ?? device;
        string displayName = effective?.DisplayName ?? "(unknown)";
        string id = effective?.Id ?? resource?.Id ?? "(no-id)";
        string kind = isApplication ? "application/bot" : user is not null ? "user" : "other";

        return new ParticipantSnapshot(
            displayName,
            id,
            kind,
            resource?.IsMuted ?? false,
            resource?.IsInLobby ?? false);
    }
}
