using MeetingBot;
using Microsoft.Graph.Communications.Client;

WebApplicationBuilder builder = WebApplication.CreateBuilder(args);

// Bind BotOptions from configuration (appsettings.json / user-secrets / env vars).
BotOptions botOptions = new BotOptions();
builder.Configuration.GetSection("Bot").Bind(botOptions);

builder.Services.AddSingleton(botOptions);
builder.Services.AddHttpClient();
builder.Services.AddSingleton<MeetingScheduler>();
builder.Services.AddSingleton<CallingBotService>();

WebApplication app = builder.Build();

// Health probe.
app.MapGet("/", () => Results.Ok(new { status = "ok", callback = botOptions.CallbackPath }));

// Signaling callback: Microsoft Graph POSTs call/roster notifications here.
app.MapPost(botOptions.CallbackPath, async (HttpContext context, CallingBotService bot) =>
{
    using HttpRequestMessage requestMessage = await context.Request.ToHttpRequestMessageAsync();
    HttpResponseMessage responseMessage = await bot.Client.ProcessNotificationAsync(requestMessage);
    await responseMessage.CopyToAsync(context.Response);
});

// Demo endpoint: schedule a meeting, then have the bot join it.
app.MapPost("/schedule-and-join", async (CallingBotService bot, MeetingScheduler scheduler, int? minutes, string? subject, CancellationToken cancellationToken) =>
{
    DateTimeOffset start = DateTimeOffset.UtcNow.AddMinutes(1);
    DateTimeOffset end = start.AddMinutes(minutes ?? 30);

    ScheduledMeeting meeting = await scheduler.ScheduleMeetingAsync(subject ?? "Meeting Bot POC", start, end, cancellationToken);
    string joinWebUrl = meeting.JoinWebUrl
        ?? throw new InvalidOperationException("The scheduled meeting did not include a join URL.");
    string threadId = JoinUrlParser.Parse(joinWebUrl).ChatInfo.ThreadId
        ?? throw new FormatException("The meeting join URL did not include a thread ID.");
    string callId = await bot.JoinMeetingAsync(joinWebUrl, cancellationToken);

    return Results.Ok(new
    {
        meetingId = meeting.Id,
        threadId,
        joinWebUrl,
        callId,
    });
});

// Have the bot join an already-existing meeting by join URL.
app.MapPost("/join", async (CallingBotService bot, string joinUrl, CancellationToken cancellationToken) =>
{
    string callId = await bot.JoinMeetingAsync(joinUrl, cancellationToken);
    return Results.Ok(new { callId });
});

// List the ids of every call the bot is currently joined to.
app.MapGet("/calls", (CallingBotService bot) => Results.Ok(new { callIds = bot.GetActiveCallIds() }));

// Determine whether an ISV bot joined a meeting. This returns a mock result until Kusto is available.
app.MapPost("/join-status", (IsvBotJoinStatusRequest request) =>
{
    if (string.IsNullOrWhiteSpace(request.ThreadId) || string.IsNullOrWhiteSpace(request.customApplicationTokenAppId))
    {
        return Results.BadRequest(new { error = "threadId and customApplicationTokenAppId are required." });
    }

    return Results.Ok(new
    {
        request.ThreadId,
        request.customApplicationTokenAppId,
        isJoined = true,
    });
});

// List (and print) the participant roster for ONE specific call.
app.MapGet("/participants/{callId}", (CallingBotService bot, string callId) =>
{
    IReadOnlyList<ParticipantSnapshot>? participants = bot.ListParticipants(callId);
    if (participants is null)
    {
        return Results.NotFound(new { error = $"No active call with id '{callId}'.", activeCalls = bot.GetActiveCallIds() });
    }

    bot.PrintRoster(callId);
    return Results.Ok(new { callId, participants });
});

// Leave ONE specific meeting by call id.
app.MapPost("/leave/{callId}", async (CallingBotService bot, string callId, CancellationToken cancellationToken) =>
{
    bool left = await bot.LeaveMeetingAsync(callId, cancellationToken);
    return left
        ? Results.Ok(new { callId, left = true })
        : Results.NotFound(new { error = $"No active call with id '{callId}'.", activeCalls = bot.GetActiveCallIds() });
});

app.Run();

internal sealed record IsvBotJoinStatusRequest(string ThreadId, string customApplicationTokenAppId);
