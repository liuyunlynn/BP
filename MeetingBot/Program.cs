using Azure.Core;
using Azure.Identity;
using MeetingBot;
using Microsoft.Graph.Communications.Client;

WebApplicationBuilder builder = WebApplication.CreateBuilder(args);

// Bind BotOptions from configuration (appsettings.json / user-secrets / env vars).
BotOptions botOptions = new BotOptions();
builder.Configuration.GetSection("Bot").Bind(botOptions);
KustoOptions kustoOptions = new KustoOptions();
builder.Configuration.GetSection("Kusto").Bind(kustoOptions);

builder.Services.AddSingleton(botOptions);
builder.Services.AddSingleton(kustoOptions);
builder.Services.AddSingleton<TokenCredential>(_ => new DefaultAzureCredential());
builder.Services.AddHttpClient();
builder.Services.AddHttpClient<KustoJoinStatusService>();
builder.Services.AddSingleton<MeetingScheduler>();
builder.Services.AddSingleton<CallingBotService>();

WebApplication app = builder.Build();

// Signaling callback: Microsoft Graph POSTs call/roster notifications here.
app.MapPost(botOptions.CallbackPath, async (HttpContext context, CallingBotService bot) =>
{
    using HttpRequestMessage requestMessage = await context.Request.ToHttpRequestMessageAsync();
    HttpResponseMessage responseMessage = await bot.Client.ProcessNotificationAsync(requestMessage);
    await responseMessage.CopyToAsync(context.Response);
});

// Schedule a meeting, have the bot join it, and retain the app-to-meeting correlation for validation.
app.MapPost("/schedule-and-join", async (CallingBotService bot, MeetingScheduler scheduler, string? appId, int? minutes, string? subject, CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(appId))
    {
        return Results.BadRequest(new { error = "appId is required." });
    }

    DateTimeOffset start = DateTimeOffset.UtcNow.AddMinutes(1);
    DateTimeOffset end = start.AddMinutes(minutes ?? 30);

    ScheduledMeeting meeting = await scheduler.ScheduleMeetingAsync(subject ?? "Meeting Bot POC", start, end, cancellationToken);
    string joinWebUrl = meeting.JoinWebUrl
        ?? throw new InvalidOperationException("The scheduled meeting did not include a join URL.");
    await bot.JoinMeetingAsync(joinWebUrl, cancellationToken);
    bot.StoreVerificationMeeting(appId, joinWebUrl);

    return Results.Ok(new { joinWebUrl });
});

// Determine whether an ISV bot joined a meeting based on Kusto telemetry.
app.MapPost("/join-status", async (IsvBotJoinStatusRequest request, CallingBotService bot, KustoJoinStatusService joinStatus, CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(request.AppId))
    {
        return Results.BadRequest(new { error = "appId is required." });
    }

    if (!bot.TryGetVerificationMeeting(request.AppId, out VerificationMeetingRecord meeting))
    {
        return Results.Ok(new
        {
            meetingFound = false,
            meetingOpened = false,
            isJoined = false,
            eventTime = (DateTimeOffset?)null,
            meetingUrl = (string?)null,
        });
    }

    DateTimeOffset? eventTime = await joinStatus.GetJoinEventTimeAsync(
        meeting.ThreadId,
        request.AppId,
        cancellationToken);
    if (eventTime.HasValue)
    {
        bot.RemoveVerificationMeeting(request.AppId, meeting);
    }

    return Results.Ok(new
    {
        meetingFound = true,
        meetingOpened = meeting.MeetingOpened,
        isJoined = eventTime.HasValue,
        eventTime,
        meetingUrl = meeting.MeetingOpened ? null : meeting.JoinWebUrl,
    });
});

// Mark a verification meeting as opened (the bot has joined the meeting).
app.MapPost("/meeting-opened", (IsvBotJoinStatusRequest request, CallingBotService bot) =>
{
    if (string.IsNullOrWhiteSpace(request.AppId))
    {
        return Results.BadRequest(new { error = "appId is required." });
    }

    return bot.MarkVerificationMeetingOpened(request.AppId)
        ? Results.Ok(new { marked = true })
        : Results.NotFound(new { error = "No active verification meeting was found for the app ID." });
});

app.Run();

internal sealed record IsvBotJoinStatusRequest(string AppId);
