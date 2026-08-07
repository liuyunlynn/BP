# Participant Metadata REST Test

## What was added

The bot already reads the live roster from the Graph Communications Calling SDK:

```csharp
ICall.Participants
```

A second method was added to call the Microsoft Graph REST API directly:

```http
GET https://graph.microsoft.com/v1.0/communications/calls/{callId}/participants
```

The POC exposes this through:

```http
GET /participants-rest/{callId}
```

This endpoint returns the unmodified Graph JSON response so fields such as
`metadata` remain available. It also writes each participant's raw `metadata`
value to the bot console.

The original endpoint remains available:

```http
GET /participants/{callId}
```

It reads the notification-backed `ICall.Participants` collection and returns the
existing flattened participant model.

## Important limitations

- The calling bot must be joined to the active call.
- `metadata` is an opaque value published by the participant.
- It has no fixed application-defined schema.
- Regular Teams users commonly omit it or publish an empty value.

## Prerequisites

- .NET 8 SDK
- A configured Azure Bot with Teams calling enabled
- Required Microsoft Graph application permissions and admin consent
- A public HTTPS callback URL
- Bot secrets configured as described in `README.md`
- The Azure Bot calling webhook and `Bot:BotBaseUrl` configured with the same
  dev tunnel URL

## Test procedure

### 1. Start the callback tunnel

Open a PowerShell window:

```powershell
devtunnel user login
devtunnel host -p 5275 --allow-anonymous
```

If the generated URL changed, update the local secret:

```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot
dotnet user-secrets set "Bot:BotBaseUrl" "https://<tunnel-id>-5275.<region>.devtunnels.ms"
```

Also update the Azure Bot calling webhook to:

```text
https://<tunnel-id>-5275.<region>.devtunnels.ms/api/calling
```

### 2. Start the bot

Open another PowerShell window:

```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot
dotnet run
```

Keep this window visible because participant metadata is logged here.

### 3. Create a meeting and join the bot

```powershell
curl.exe -s -X POST "http://localhost:5275/schedule-and-join?minutes=30&subject=MetadataTest"
```

The response contains a `joinWebUrl` and `callId`:

```json
{
  "meetingId": "<meeting-id>",
  "joinWebUrl": "<Teams meeting URL>",
  "callId": "<call-id>"
}
```

Open `joinWebUrl` and join the meeting as a real Teams user. Wait until the bot
and the user are both visible in the meeting.

### 4. Verify the existing SDK roster

Replace `<callId>` with the value returned above:

```powershell
curl.exe -s "http://localhost:5275/participants/<callId>"
```

This confirms that the bot is receiving participant roster notifications.

### 5. Test the direct Graph REST API

```powershell
curl.exe -s "http://localhost:5275/participants-rest/<callId>"
```

The command returns the raw response from:

```http
GET /v1.0/communications/calls/{callId}/participants
```

Inspect each participant object for:

```json
{
  "id": "<participant-id>",
  "metadata": "<opaque participant value>"
}
```

The `metadata` property may be absent, empty, or contain an opaque value.

### 6. Verify console output

The `dotnet run` window logs one line per participant:

```text
Call '<callId>' participant '<participantId>' metadata: <value>
```

If Graph does not include the field, the output is:

```text
Call '<callId>' participant '<participantId>' metadata: (not present)
```

### 7. Leave the meeting

```powershell
curl.exe -s -X POST "http://localhost:5275/leave/<callId>"
```

## Troubleshooting

| Symptom | Check |
|---|---|
| `dotnet` reports that no SDK is installed | Install the .NET 8 SDK and restart PowerShell. |
| The join request fails | Confirm Graph permissions, admin consent, Teams calling configuration, and the application access policy. |
| The roster is empty | Join the meeting as a real user and confirm Graph notifications reach `/api/calling`. |
| The REST endpoint returns `404` | Confirm the `callId` is current and the bot is still joined. |
| The REST endpoint returns `403` | Confirm the bot app has the required Graph calling permissions and admin consent. |
| `metadata` is missing or empty | This is expected when the participant does not publish metadata. |

## Relevant documentation

- ICall interface:
  https://microsoftgraph.github.io/microsoft-graph-comms-samples/docs/calls/Microsoft.Graph.Communications.Calls.ICall.html
- List call participants:
  https://learn.microsoft.com/en-us/graph/api/call-list-participants?view=graph-rest-1.0
- Participant resource:
  https://learn.microsoft.com/en-us/graph/api/resources/participant?view=graph-rest-1.0
