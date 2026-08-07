# Testing Runbook

## Status
- ✅ **Auth** — client-credentials token acquired.
- ✅ **Permissions** — `OnlineMeetings.ReadWrite.All`, `Calls.JoinGroupCall.All`, `Calls.JoinGroupCallAsGuest.All` consented.
- ✅ **Test A — Scheduling** — meeting created end-to-end via `POST /users/{id}/onlineMeetings`.
- ⬜ **Test B — Bot join + live roster** — steps below.

Organizer configured in user-secrets: `<ORGANIZER_UPN>`
(object id `<ORGANIZER_OBJECT_ID_GUID>`).

---

## Test B — bot joins & lists participants

### 1. Start a public tunnel (once)
```powershell
devtunnel user login                 # interactive browser sign-in
devtunnel host -p 5275 --allow-anonymous
```
Copy the printed **`https://<id>-5275.<region>.devtunnels.ms`** URL. Keep this window open.

> `--allow-anonymous` is required so Microsoft Graph can POST notifications without auth.

### 2. Point config at the tunnel
```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot
dotnet user-secrets set "Bot:BotBaseUrl" "https://<id>-5275.<region>.devtunnels.ms"
```

### 3. Configure the Azure Bot calling webhook (once per tunnel URL)
Azure Portal → your Bot resource (App ID `<APP_ID_GUID>`) → **Configuration** →
**Calling** tab → enable calling → **Webhook (for calling)** =
`https://<id>-5275.<region>.devtunnels.ms/api/calling` → Save.
Also ensure the **Microsoft Teams** channel is added with **Calling** enabled.

### 4. Run the bot
```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot
dotnet run
```
Watch this console — roster changes are logged here.

### 5. Schedule + join, then read the roster
```powershell
# schedule a new meeting AND have the bot join it (capture the returned callId):
curl.exe -s -X POST "http://localhost:5275/schedule-and-join?minutes=30&subject=RosterTest"

# a REAL user (the organizer) must now open the returned joinWebUrl and join,
# so the meeting session actually starts.

# list the live participants for THAT call (also prints to the bot console):
curl.exe -s "http://localhost:5275/participants/<callId>"

# call GET /v1.0/communications/calls/{id}/participants directly and inspect
# the raw Graph response, including metadata when a participant publishes it:
curl.exe -s "http://localhost:5275/participants-rest/<callId>"

# optional — every meeting the bot is currently in:
curl.exe -s "http://localhost:5275/calls"
```

### 6. Leave
```powershell
curl.exe -s -X POST "http://localhost:5275/leave/<callId>"
```

---

## Troubleshooting
| Symptom | Cause / fix |
|---|---|
| `/schedule-and-join` 403 on join | `Calls.JoinGroupCall.All` not consented, or Teams channel calling not enabled. |
| Bot never joins / no notifications | Webhook URL wrong, tunnel not `--allow-anonymous`, or `BotBaseUrl` mismatch. |
| `/participants/<callId>` stays `[]` | No human has joined yet (meeting session not started), or notifications not reaching `/api/calling`. |
| `/participants/<callId>` returns 404 | Wrong/expired `callId` — call `GET /calls` to list active call ids. |
| 400 "not a valid GUID" | `Bot:OrganizerUserId` must be the object-id GUID, not the UPN. |
