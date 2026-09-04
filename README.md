# Teams Meeting Calling Bot — POC & Onboarding Guide

Welcome! 👋 This repo is a **standalone C#/.NET 8 proof-of-concept** that proves out a full
Microsoft Teams calling-bot flow. This README is the **onboarding document** — read it
top-to-bottom and you'll understand *what* the demo does, *how* it's built, and *how to run
and test it yourself*.

---

## 1. What the demo does

The POC demonstrates the complete lifecycle in four steps:

1. **Schedule** a Teams meeting via Microsoft Graph REST.
2. Have a **bot / application join** that meeting using the Graph Communications Calling SDK.
3. **List** the live meeting participants — *including the bot itself*.
4. **Print** the participant roster to the console and return it as JSON.

It answers the underlying question every new hire asks:
> *Can a Microsoft Graph bot join a Teams meeting and enumerate who's in the room — live?*

**Answer: yes — but only while a bot is actually joined to the call.** This POC proves it end-to-end.

---

## 2. Feasibility verdict (the "why")

| Capability | Feasible? | How |
|---|---|---|
| Schedule a meeting | ✅ Yes | `POST /users/{id}/onlineMeetings` (app perm `OnlineMeetings.ReadWrite.All` + application access policy) |
| Bot **joins** the meeting programmatically | ✅ Yes | Graph Communications Calling SDK — `JoinMeetingParameters` → `ICommunicationsClient.Calls().AddAsync(...)` |
| List **live** participants (incl. the bot) | ✅ Yes — **only** while a bot is in the call | `ICall.Participants`, populated by the calling platform's roster notifications |
| List participants **without** a bot | ⚠️ Limited | Post-meeting `attendanceReports` only — **excludes bots**, needs an organizer to have joined. No plain-REST live roster exists. |

**Key insight for new hires:** There is **no** Microsoft Graph REST endpoint that returns
the live in-meeting participant list. Live roster data arrives as **call notifications** to an
application that has *joined the call* as a calling bot. Having the bot join is the *only* way
to get a live roster that includes the bot itself.

---

## 3. Architecture at a glance

```
                    ┌─────────────────────────────────────────────┐
                    │             MeetingBot (ASP.NET Core)         │
                    │                                               │
  POST /schedule-and-join ──► MeetingScheduler ──► Graph REST       │
                    │            (create onlineMeeting)             │
                    │                    │                          │
                    │                    ▼                          │
                    │           CallingBotService ──► Calling SDK   │
                    │           (join call, hold roster)   │        │
                    │                                       ▼        │
   GET /participants ──────────► ListParticipants() ──► ICall.Participants
                    │                                       ▲        │
                    │   POST /api/calling ◄─── Graph call notifications
                    └───────────────────────────┬───────────────────┘
                                                 │
             Azure Bot "Calling webhook"  ───────┘  (points at {tunnel}/api/calling)
```

- **Graph → bot** notifications must reach a **public HTTPS URL** (a dev tunnel in dev).
- The Azure Bot resource's **calling webhook** and the app's **`Bot:BotBaseUrl`** must point at the **same** tunnel URL.

### Design choice: service-hosted media
The bot joins with **service-hosted media** (no `IMediaSession`, empty modalities). We only
need the roster — not raw audio/video — so we avoid the Windows-only local Media Platform,
TLS media certificates, and UDP media ports. This keeps the POC lightweight while still
receiving participant/roster notifications.

---

## 4. Project layout

```
TeamsMeetingBotPoc/
├─ README.md                ← you are here (onboarding + feasibility)
├─ TESTING.md               ← step-by-step test runbook
├─ AZURE_BOT_SETUP.md       ← how to create the Azure Bot resource + calling webhook
├─ NuGet.config             ← pins package sources (see "Environment gotchas")
├─ TeamsMeetingBotPoc.sln
└─ MeetingBot/
   ├─ MeetingBot.csproj     ← net8.0; Azure.Identity + Graph Communications Calls/Client 1.2.0
   ├─ Program.cs            ← minimal-API host + all HTTP endpoints
   ├─ BotOptions.cs         ← strongly-typed config (tenant, app id/secret, organizer, callback)
   ├─ MeetingScheduler.cs   ← schedules the meeting via raw Graph REST
   ├─ GraphTokenProvider.cs ← client-credential tokens for the calling SDK
   ├─ JoinUrlParser.cs      ← parses a joinWebUrl → ChatInfo + OrganizerMeetingInfo
   ├─ CallingBotService.cs  ← builds calling client, joins, holds & prints the roster
   └─ HttpTranslation.cs    ← bridges ASP.NET Core ↔ SDK notification pipeline
```

| File | Responsibility |
|---|---|
| `BotOptions.cs` | Config model. `CallbackUri = BotBaseUrl + "/api/calling"`. |
| `MeetingScheduler.cs` | `POST /users/{id}/onlineMeetings` via `ClientSecretCredential` → `ScheduledMeeting {Id, Subject, JoinWebUrl}`. |
| `GraphTokenProvider.cs` | `ITokenProvider` using `ClientSecretCredential`, scope `graph/.default`. |
| `JoinUrlParser.cs` | Extracts `threadId`, `Tid`, `Oid` from `joinWebUrl`. |
| `CallingBotService.cs` | Builds the calling client, joins meetings, and keeps a unique app-ID-to-validation-meeting correlation for up to 24 hours. |
| `HttpTranslation.cs` | Converts ASP.NET Core `HttpRequest`/`HttpResponse` ↔ `HttpRequestMessage`/`HttpResponseMessage`. |
| `Program.cs` | Minimal-API host + endpoints. |

---

## 5. HTTP endpoints

| Method + path | Purpose |
|---|---|
| `GET /` | Health probe. Returns `{ status, callback }`. |
| `POST /api/calling` | **Signaling webhook** — Microsoft Graph POSTs call/roster notifications here. Not called by hand. |
| `POST /schedule-and-join?appId=<app-id>&minutes=30&subject=Foo` | Schedule a new meeting, have the bot join it, and store the app ID correlation. Returns `{ joinWebUrl }`. |
| `POST /join-status` | Check the stored meeting for a JSON `appId`. Returns `{ isJoined, eventTime, meetingUrl }`; removes the record after Kusto confirms the join. |
| `POST /join?joinUrl=<url>` | Have the bot join an **existing** meeting by its join URL. Returns `{ callId }`. |
| `GET /calls` | List the ids of every meeting the bot is currently in. Returns `{ callIds: [...] }`. |
| `GET /participants/{callId}` | Return the live roster **for that one call** as JSON **and** print it to the bot console. `404` if the id isn't an active call. |
| `POST /leave/{callId}` | Bot leaves **that one** meeting. `404` if the id isn't an active call. |

The bot can be in **multiple meetings at once** — every call is tracked independently by its
`callId`, so roster reads and leaves always target exactly one meeting. Ended calls are
auto-removed from the active set.

The app listens on **http://localhost:5275** in dev (see `Properties/launchSettings.json`).

---

## 6. Prerequisites

**Tooling**
- .NET 8 SDK — `dotnet --version` should print `8.x`.
- [`devtunnel`](https://learn.microsoft.com/azure/developer/dev-tunnels/get-started) CLI (for the public HTTPS callback).
- `curl.exe` (ships with Windows) or any HTTP client.

**Azure / tenant setup (one-time, already done for this POC)**
1. **App registration** with **admin-consented Graph application permissions**:
   - `OnlineMeetings.ReadWrite.All` — create the meeting
   - `Calls.JoinGroupCall.All` — join the meeting
   - `Calls.JoinGroupCallAsGuest.All` — (optional) guest join
   - `Calls.AccessMedia.All` — *only* if switching to application-hosted media (not used here)
2. **Application access policy** (Teams PowerShell `New-CsApplicationAccessPolicy` /
   `Grant-CsApplicationAccessPolicy`) so the app may create `onlineMeetings` on behalf of the organizer.
3. **Azure Bot resource** reusing that app registration, with the **Teams channel + calling
   webhook** enabled → see **`AZURE_BOT_SETUP.md`**.

---

## 7. Configuration (secrets stay OUT of source)

Secrets live in **.NET user-secrets**, never in `appsettings.json`. From the `MeetingBot` folder:

```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot

dotnet user-secrets set "Bot:TenantId"        "<tenant-guid>"
dotnet user-secrets set "Bot:AppId"           "<app-guid>"
dotnet user-secrets set "Bot:AppSecret"       "<secret>"
dotnet user-secrets set "Bot:OrganizerUserId" "<organizer-object-id-GUID>"   # NOT the UPN
dotnet user-secrets set "Bot:BotBaseUrl"      "https://<your-tunnel>-5275.<region>.devtunnels.ms"
```

> ⚠️ **`Bot:OrganizerUserId` must be the object-id GUID**, not the UPN — Graph's
> `onlineMeetings` URL rejects the UPN with a "not a valid GUID" error.

---

## 8. Run & test — quick start

Do these in **three terminals**.

### Terminal 1 — public tunnel (leave open)
```powershell
devtunnel user login                       # once, interactive browser sign-in
devtunnel host -p 5275 --allow-anonymous   # copy the printed https://...-5275.<region>.devtunnels.ms URL
```
`--allow-anonymous` is required so Microsoft Graph can POST notifications without auth.

### One-time — point config + Azure Bot at the tunnel
```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot
dotnet user-secrets set "Bot:BotBaseUrl" "https://<id>-5275.<region>.devtunnels.ms"
```
Then in the Azure Portal set the Bot's **Calling webhook** to
`https://<id>-5275.<region>.devtunnels.ms/api/calling` (see `AZURE_BOT_SETUP.md`).

> 🔁 **Every time the tunnel URL changes, update BOTH** the `Bot:BotBaseUrl` secret **and**
> the Azure Bot calling webhook to match — otherwise notifications never arrive.

### Terminal 2 — run the bot (leave open; roster changes log here)
```powershell
cd C:\Users\libwei\source\TeamsMeetingBotPoc\MeetingBot
dotnet run
```

### Terminal 3 — drive the demo
```powershell
# 1) schedule a meeting AND have the bot join it:
curl.exe -s -X POST "http://localhost:5275/schedule-and-join?appId=<app-id>&minutes=30&subject=RosterTest"

# 2) a REAL user (the organizer) opens the returned joinWebUrl and joins,
#    so the meeting session actually starts.

# 3) list the live participants FOR THAT CALL (also prints to Terminal 2):
curl.exe -s "http://localhost:5275/participants/<callId>"

# (optional) see every meeting the bot is currently in:
curl.exe -s "http://localhost:5275/calls"

# 4) when done, bot leaves THAT meeting:
curl.exe -s -X POST "http://localhost:5275/leave/<callId>"
```

**Expected result:** after a human joins, `GET /participants/<callId>` returns two entries — the bot
(`kind: application/bot`) and the human (`kind: user`, with mute state) — proving the full flow.
Because each call is tracked separately, you can repeat step 1 for a second meeting and read its
roster with that meeting's own `callId` — the two rosters never mix.

For the detailed runbook and troubleshooting table, see **`TESTING.md`**.

---

## 9. Environment gotchas (save yourself time)

- **NuGet 401 / nuget.org disabled** — the parent machine config disables nuget.org and points
  at a private feed. The local `NuGet.config` fixes this with `<clear/>` + a disabled-sources
  `<clear/>`. Don't delete it.
- **Package versions are pinned** — the Communications SDK `1.2.0.17950` brings
  `Microsoft.Graph 5.92.0` (models under `Microsoft.Graph.Models`). **Do NOT add
  `Microsoft.Graph 6.x`** — it causes model conflicts. Meeting scheduling deliberately uses raw
  REST to stay version-independent.
- **`/participants/{callId}` stays `[]`** — no human has joined yet (the meeting session hasn't
  started), or notifications aren't reaching `/api/calling` (wrong webhook / tunnel not
  `--allow-anonymous` / `BotBaseUrl` mismatch). A `404` instead means the `callId` isn't active —
  call `GET /calls` to see valid ids.
- **"Policy hook failed" log lines** sometimes appear even though the operation succeeded — check
  the actual HTTP response before assuming failure.
- **Security** — never commit secrets. If a client secret is ever shared in chat/logs, **rotate it**.

---

## 10. References

- Bots for Teams calls & online meetings — https://learn.microsoft.com/microsoftteams/platform/bots/calls-and-meetings/calls-meetings-bots-overview
- Graph Communications Calling SDK — https://microsoftgraph.github.io/microsoft-graph-comms-samples/docs/
- Sample: `OfficeDev/Microsoft-Teams-Samples/samples/bot-calling-meeting/csharp`
- Create onlineMeeting — https://learn.microsoft.com/graph/api/application-post-onlinemeetings
- Dev tunnels — https://learn.microsoft.com/azure/developer/dev-tunnels/get-started
