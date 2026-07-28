# Creating the Azure Bot resource (for the calling POC)

You already have: app registration (App ID `<APP_ID_GUID>`),
client secret, and consented Graph permissions. You now need an **Azure Bot resource**
that reuses that app registration, with the **Teams channel + calling webhook**.

Prereq: an **Azure subscription in the `<TENANT_NAME>` tenant** (the Azure Bot's Microsoft
App ID must match your app registration, which lives in that tenant).

---

## Part A — Create the Azure Bot resource (Portal)

1. Portal → **Create a resource** → search **Azure Bot** → **Create**.
2. **Bot handle**: e.g. `meeting-bot-poc` (globally-ish unique display name).
3. **Subscription** / **Resource group**: pick one in the `<TENANT_NAME>` tenant.
4. **Pricing tier**: **F0 (Free)** is fine for the POC.
5. **Microsoft App ID** section:
   - **Type of App**: **Single Tenant** (multi-tenant is deprecated for new bots).
   - **Creation type**: **Use existing app registration**.
   - **App ID**: `<APP_ID_GUID>`
   - **App tenant ID**: `<TENANT_ID_GUID>`
6. **Review + create** → **Create**.

> If the portal blocks "use existing app registration", use the CLI in Part D.

---

## Part B — Add the Teams channel + enable calling

1. Open the new Azure Bot resource → left nav **Channels**.
2. Click **Microsoft Teams** → accept terms.
3. **Messaging** tab: choose **Microsoft Teams** (commercial) cloud.
4. Switch to the **Calling** tab:
   - Check **Enable calling**.
   - **Webhook (for calling)** = `https://<your-tunnel>-5275.<region>.devtunnels.ms/api/calling`
     (the same tunnel URL you'll set as `Bot:BotBaseUrl`; note the **`/api/calling`** path).
5. **Apply / Save**.

> Every time your devtunnel URL changes, update this webhook to match.

---

## Part C — Set the messaging endpoint (Configuration) — OPTIONAL

> Not needed for this calling-only POC (the bot never uses `/api/messages`).
> Only do this if the portal refuses to save with an empty Messaging endpoint —
> then a placeholder tunnel URL is enough; nothing will actually call it.

1. Azure Bot → **Configuration**.
2. **Messaging endpoint**: `https://<your-tunnel>-5275.<region>.devtunnels.ms/api/messages`
   (our POC is calling-only and doesn't implement `/api/messages`, but the field is
   required — a placeholder on the tunnel host is fine).
3. Confirm **Microsoft App ID** = `<APP_ID_GUID>`. **Save**.

---

## Part D — CLI alternative (if the portal path is blocked)

```powershell
# sign in to the tenant that owns the app registration:
az login --tenant <TENANT_DOMAIN>
az account set --subscription "<subscription-id-in-<TENANT_NAME>>"

az bot create `
  --resource-group "<rg>" `
  --name "meeting-bot-poc" `
  --app-type SingleTenant `
  --appid "<APP_ID_GUID>" `
  --tenant-id "<TENANT_ID_GUID>" `
  --sku F0
```
Then enable the Teams channel + calling webhook in the portal (Part B) — the calling
webhook isn't exposed through `az bot` cleanly, so use the portal for that part.

---

## How this ties into the POC

```
Azure Bot "Calling webhook"  ─┐
                              ├─►  {tunnel}/api/calling   (Program.cs MapPost)
Bot:BotBaseUrl (user-secret) ─┘        │
                                       └► ProcessNotificationAsync ► ICall.Participants
```
The Azure Bot's calling webhook and the POC's `Bot:BotBaseUrl` must point at the **same
tunnel URL**. After this is set, follow `TESTING.md` Part B to run the join + roster test.

---

## Common gotchas
| Problem | Fix |
|---|---|
| Can't "use existing app registration" in portal | Use `az bot create` (Part D). |
| Calling tab missing | Finish the Teams channel **Messaging** tab first, then reopen Channels. |
| Bot joins but roster empty | Calling webhook URL wrong or not `/api/calling`; tunnel not `--allow-anonymous`. |
| No Azure subscription in <TENANT_NAME> | You need one there (or an owner to add you) to host the Azure Bot resource. |
