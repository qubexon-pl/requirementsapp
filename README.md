# Azure Entra ID Copilot Studio Chat

This repository now contains a deployable static web application that authenticates users with Microsoft Entra ID and calls your authenticated Copilot Studio / Power Platform data agent at:

- `https://default1dc9b339fadb432e86df423c38a0fc.b8.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/cre2f_offerings/conversations?api-version=2022-03-01-preview`

## What is included

- `index.html` — the UI shell.
- `styles.css` — responsive styling for the Azure-ready chat app.
- `app.js` — MSAL authentication flow, conversation creation, and agent messaging.
- `config.sample.js` — template for Entra ID and Power Platform settings.
- `config.js` — active runtime configuration file copied from the sample.

## Configure Microsoft Entra ID

1. In the Azure portal, create or reuse an **App registration**.
2. Add a **Single-page application** redirect URI:
   - Local development: `http://localhost:8000`
   - Azure deployment: your final HTTPS site URL.
3. Enable **Access tokens** for the SPA platform if your tenant configuration requires it.
4. Grant delegated API permissions for the Power Platform resource your tenant requires.
   - The sample is preconfigured with `https://service.powerapps.com/user_impersonation`.
   - If your environment requires a different scope, update `config.js`.
5. Copy `config.sample.js` to `config.js` and fill in:
   - `clientId`
   - `tenantId`
   - optional redirect overrides

## Run locally

Because this is a static app, any small HTTP server works.

```bash
python3 -m http.server 8000
```

Then open `http://localhost:8000`.

## Deploy to Azure

### Option 1: Azure Static Web Apps

1. Push this repository to GitHub.
2. Create a **Static Web App** in Azure.
3. Set the app location to `/` and leave the build location empty.
4. After deployment, add the Static Web App URL as a redirect URI in your Entra app registration.
5. Update `config.js` so `redirectUri` matches the deployed site origin.

### Option 2: Azure App Service

1. Create a Linux or Windows Web App.
2. Deploy the repository contents as static files.
3. Configure the default document to `index.html`.
4. Add the final site URL to the SPA redirect URIs in your Entra app registration.

## Notes about the agent call

The app:

1. Signs the user in with MSAL.
2. Acquires a delegated access token for the configured scopes.
3. Creates a conversation by POSTing to the provided `/conversations` endpoint.
4. Sends messages to `/{conversationId}/activities?api-version=2022-03-01-preview`.
5. Shows any text responses returned immediately by the API and logs raw payloads in the Diagnostics panel.

If your agent returns a different payload shape, you can adjust the parsing logic in `extractReplies()` inside `app.js`.
