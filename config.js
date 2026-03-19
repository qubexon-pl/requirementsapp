window.APP_CONFIG = {
  msal: {
    clientId: "YOUR-ENTRA-APP-CLIENT-ID",
    tenantId: "YOUR-TENANT-ID-OR-DOMAIN",
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin
  },
  powerPlatform: {
    agentUri: "https://default1dc9b339fadb432e86df423c38a0fc.b8.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/cre2f_offerings/conversations?api-version=2022-03-01-preview",
    scopes: [
      "https://service.powerapps.com/user_impersonation"
    ]
  },
  chat: {
    locale: "en-US",
    welcomeMessage: "Welcome to the Offerings data agent."
  }
};
