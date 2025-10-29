export const msalConfig = {
  auth: {
    clientId: "5c1e64c0-76f2-4200-8ee5-b3b3d19b53da",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  }
};

export const loginRequest = {
  scopes: [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Sites.ReadWrite.All"
  ]
};
