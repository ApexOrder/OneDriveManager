import { PublicClientApplication } from "@azure/msal-browser";
import * as microsoftTeams from "@microsoft/teams-js";

const CLIENT_ID = process.env.REACT_APP_CLIENT_ID;
const TENANT_ID = process.env.REACT_APP_TENANT_ID;
const SCOPES = [process.env.NEXT_PUBLIC_SCOPES];

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
});

export async function authenticateWithGraph() {
  await microsoftTeams.app.initialize();
  console.log("✅ Teams SDK initialized");

  return new Promise((resolve, reject) => {
    microsoftTeams.authentication.getAuthToken({
      successCallback: async (ssoToken) => {
        console.log("✅ SSO token received");

        try {
          const account = msalInstance.getAllAccounts()[0];
          if (!account) {
            console.warn("🟡 No account found in MSAL");
            await msalInstance.ssoSilent({ scopes: SCOPES });
          }

          const response = await msalInstance.acquireTokenSilent({
            scopes: SCOPES,
            account: msalInstance.getAllAccounts()[0]
          });

          resolve(response.accessToken);
        } catch (err) {
          console.error("❌ MSAL token exchange failed:", err);
          reject(err);
        }
      },
      failureCallback: (err) => {
        console.error("❌ Teams SSO failed:", err);
        reject(err);
      }
    });
  });
}
