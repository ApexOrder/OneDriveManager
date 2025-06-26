import React, { useEffect, useState, useRef } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import axios from "axios";
import { app } from "@microsoft/teams-js";

const authConfig = {
  clientId: process.env.REACT_APP_CLIENT_ID,
  tenantId: process.env.REACT_APP_TENANT_ID,
  authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID}`,
  redirectUri: window.location.origin + "/auth.html",
  scopes: ["User.Read", "Files.ReadWrite.All", "Sites.Read.All"]
};

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: authConfig.clientId,
    authority: authConfig.authority,
    redirectUri: authConfig.redirectUri
  }
});

const OneDriveManager = () => {
  const [token, setToken] = useState(null);
  const [files, setFiles] = useState([]);
  const [error, setError] = useState("");
  const [debugLogs, setDebugLogs] = useState([]);
  const authRef = useRef(false);

  const log = (msg, level = "log") => {
    if (console[level]) console[level](msg);
    setDebugLogs(prev => [...prev, msg]);
  };

  useEffect(() => {
    if (authRef.current) return;
    authRef.current = true;

    log("🔄 Starting auth flow...");
    msalInstance.handleRedirectPromise().then(resp => {
      if (resp && resp.accessToken) {
        log("✅ Redirect token received");
        setToken(resp.accessToken);
        return;
      }

      log("🧠 No redirect token. Checking Teams context...");

      app.initialize().then(() => {
        log("✅ Teams SDK initialized");
        app.getContext().then(() => {
          log("📦 Running inside Teams. Using popup login...");

          msalInstance.loginPopup({ scopes: authConfig.scopes })
            .then(resp => {
              if (resp.accessToken) {
                log("✅ Access token received via popup");
                setToken(resp.accessToken);
              } else {
                log("❌ No token in popup response");
              }
            })
            .catch(err => {
              log("❌ Popup login failed: " + err.message, "error");
              setError("Popup login failed: " + err.message);
            });

        }).catch(err => {
          log("⚠️ Failed to get Teams context: " + err.message);
          fallbackToWebAuth();
        });
      }).catch(err => {
        log("⚠️ Teams SDK init failed: " + err.message);
        fallbackToWebAuth();
      });
    }).catch(err => {
      log("❌ Redirect handling error: " + err.message, "error");
      setError("Auth error: " + err.message);
    });
  }, []);

  const fallbackToWebAuth = () => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      log("🔐 Found cached account: " + accounts[0].username);
      msalInstance.acquireTokenSilent({
        scopes: authConfig.scopes,
        account: accounts[0]
      }).then(resp => {
        setToken(resp.accessToken);
        log("✅ Token acquired silently");
      }).catch(err => {
        log("⚠️ Silent token failed: " + err.message);
        msalInstance.loginRedirect({ scopes: authConfig.scopes });
      });
    } else {
      log("🔁 No account found. Using loginRedirect...");
      msalInstance.loginRedirect({ scopes: authConfig.scopes });
    }
  };

  useEffect(() => {
    if (!token) return;

    log("📁 Fetching OneDrive files...");
    axios.get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${token}` }
    })
    .then(res => {
      log(`✅ ${res.data.value.length} file(s) retrieved`);
      setFiles(res.data.value);
    })
    .catch(err => {
      log("❌ OneDrive fetch error: " + err.message, "error");
      setError("Failed to fetch files: " + err.message);
    });
  }, [token]);

  return (
    <div style={{ padding: 20, fontFamily: "monospace", color: "#fff", background: "#121212" }}>
      <h2 style={{ color: "#ffcc00" }}>OneDrive File Viewer</h2>
      {error && <div style={{ color: "red" }}>{error}</div>}
      {!token && <p>🔐 Authenticating with Microsoft...</p>}
      {token && files.length === 0 && <p>📦 Loading files...</p>}
      <ul>
        {files.map(file => (
          <li key={file.id}>
            {file.name}
            {file.name.endsWith(".docx") && (
              <button onClick={() => alert("Convert to PDF coming soon.")}>Convert to PDF</button>
            )}
          </li>
        ))}
      </ul>
      <div style={{ marginTop: 30, padding: 10, background: "#222", borderRadius: 5 }}>
        <strong>Debug Log</strong>
        <pre style={{ fontSize: 12, maxHeight: 300, overflowY: "auto" }}>
{debugLogs.map((msg, i) => `• ${msg}\n`)}
        </pre>
      </div>
    </div>
  );
};

export default OneDriveManager;
