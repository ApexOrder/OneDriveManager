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
  const [debug, setDebug] = useState("Initializing...");
  const [debugLogs, setDebugLogs] = useState([]);
  const authRef = useRef(false);

  const log = (msg, level = "log") => {
    if (console[level]) console[level](msg);
    setDebugLogs(prev => [...prev, msg]);
    setDebug(msg);
  };

  useEffect(() => {
    if (authRef.current) return;
    authRef.current = true;

    log("Handling redirect result...");
    msalInstance.handleRedirectPromise().then(resp => {
      log(`[MSAL] Redirect result: ${JSON.stringify(resp)}`);

      if (resp && resp.accessToken) {
        log("[Auth] Redirect token received");
        setToken(resp.accessToken);
        log("Redirect token processed");
        return;
      }

      log("[Teams] Waiting for Teams SDK to initialize...");
      app.initialize().then(() => {
        log("[Teams] SDK initialized.");
        const accounts = msalInstance.getAllAccounts();

        if (accounts.length > 0) {
          log(`[Auth] Found cached account: ${accounts[0].username}`);
          msalInstance.acquireTokenSilent({
            scopes: authConfig.scopes,
            account: accounts[0]
          }).then(resp => {
            setToken(resp.accessToken);
            log("Token acquired silently.");
          }).catch(err => {
            log(`[Auth] Silent token failed: ${err.message}`, "warn");
            log("Redirecting for login...");
            msalInstance.loginRedirect({ scopes: authConfig.scopes });
          });
        } else {
          log("[Auth] No cached account, redirecting to login...");
          msalInstance.loginRedirect({ scopes: authConfig.scopes });
        }
      }).catch(err => {
        log(`[Teams] SDK init failed: ${err.message}`, "error");
        setError("Teams SDK init failed: " + err.message);
      });
    }).catch(err => {
      log(`[MSAL] Redirect handling error: ${err.message}`, "error");
      setError("Auth error: " + err.message);
    });
  }, []);

  useEffect(() => {
    if (!token) return;

    log("[Graph] Fetching OneDrive files...");
    axios.get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${token}` }
    })
    .then(res => {
      log(`[Graph] Files retrieved: ${res.data.value.length} items`);
      setFiles(res.data.value);
    })
    .catch(err => {
      log(`[Graph] Fetch failed: ${err.message}`, "error");
      setError("Failed to fetch files: " + err.message);
    });
  }, [token]);

  const convertToPdf = async (itemId, name) => {
    log(`[Convert] Converting ${name} to PDF...`);
    try {
      const res = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content?format=pdf`, {
        responseType: "blob",
        headers: { Authorization: `Bearer ${token}` }
      });
      const url = URL.createObjectURL(res.data);
      const a = document.createElement("a");
      a.href = url;
      a.download = name.replace(/\.docx$/, ".pdf");
      a.click();
      log("[Convert] Download triggered.");
    } catch (err) {
      log(`[Convert] Failed to convert file: ${err.message}`, "error");
      setError("Convert failed: " + err.message);
    }
  };

  return (
    <div style={{ padding: 20, color: "#fff", fontFamily: "Arial" }}>
      <h2 style={{ color: "yellow" }}>My OneDrive Files</h2>
      <div style={{ fontSize: 13, color: "#aaa", marginBottom: 10 }}>
        Debug: {debug}
      </div>

      {error && <div style={{ color: "red", marginBottom: 10 }}>{error}</div>}
      {!token && <p>Authenticating with Microsoft...</p>}
      {token && files.length === 0 && <p>Loading files from OneDrive...</p>}

      <ul>
        {files.map(file => (
          <li key={file.id}>
            {file.name}
            {file.name.endsWith(".docx") && (
              <button onClick={() => convertToPdf(file.id, file.name)}>Convert to PDF</button>
            )}
          </li>
        ))}
      </ul>

      <div style={{ marginTop: 30, padding: 10, background: "#222", borderRadius: 5 }}>
        <strong>Debug Log</strong>
        <pre style={{ fontSize: 12, maxHeight: 200, overflowY: "auto" }}>
{debugLogs.map((msg, i) => `• ${msg}\n`)}
        </pre>
      </div>
    </div>
  );
};

export default OneDriveManager;
