// 📁 src/OneDriveManager.jsx
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
  const authRef = useRef(false);

  useEffect(() => {
    if (authRef.current) return;
    authRef.current = true;

    setDebug("Handling redirect result...");
    msalInstance.handleRedirectPromise().then(resp => {
      console.log("[MSAL] Redirect result:", resp);

      if (resp && resp.accessToken) {
        console.log("[Auth] Redirect token received");
        setToken(resp.accessToken);
        setDebug("Redirect token processed");
        return;
      }

      setDebug("Initializing Teams SDK...");
      app.initialize().then(() => {
        console.log("[Teams] SDK initialized.");
        setDebug("Teams SDK initialized. Checking for accounts...");

        const accounts = msalInstance.getAllAccounts();

        if (accounts.length > 0) {
          console.log("[Auth] Found cached account:", accounts[0]);
          setDebug("Cached account found. Getting token silently...");

          msalInstance.acquireTokenSilent({
            scopes: authConfig.scopes,
            account: accounts[0]
          }).then(resp => {
            setToken(resp.accessToken);
            setDebug("Token acquired silently.");
          }).catch(err => {
            console.warn("[Auth] Silent token failed, redirecting...");
            setDebug("Silent token failed. Redirecting to login...");
            msalInstance.loginRedirect({ scopes: authConfig.scopes });
          });
        } else {
          console.log("[Auth] No cached account, redirecting to login...");
          setDebug("No account. Redirecting to login...");
          msalInstance.loginRedirect({ scopes: authConfig.scopes });
        }
      }).catch(err => {
        console.error("[Teams] SDK init failed:", err);
        setError("Teams SDK init failed: " + err.message);
        setDebug("Teams SDK init failed");
      });
    }).catch(err => {
      console.error("[MSAL] Redirect handling error:", err);
      setError("Auth error: " + err.message);
      setDebug("Redirect handling failed");
    });
  }, []);

  useEffect(() => {
    if (!token) return;

    console.log("[Graph] Fetching OneDrive files...");
    setDebug("Fetching files from OneDrive...");

    axios.get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${token}` }
    })
    .then(res => {
      console.log("[Graph] Files retrieved:", res.data.value);
      setFiles(res.data.value);
      setDebug("Files loaded.");
    })
    .catch(err => {
      console.error("[Graph] Failed to fetch files:", err);
      setError("Failed to fetch files: " + err.message);
      setDebug("Failed to fetch files");
    });
  }, [token]);

  const convertToPdf = async (itemId, name) => {
    console.log(`[Convert] Converting ${name} to PDF...`);
    setDebug(`Converting ${name} to PDF...`);

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
      console.log("[Convert] Download triggered.");
      setDebug("Download triggered");
    } catch (err) {
      console.error("[Convert] Failed to convert file:", err);
      setError("Convert failed: " + err.message);
      setDebug("Convert failed");
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>My OneDrive Files</h2>
      <div style={{ fontSize: 13, color: "#aaa", marginBottom: 10 }}>Debug: {debug}</div>

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
    </div>
  );
};

export default OneDriveManager;
