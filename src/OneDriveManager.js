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
  const authRef = useRef(false);

  useEffect(() => {
    if (authRef.current) return;
    authRef.current = true;

    console.log("[Teams] Waiting for Teams SDK to initialize...");
    app.initialize().then(() => {
      console.log("[Teams] Teams SDK initialized.");

      const accounts = msalInstance.getAllAccounts();

      if (accounts.length > 0) {
        console.log("[Auth] Using cached account:", accounts[0]);
        msalInstance.acquireTokenSilent({
          scopes: authConfig.scopes,
          account: accounts[0]
        }).then(resp => {
          setToken(resp.accessToken);
        }).catch(err => {
          console.warn("[Auth] Silent token failed. Falling back to popup...");
          msalInstance.loginPopup({ scopes: authConfig.scopes }).then(resp => {
            setToken(resp.accessToken);
          }).catch(loginErr => {
            console.error("[Auth] Popup failed:", loginErr);
            setError("Login failed: " + loginErr.message);
          });
        });
      } else {
        console.log("[Auth] No cached account, using loginPopup...");
        msalInstance.loginPopup({ scopes: authConfig.scopes }).then(resp => {
          setToken(resp.accessToken);
        }).catch(err => {
          console.error("[Auth] Login failed:", err);
          setError("Login failed: " + err.message);
        });
      }
    }).catch(err => {
      console.error("[Teams] Failed to initialize Teams SDK:", err);
      setError("Teams SDK init failed: " + err.message);
    });
  }, []);

  useEffect(() => {
    if (!token) return;

    console.log("[Graph] Fetching OneDrive files...");
    axios.get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${token}` }
    })
    .then(res => {
      console.log("[Graph] Files retrieved:", res.data.value);
      setFiles(res.data.value);
    })
    .catch(err => {
      console.error("[Graph] Failed to fetch files:", err);
      setError("Failed to fetch files: " + err.message);
    });
  }, [token]);

  const convertToPdf = async (itemId, name) => {
    console.log(`[Convert] Converting ${name} to PDF...`);
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
    } catch (err) {
      console.error("[Convert] Failed to convert file:", err);
      setError("Convert failed: " + err.message);
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>My OneDrive Files</h2>

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
