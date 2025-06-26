// 📁 src/OneDriveManager.js
import React, { useEffect, useState } from "react";
import axios from "axios";
import { app, authentication } from "@microsoft/teams-js";

const OneDriveManager = () => {
  const [token, setToken] = useState(null);
  const [files, setFiles] = useState([]);
  const [error, setError] = useState("");
  const [debugLog, setDebugLog] = useState(["🚀 Initializing..."]);

  const addLog = (msg) => setDebugLog(logs => [...logs, msg]);

  useEffect(() => {
    app.initialize().then(() => {
      addLog("✅ Teams SDK initialized. Requesting SSO token...");

      authentication.getAuthToken({
        resources: ["https://graph.microsoft.com"],
        successCallback: (token) => {
          console.log("✅ SSO Token:", token);
          addLog("✅ SSO token received. Fetching OneDrive files...");
          setToken(token);
        },
        failureCallback: (err) => {
          console.error("❌ Teams SSO failed:", err);
          addLog("❌ Teams SSO failed: " + err);
          setError("Teams SSO failed: " + err);
        }
      });
    }).catch(err => {
      console.error("❌ Teams SDK init failed:", err);
      addLog("❌ Teams SDK init failed: " + err.message);
      setError("Teams SDK init failed: " + err.message);
    });
  }, []);

  useEffect(() => {
    if (!token) return;

    addLog("📂 Calling Graph API for OneDrive files...");
    axios.get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${token}` }
    }).then(res => {
      setFiles(res.data.value);
      addLog("📂 OneDrive files fetched.");
    }).catch(err => {
      console.error("❌ Graph API failed:", err);
      addLog("❌ Graph API failed: " + err.message);
      setError("Failed to fetch files: " + err.message);
    });
  }, [token]);

  const convertToPdf = async (itemId, name) => {
    addLog(`🌀 Converting '${name}' to PDF...`);
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

      addLog("✅ Download triggered.");
    } catch (err) {
      console.error("❌ Convert failed:", err);
      addLog("❌ Convert failed: " + err.message);
      setError("Convert failed: " + err.message);
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>OneDrive File Viewer (Teams SSO)</h2>

      <div style={{ fontSize: 13, color: "#888", marginBottom: 10 }}>
        <strong>Debug Log</strong>
        <ul style={{ paddingLeft: 20 }}>
          {debugLog.map((line, i) => <li key={i}>{line}</li>)}
        </ul>
      </div>

      {error && <div style={{ color: "red" }}>{error}</div>}

      {!token && <p>🔐 Authenticating via Teams SSO...</p>}
      {token && files.length === 0 && <p>📂 Loading files...</p>}

      <ul>
        {files.map(file => (
          <li key={file.id}>
            {file.name}
            {file.name.endsWith(".docx") && (
              <button onClick={() => convertToPdf(file.id, file.name)} style={{ marginLeft: 10 }}>
                Convert to PDF
              </button>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
};

export default OneDriveManager;
