// 📁 OneDriveManager.js
import React, { useEffect, useState } from "react";
import axios from "axios";
import { app as teamsApp, authentication } from "@microsoft/teams-js";

const OneDriveManager = () => {
  const [token, setToken] = useState(null);
  const [files, setFiles] = useState([]);
  const [debug, setDebug] = useState("Initializing...");
  const [error, setError] = useState("");
  const [debugLogs, setDebugLogs] = useState([]);

  const log = (msg, level = "log") => {
    if (console[level]) console[level](msg);
    setDebugLogs((prev) => [...prev, msg]);
    setDebug(msg);
  };

  useEffect(() => {
    log("Initializing Teams SDK...");
    teamsApp.initialize();

    authentication.getAuthToken({
      successCallback: async (teamsToken) => {
        log("✅ Received Teams token. Exchanging for Graph token...");

        try {
          const res = await axios.post("/api/token", { token: teamsToken });
          const graphToken = res.data.access_token;
          setToken(graphToken);
          log("✅ Graph token acquired. Fetching OneDrive files...");

          const filesRes = await axios.get(
            "https://graph.microsoft.com/v1.0/me/drive/root/children",
            {
              headers: { Authorization: `Bearer ${graphToken}` },
            }
          );

          setFiles(filesRes.data.value);
          log(`📁 Loaded ${filesRes.data.value.length} file(s)`);
        } catch (err) {
          log("❌ Failed to exchange token or load files: " + err.message, "error");
          setError("Token exchange or OneDrive error: " + err.message);
        }
      },
      failureCallback: (err) => {
        log("❌ Teams SSO failed: " + err, "error");
        setError("Teams SSO failed: " + err);
      },
    });
  }, []);

  return (
    <div style={{ padding: 20, fontFamily: "monospace", color: "#fff", background: "#121212" }}>
      <h2 style={{ color: "#ffcc00" }}>OneDrive File Viewer (Teams SSO)</h2>
      <div style={{ fontSize: 13, color: "#aaa", marginBottom: 10 }}>Debug: {debug}</div>
      {error && <div style={{ color: "red" }}>{error}</div>}
      {!token && <p>🔐 Authenticating via Teams SSO...</p>}
      {token && files.length === 0 && <p>📦 Loading OneDrive files...</p>}

      <ul>
        {files.map((file) => (
          <li key={file.id}>{file.name}</li>
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
