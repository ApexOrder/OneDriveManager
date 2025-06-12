import React, { useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import axios from "axios";

const authConfig = {
  clientId: process.env.CLIENT_ID,
  tenantId: process.env.TENANT_ID,
  authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
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

  useEffect(() => {
    msalInstance.loginPopup({ scopes: authConfig.scopes }).then(resp => {
      setToken(resp.accessToken);
    });
  }, []);

  useEffect(() => {
    if (!token) return;
    axios.get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${token}` }
    }).then(res => {
      setFiles(res.data.value);
    });
  }, [token]);

  const convertToPdf = async (itemId, name) => {
    const res = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content?format=pdf`, {
      responseType: "blob",
      headers: { Authorization: `Bearer ${token}` }
    });
    const url = URL.createObjectURL(res.data);
    const a = document.createElement("a");
    a.href = url;
    a.download = name.replace(/\.docx$/, ".pdf");
    a.click();
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>My OneDrive Files</h2>
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
