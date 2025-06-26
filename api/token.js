// 📁 /api/token.js (Vercel serverless function)
import axios from 'axios';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end('Method Not Allowed');

  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;
  const token = req.body.token;

  if (!token) return res.status(400).json({ error: 'Missing Teams token' });

  try {
    const params = new URLSearchParams();
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
    params.append('assertion', token);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('requested_token_use', 'on_behalf_of');

    const tokenRes = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      params,
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    return res.status(200).json(tokenRes.data);
  } catch (err) {
    console.error('[Token Exchange] Error:', err.response?.data || err.message);
    return res.status(500).json({ error: 'Token exchange failed', details: err.response?.data || err.message });
  }
}
