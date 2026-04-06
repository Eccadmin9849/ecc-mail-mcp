// ECC Exteriors — Outlook Send MCP Server
// File location: api/mcp.js
// Vercel serverless function — no Express needed

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });
  const res = await fetch(url, { method: "POST", body });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Token error: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function sendEmail({ to, subject, body, cc }) {
  const token = await getAccessToken();
  const toRecipients = (Array.isArray(to) ? to : [to]).map(addr => ({
    emailAddress: { address: addr }
  }));
  const ccRecipients = cc
    ? (Array.isArray(cc) ? cc : [cc]).map(addr => ({ emailAddress: { address: addr } }))
    : [];

  const message = {
    subject,
    body: { contentType: "HTML", content: body },
    toRecipients,
    ...(ccRecipients.length > 0 && { ccRecipients }),
  };

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${process.env.SENDER_EMAIL}/sendMail`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ message }),
    }
  );

  if (res.status === 202) return { success: true, message: "Email sent successfully" };
  const err = await res.json().catch(() => ({}));
  throw new Error(`Send failed: ${JSON.stringify(err)}`);
}

export default async function handler(req, res) {
  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();

  // Health check
  if (req.method === "GET") {
    return res.status(200).json({ status: "ECC Mail MCP Server running", version: "1.0.0" });
  }

  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { method, params, jsonrpc, id } = req.body;

  if (method === "initialize") {
    return res.status(200).json({
      jsonrpc: "2.0", id,
      result: {
        protocolVersion: "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: { name: "ecc-mail-sender", version: "1.0.0" }
      }
    });
  }

  if (method === "notifications/initialized") {
    return res.status(200).json({ jsonrpc: "2.0" });
  }

  if (method === "tools/list") {
    return res.status(200).json({
      jsonrpc: "2.0", id,
      result: {
        tools: [{
          name: "send_email",
          description: "Send an email via Outlook on behalf of ECC Exteriors. Use this to send reports, alerts, and summaries to team members.",
          inputSchema: {
            type: "object",
            required: ["to", "subject", "body"],
            properties: {
              to: { type: "string", description: "Recipient email address" },
              subject: { type: "string", description: "Email subject line" },
              body: { type: "string", description: "Email body — HTML supported for formatted reports" },
              cc: { type: "string", description: "CC email address (optional)" }
            }
          }
        }]
      }
    });
  }

  if (method === "tools/call") {
    const { name, arguments: args } = params;
    if (name === "send_email") {
      try {
        const result = await sendEmail(args);
        return res.status(200).json({
          jsonrpc: "2.0", id,
          result: { content: [{ type: "text", text: JSON.stringify(result) }] }
        });
      } catch (err) {
        return res.status(200).json({
          jsonrpc: "2.0", id,
          error: { code: -32000, message: err.message }
        });
      }
    }
  }

  return res.status(200).json({
    jsonrpc: "2.0", id,
    error: { code: -32601, message: "Method not found" }
  });
}
