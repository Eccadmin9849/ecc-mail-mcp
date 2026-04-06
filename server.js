// ECC Exteriors — Custom Outlook Send MCP Server
// Deploy to Vercel, Railway, or any Node.js host
// Required env vars: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SENDER_EMAIL

// ECC Exteriors — Custom Outlook Send MCP Server
// File: server.js
// Deploy to Vercel or any Node.js host

import express from "express";

const app = express();
app.use(express.json());

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  SENDER_EMAIL, // e.g. rich@eccexteriors.com
  PORT = 3000,
} = process.env;

// --- Get Microsoft Graph access token ---
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });
  const res = await fetch(url, { method: "POST", body });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Token error: ${JSON.stringify(data)}`);
  return data.access_token;
}

// --- Send email via Graph API ---
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
    `https://graph.microsoft.com/v1.0/users/${SENDER_EMAIL}/sendMail`,
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

// --- MCP Endpoint ---
app.post("/mcp", async (req, res) => {
  const { method, params } = req.body;

  // MCP tool discovery
  if (method === "tools/list") {
    return res.json({
      tools: [
        {
          name: "send_email",
          description: "Send an email via Outlook on behalf of ECC Exteriors. Use for sending reports, alerts, and summaries to team members.",
          inputSchema: {
            type: "object",
            required: ["to", "subject", "body"],
            properties: {
              to: {
                oneOf: [
                  { type: "string", description: "Single recipient email address" },
                  { type: "array", items: { type: "string" }, description: "Multiple recipient email addresses" }
                ]
              },
              subject: { type: "string", description: "Email subject line" },
              body: { type: "string", description: "Email body — HTML is supported for formatted reports" },
              cc: {
                oneOf: [
                  { type: "string", description: "Single CC email address" },
                  { type: "array", items: { type: "string" }, description: "Multiple CC email addresses" }
                ]
              }
            }
          }
        }
      ]
    });
  }

  // MCP tool execution
  if (method === "tools/call") {
    const { name, arguments: args } = params;
    if (name === "send_email") {
      try {
        const result = await sendEmail(args);
        return res.json({ content: [{ type: "text", text: JSON.stringify(result) }] });
      } catch (err) {
        return res.status(500).json({ error: err.message });
      }
    }
  }

  res.status(404).json({ error: "Unknown method" });
});

// Health check
app.get("/", (req, res) => res.send("ECC Mail MCP Server running"));

app.listen(PORT, () => console.log(`ECC Mail MCP running on port ${PORT}`));
