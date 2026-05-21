import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import cookieParser from "cookie-parser";
import axios from "axios";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3000);
const APP_URL = (process.env.APP_URL || `http://localhost:${PORT}`).replace(/\/$/, "");

app.use(express.json({ limit: "30mb" }));
app.use(cookieParser());

// ServiceM8 OAuth Config
const CLIENT_ID = process.env.SERVICEM8_CLIENT_ID;
const CLIENT_SECRET = process.env.SERVICEM8_CLIENT_SECRET;
const REDIRECT_URI = `${APP_URL}/auth/callback`;

const cleanServiceM8Label = (label: string) =>
  String(label || "Field")
    .replace(/[^a-zA-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const trimBadgeName = (value: string, fallback: string) =>
  cleanServiceM8Label(value || fallback || "Form").substring(0, 11);

// API routes
app.get("/api/auth/url", (req, res) => {
  if (!CLIENT_ID) {
    return res.status(500).json({ error: "SERVICEM8_CLIENT_ID not configured" });
  }

  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    redirect_uri: REDIRECT_URI,
    response_type: "code",
    scope: "manage_forms read_forms",
  });

  const authUrl = `https://api.servicem8.com/oauth/authorize?${params.toString()}`;
  res.json({ url: authUrl });
});

app.get("/auth/callback", async (req, res) => {
  const { code } = req.query;

  if (!code) {
    return res.status(400).send("No code provided");
  }

  try {
    const response = await axios.post("https://api.servicem8.com/oauth/access_token", new URLSearchParams({
      grant_type: "authorization_code",
      code: code as string,
      client_id: CLIENT_ID!,
      client_secret: CLIENT_SECRET!,
      redirect_uri: REDIRECT_URI,
    }).toString(), {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });

    const { access_token, refresh_token } = response.data;

    // Set cookies
    const secureCookie = APP_URL.startsWith("https://");
    res.cookie("sm8_access_token", access_token, {
      secure: secureCookie,
      sameSite: secureCookie ? "none" : "lax",
      httpOnly: true,
      maxAge: 3600 * 1000, // 1 hour
    });

    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              window.close();
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. This window should close automatically.</p>
        </body>
      </html>
    `);
  } catch (error: any) {
    console.error("OAuth Error:", error.response?.data || error.message);
    res.status(500).send("Authentication failed");
  }
});

app.get("/api/auth/status", (req, res) => {
  const token = req.cookies.sm8_access_token;
  res.json({ authenticated: !!token });
});

app.post("/api/generate-form", (_req, res) => {
  res.status(501).json({
    error: "AI generation is handled by Codex in the chat workflow, not by the local web app. For local testing, edit a spec JSON and run npm run sm8f:sample or node scripts/generate-sm8f.mjs <spec.json> <output.sm8f>.",
  });
});

app.post("/api/servicem8/push-form", async (req, res) => {
  const { form, apiKey } = req.body;
  const oauthToken = req.cookies.sm8_access_token;

  console.log("[ServiceM8 Push] Request received");
  console.log("[ServiceM8 Push] Auth Method:", apiKey ? "API Key" : (oauthToken ? "OAuth" : "None"));

  if (!oauthToken && !apiKey) {
    console.error("[ServiceM8 Push] Error: No authentication provided");
    return res.status(401).json({ error: "Not authenticated with ServiceM8 (OAuth or API Key required)" });
  }

  if (!form) {
    console.error("[ServiceM8 Push] Error: No form data provided");
    return res.status(400).json({ error: "No form data provided" });
  }

  const headers: any = {
    "Content-Type": "application/json",
    "Accept": "application/json",
  };

  if (apiKey) {
    // Some ServiceM8 setups use X-Api-Key, others use Basic Auth
    headers["X-Api-Key"] = apiKey;
    
    // Also add Basic Auth as it's the standard for ServiceM8 API keys
    const authBuffer = Buffer.from(`${apiKey}:`).toString('base64');
    headers["Authorization"] = `Basic ${authBuffer}`;
    
    console.log("[ServiceM8 Push] Using API Key (X-Api-Key and Basic Auth)");
  } else {
    headers["Authorization"] = `Bearer ${oauthToken}`;
    console.log("[ServiceM8 Push] Using OAuth Bearer Token");
  }

  try {
    // 1. Create the Form
    console.log("[ServiceM8 Push] Creating form:", form.title);
    
    // Ensure badge_name is exactly what ServiceM8 expects (max 11 chars)
    const badgeName = trimBadgeName(form.badgeName, form.title);
    
    const formResponse = await axios.post("https://api.servicem8.com/api_1.0/form.json", {
      name: form.title,
      badge_name: badgeName,
      description: form.description,
      active: 1,
    }, { headers });

    console.log("[ServiceM8 Push] Form creation response status:", formResponse.status);
    console.log("[ServiceM8 Push] Form creation response headers:", JSON.stringify(formResponse.headers, null, 2));
    
    // Try different casing for the header
    const formUuid = formResponse.headers["x-record-uuid"] || formResponse.headers["X-Record-Uuid"];
    console.log("[ServiceM8 Push] Form created successfully. UUID:", formUuid);

    if (!formUuid) {
      throw new Error("Failed to retrieve form UUID from ServiceM8 response headers");
    }

    // 2. Create Form Fields using the correct structure
    console.log("[ServiceM8 Push] Creating", form.fields.length, "fields...");
    
    const fieldPromises = form.fields.map((field: any, index: number) => {
      // Map internal types to ServiceM8 fieldType based on the provided UI screenshot
      const typeMap: Record<string, string> = {
        text: "Text",
        number: "Number",
        date: "Date",
        checkbox: "Multiple Choice",
        select: "Multiple Choice",
        signature: "Signature",
        photo: "Photo"
      };

      const fieldData: any = {
        fieldType: typeMap[field.type] || "Text",
        additionalDetails: field.additionalDetails || "",
        mandatory: !!field.required,
        conditions: [
          { question: "", operator: "", value: "" },
          { question: "", operator: "", value: "" },
          { question: "", operator: "", value: "" }
        ],
        conditionMethod: "AND"
      };

      // For Multiple Choice, ServiceM8 uses 'choices' as an array of strings
      if (field.type === 'select' || field.type === 'checkbox') {
        fieldData.choices = field.options && field.options.length > 0 
          ? field.options 
          : ["Yes", "No"];
      }

      return axios.post("https://api.servicem8.com/api_1.0/formfield.json", {
        form_uuid: formUuid,
        name: cleanServiceM8Label(field.label),
        field_data_json: JSON.stringify(fieldData),
        sort_order: String(index + 1),
      }, { headers });
    });

    await Promise.all(fieldPromises);
    console.log("[ServiceM8 Push] All fields created successfully");

    res.json({ success: true, formUuid });
  } catch (error: any) {
    const errorData = error.response?.data;
    const errorStatus = error.response?.status;
    
    console.error("[ServiceM8 Push] API Error!");
    console.error("Status:", errorStatus);
    console.error("Data:", JSON.stringify(errorData, null, 2));
    console.error("Message:", error.message);

    res.status(500).json({ 
      error: "Failed to push form to ServiceM8", 
      details: errorData || error.message,
      status: errorStatus
    });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
