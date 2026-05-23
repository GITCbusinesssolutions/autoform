import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import cookieParser from "cookie-parser";
import axios from "axios";
import dotenv from "dotenv";
import fs from "node:fs/promises";
import os from "node:os";
import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3000);
const APP_URL = (process.env.APP_URL || `http://localhost:${PORT}`).replace(/\/$/, "");
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5.4-mini";
const CODEX_BRIDGE_DIR = process.env.CODEX_BRIDGE_DIR || (process.platform === "win32" ? "C:\\tmp\\autoform-codex-bridge" : path.join(os.tmpdir(), "autoform-codex-bridge"));
const LOCAL_DB_PATH = process.env.LOCAL_DB_PATH || path.join(process.cwd(), "data", "local-db.json");
const ENV_PATH = path.join(process.cwd(), ".env");

app.use(express.json({ limit: "30mb" }));
app.use(cookieParser());
app.get("/favicon.ico", (_req, res) => res.status(204).end());

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

app.post("/api/auth/check", (req, res) => {
  const configured = process.env.APP_PASSWORD;
  if (!configured) return res.json({ ok: true });
  res.json({ ok: req.body?.password === configured });
});

async function readLocalDb() {
  try {
    return JSON.parse(await fs.readFile(LOCAL_DB_PATH, "utf8"));
  } catch {
    return { projects: [] };
  }
}

async function writeLocalDb(data: any) {
  await fs.mkdir(path.dirname(LOCAL_DB_PATH), { recursive: true });
  await fs.writeFile(LOCAL_DB_PATH, JSON.stringify(data, null, 2), "utf8");
}

async function updateEnvValue(key: string, value: string) {
  let raw = "";
  try {
    raw = await fs.readFile(ENV_PATH, "utf8");
  } catch {
    raw = "";
  }

  const lines = raw.split(/\r?\n/).filter((line, index, all) => line.length || index < all.length - 1);
  const nextLine = `${key}=${value}`;
  const index = lines.findIndex((line) => line.startsWith(`${key}=`));

  if (index >= 0) lines[index] = nextLine;
  else lines.push(nextLine);

  await fs.writeFile(ENV_PATH, `${lines.join("\n")}\n`, "utf8");
}

app.get("/api/projects", async (_req, res) => {
  res.json(await readLocalDb());
});

app.put("/api/projects", async (req, res) => {
  const projects = Array.isArray(req.body?.projects) ? req.body.projects : [];
  await writeLocalDb({ projects, updatedAt: new Date().toISOString() });
  res.json({ ok: true });
});

app.post("/api/settings/openai-key", async (req, res) => {
  const apiKey = String(req.body?.apiKey || "").trim();
  if (apiKey) {
    process.env.OPENAI_API_KEY = apiKey;
    await updateEnvValue("OPENAI_API_KEY", apiKey);
  }
  res.json({
    ok: !!apiKey,
    message: apiKey
      ? "OpenAI API key saved to local .env and active for this server session. Use Vercel environment variables when deployed."
      : "Enter an API key to enable OpenAI generation.",
    received: !!apiKey,
  });
});

app.post("/api/generate-form", (_req, res) => {
  res.status(501).json({
    error: "AI generation is handled by Codex in the chat workflow, not by the local web app. For local testing, edit a spec JSON and run npm run sm8f:sample or node scripts/generate-sm8f.mjs <spec.json> <output.sm8f>.",
  });
});

const formPlanSchema = {
  name: "autoform_plan",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    required: ["summary", "assumptions", "clarificationQuestions", "designNotes", "nextSteps", "spec"],
    properties: {
      summary: { type: "string" },
      assumptions: { type: "array", items: { type: "string" } },
      clarificationQuestions: { type: "array", items: { type: "string" } },
      designNotes: { type: "array", items: { type: "string" } },
      nextSteps: { type: "array", items: { type: "string" } },
      spec: {
        type: "object",
        additionalProperties: false,
        required: ["title", "badgeName", "description", "fields", "docxContent", "designSettings"],
        properties: {
          title: { type: "string" },
          badgeName: { type: "string" },
          description: { type: "string" },
          fields: {
            type: "array",
            items: {
              type: "object",
              additionalProperties: false,
              required: ["label", "type", "required", "additionalDetails", "options", "conditions", "sortOrder", "docxOutput"],
              properties: {
                label: { type: "string" },
                type: { enum: ["text", "textarea", "number", "date", "time", "checkbox", "multi_answer", "select", "signature", "photo"] },
                required: { type: "boolean" },
                additionalDetails: { type: "string" },
                options: { type: "array", items: { type: "string" } },
                conditions: {
                  type: "array",
                  items: {
                    type: "object",
                    additionalProperties: false,
                    required: ["questionLabel", "operator", "value"],
                    properties: {
                      questionLabel: { type: "string" },
                      operator: { enum: ["EQ", "NEQ", "CON", "NCON", "LT", "GT", "LTE", "GTE"] },
                      value: { type: "string" },
                    },
                  },
                },
                sortOrder: { type: "string" },
                docxOutput: { enum: ["raw", "checkboxList"] },
              },
            },
          },
          docxContent: {
            type: "object",
            additionalProperties: false,
            required: ["style", "sections"],
            properties: {
              style: { type: "string" },
              sections: {
                type: "array",
                items: {
                  type: "object",
                  additionalProperties: false,
                  required: ["title", "content", "isStandardHeader", "isStandardFooter", "layout"],
                  properties: {
                    title: { type: "string" },
                    content: { type: "string" },
                    isStandardHeader: { type: "boolean" },
                    isStandardFooter: { type: "boolean" },
                    layout: { enum: ["text", "table", "grid"] },
                  },
                },
              },
            },
          },
          designSettings: {
            type: "object",
            additionalProperties: false,
            required: [
              "companyName",
              "designBrief",
              "aiChooseDesign",
              "stylePreset",
              "primaryColor",
              "accentColor",
              "headerColor",
              "footerColor",
              "tableHeaderColor",
              "tableBorderColor",
              "fontFamily",
              "bodyFontSize",
              "headingStyle",
              "tableStyle",
              "logoPlacement",
              "logoWidth",
              "pageMargin",
              "headerText",
              "footerText",
            ],
            properties: {
              companyName: { type: "string" },
              designBrief: { type: "string" },
              aiChooseDesign: { type: "boolean" },
              stylePreset: { enum: ["ai", "clean_trade", "corporate", "source_replica", "minimal"] },
              primaryColor: { type: "string" },
              accentColor: { type: "string" },
              headerColor: { type: "string" },
              footerColor: { type: "string" },
              tableHeaderColor: { type: "string" },
              tableBorderColor: { type: "string" },
              fontFamily: { enum: ["Arial", "Calibri", "Inter", "Aptos"] },
              bodyFontSize: { type: "number" },
              headingStyle: { enum: ["bar", "underline", "boxed"] },
              tableStyle: { enum: ["source", "minimal", "grid"] },
              logoPlacement: { enum: ["left", "center", "right"] },
              logoWidth: { type: "number" },
              pageMargin: { type: "number" },
              headerText: { type: "string" },
              footerText: { type: "string" },
            },
          },
        },
      },
    },
  },
};

function codexHandoffBrief(body: any) {
  const fileLines = (body.attachments || [])
    .map((file: any) => `- ${file.name} (${file.mimeType || "unknown"}, ${file.size || 0} bytes)`)
    .join("\n") || "- No files uploaded in the web app yet.";

  return [
    "Codex handoff brief",
    "",
    `Mode: ${body.mode || "create_sm8f"}`,
    "",
    "User request:",
    body.prompt || "Please inspect the uploaded files and draft the ServiceM8 form plan.",
    "",
    "Uploaded files:",
    fileLines,
    "",
    "Design settings:",
    JSON.stringify(body.designSettings || {}, null, 2),
    "",
    "Codex should reply with a plan first, ask clarification questions where needed, list the proposed ServiceM8 questions/sections/logic for user review, then produce/import a FormSpec JSON for this app to build into .sm8f. Include designSettings in the FormSpec. If aiChooseDesign is true, choose suitable colours, typography, heading/table style and report treatment from the user design brief.",
  ].join("\n");
}

async function writeCodexBridgeRequest(body: any) {
  const requestId = randomUUID();
  const requestsDir = path.join(CODEX_BRIDGE_DIR, "requests");
  const responsesDir = path.join(CODEX_BRIDGE_DIR, "responses");
  await fs.mkdir(requestsDir, { recursive: true });
  await fs.mkdir(responsesDir, { recursive: true });

  const brief = codexHandoffBrief(body);
  const requestPath = path.join(requestsDir, `${requestId}.md`);
  const responsePath = path.join(responsesDir, `${requestId}.json`);
  const payload = [
    brief,
    "",
    "Response file expected at:",
    responsePath,
    "",
    "Response format:",
    "{",
    '  "summary": "...",',
    '  "assumptions": [],',
    '  "clarificationQuestions": [],',
    '  "designNotes": [],',
    '  "nextSteps": [],',
    '  "spec": { "title": "...", "badgeName": "...", "description": "...", "fields": [], "docxContent": { "sections": [] }, "designSettings": { "...": "..." } }',
    "}",
  ].join("\n");

  await fs.writeFile(requestPath, payload, "utf8");
  await fs.writeFile(path.join(CODEX_BRIDGE_DIR, "latest-request.md"), payload, "utf8");

  return { requestId, brief: payload, requestPath, responsePath };
}

function openAiInstructions(mode: string) {
  return `You are Autoform, an expert ServiceM8 .sm8f form designer. Build structured ServiceM8 form specs from user scopes and uploaded reference files.
Mode: ${mode}.
For create_from_prompt, build from the user's written scope first and return the proposed questions, sections, answer types, required flags, and conditional logic for review. Return a practical draft users can edit. Use clean alphanumeric field labels. Badge names must be 11 chars or less. Prefer existing ServiceM8/job merge fields for admin data. For multi-answer driven display logic, write display conditions as EQ/CON in the spec; the builder will invert multi-answer skip logic to NCON. DOCX sections should use editable table layout unless the user asks otherwise. Include designSettings in the spec. If aiChooseDesign is true, choose design elements that suit the trade/compliance context and explain the choices in designNotes.`;
}

function attachmentPlanningContext(attachments: any[] = []) {
  return attachments.map((file) => {
    const summary = `Uploaded file: ${file.name} (${file.mimeType}, ${file.size} bytes).`;
    const canReadInline =
      String(file.mimeType || "").startsWith("text/") ||
      /\.(csv|json|md|txt|xml|html)$/i.test(String(file.name || ""));

    if (!canReadInline || !file.data) {
      return `${summary} Binary content is available to the application; infer structure from the file name and user instructions unless explicit extracted text is provided.`;
    }

    try {
      const base64 = String(file.data).split(",").pop() || "";
      const text = Buffer.from(base64, "base64").toString("utf8").slice(0, 12000);
      return `${summary}\nExtracted text excerpt:\n${text}`;
    } catch {
      return `${summary} The content could not be decoded for inline planning.`;
    }
  }).join("\n\n");
}

async function callOpenAiPlan(body: any, revision = false) {
  if (!process.env.OPENAI_API_KEY) {
    const bridge = await writeCodexBridgeRequest(body);
    const error = new Error("Codex local mode is active. The web app will not generate test AI fields without OPENAI_API_KEY.");
    (error as any).statusCode = 409;
    (error as any).codexBrief = bridge.brief;
    (error as any).requestId = bridge.requestId;
    (error as any).requestPath = bridge.requestPath;
    (error as any).responsePath = bridge.responsePath;
    throw error;
  }

  const fileText = attachmentPlanningContext(body.attachments || []);
  const currentSpec = revision ? `Current spec:\n${JSON.stringify(body.currentSpec, null, 2)}` : "";
  const input = `${openAiInstructions(body.mode || "create_sm8f")}\n\nUser request:\n${body.prompt || ""}\n\n${fileText}\n\nDesign settings:\n${JSON.stringify(body.designSettings || {}, null, 2)}\n\n${currentSpec}`;
  const fileInputs = (body.attachments || [])
    .filter((file: any) => file.data && /pdf|officedocument|msword|text|csv|json|image/i.test(`${file.mimeType} ${file.name}`))
    .slice(0, 5)
    .map((file: any) => ({
      type: "input_file",
      filename: file.name,
      file_data: `data:${file.mimeType || "application/octet-stream"};base64,${file.data}`,
    }));

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      input: [
        {
          role: "user",
          content: [
            { type: "input_text", text: input },
            ...fileInputs,
          ],
        },
      ],
      text: {
        format: {
          type: "json_schema",
          ...formPlanSchema,
        },
      },
    }),
  });

  const data: any = await response.json();
  if (!response.ok) {
    console.error("[OpenAI] Error:", JSON.stringify(data, null, 2));
    throw new Error(data.error?.message || "OpenAI plan generation failed");
  }

  const outputText = data.output_text || data.output?.flatMap((item: any) => item.content || []).find((item: any) => item.type === "output_text")?.text;
  if (!outputText) throw new Error("OpenAI returned no structured plan text");
  return JSON.parse(outputText);
}

app.post("/api/ai/form-plan", async (req, res) => {
  try {
    res.json(await callOpenAiPlan(req.body));
  } catch (error: any) {
    res.status(error.statusCode || 500).json({
      error: error.message || "Failed to generate form plan",
      codexBrief: error.codexBrief,
      requestId: error.requestId,
      requestPath: error.requestPath,
      responsePath: error.responsePath,
    });
  }
});

app.post("/api/ai/revise-plan", async (req, res) => {
  try {
    res.json(await callOpenAiPlan(req.body, true));
  } catch (error: any) {
    res.status(error.statusCode || 500).json({
      error: error.message || "Failed to revise form plan",
      codexBrief: error.codexBrief,
      requestId: error.requestId,
      requestPath: error.requestPath,
      responsePath: error.responsePath,
    });
  }
});

app.get("/api/ai/status", (_req, res) => {
  const hasOpenAiKey = !!process.env.OPENAI_API_KEY;
  res.json({ hasOpenAiKey, mode: hasOpenAiKey ? "openai" : "codex_handoff" });
});

app.get("/api/codex/response/:requestId", async (req, res) => {
  try {
    const requestId = String(req.params.requestId || "").replace(/[^a-f0-9-]/gi, "");
    if (!requestId) return res.status(400).json({ error: "Missing request id" });

    const responsePath = path.join(CODEX_BRIDGE_DIR, "responses", `${requestId}.json`);
    const raw = await fs.readFile(responsePath, "utf8");
    res.json(JSON.parse(raw));
  } catch (error: any) {
    if (error.code === "ENOENT") {
      res.status(404).json({ error: "Codex response is not ready yet" });
    } else {
      res.status(500).json({ error: error.message || "Failed to load Codex response" });
    }
  }
});

function runNodeScript(scriptPath: string, args: string[]) {
  return new Promise<void>((resolve, reject) => {
    const child = spawn(process.execPath, [scriptPath, ...args], { cwd: process.cwd(), stdio: "pipe" });
    let stderr = "";
    child.stderr.on("data", (chunk) => { stderr += chunk.toString(); });
    child.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(stderr || `Script failed with exit code ${code}`));
    });
  });
}

app.post("/api/sm8f/build", async (req, res) => {
  try {
    const spec = req.body?.spec;
    if (!spec) return res.status(400).json({ error: "Missing spec" });
    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "autoform-"));
    const specPath = path.join(tempDir, "spec.json");
    const outputPath = path.join(tempDir, "output.sm8f");
    const buildSpec = { ...spec, docxContent: { ...(spec.docxContent || {}) } };
    const logo = spec.designSettings?.logo;
    if (logo?.data && logo?.mimeType?.includes("png")) {
      const logoPath = path.join(tempDir, "logo.png");
      await fs.writeFile(logoPath, Buffer.from(logo.data, "base64"));
      buildSpec.docxContent.logoPath = logoPath;
    }
    await fs.writeFile(specPath, JSON.stringify(buildSpec, null, 2), "utf8");
    await runNodeScript(path.join(process.cwd(), "scripts", "generate-sm8f.mjs"), [specPath, outputPath]);
    const file = await fs.readFile(outputPath);
    res.setHeader("Content-Type", "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="${cleanServiceM8Label(spec.title || "autoform").replace(/\s+/g, "-").toLowerCase()}.sm8f"`);
    res.send(file);
  } catch (error: any) {
    res.status(500).json({ error: error.message || "Failed to build SM8F" });
  }
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

if (!process.env.VERCEL) {
  startServer();
}

export default app;
