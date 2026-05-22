import { AiPlanResponse, AttachmentPayload, DesignSettings, FormSpec, ToolMode } from "../types";

export type ReferenceFilePayload = AttachmentPayload;

async function readJson<T>(response: Response): Promise<T> {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.error || data.message || "Request failed");
  }
  return data as T;
}

export async function requestFormPlan(input: {
  prompt: string;
  mode: ToolMode;
  attachments: AttachmentPayload[];
  designSettings: DesignSettings;
}): Promise<AiPlanResponse> {
  const response = await fetch("/api/ai/form-plan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });

  return readJson<AiPlanResponse>(response);
}

export async function reviseFormPlan(input: {
  prompt: string;
  mode: ToolMode;
  currentSpec: FormSpec;
  designSettings: DesignSettings;
}): Promise<AiPlanResponse> {
  const response = await fetch("/api/ai/revise-plan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });

  return readJson<AiPlanResponse>(response);
}

export async function buildSm8f(spec: FormSpec): Promise<Blob> {
  const response = await fetch("/api/sm8f/build", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ spec }),
  });

  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || "Failed to build SM8F");
  }

  return response.blob();
}

export async function checkAdminAccess(password: string): Promise<boolean> {
  const response = await fetch("/api/auth/check", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ password }),
  });

  const data = await readJson<{ ok: boolean }>(response);
  return data.ok;
}

export async function generateFormStructure(): Promise<FormSpec> {
  throw new Error("Use requestFormPlan instead.");
}
