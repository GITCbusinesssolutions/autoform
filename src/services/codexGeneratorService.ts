import { AiPlanResponse, AttachmentPayload, DesignSettings, FormSpec, SwmsLibraryItem, ToolMode } from "../types";
import type { ProjectRecord } from "../types";

export type ReferenceFilePayload = AttachmentPayload;

async function readJson<T>(response: Response): Promise<T> {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(data.error || data.message || "Request failed") as Error & Record<string, unknown>;
    Object.assign(error, data);
    throw error;
  }
  return data as T;
}

export async function getAiStatus(): Promise<{ hasOpenAiKey: boolean; mode: "openai" | "codex_handoff" }> {
  const response = await fetch("/api/ai/status");
  return readJson(response);
}

export async function loadCodexResponse(requestId: string): Promise<AiPlanResponse> {
  const response = await fetch(`/api/codex/response/${encodeURIComponent(requestId)}`);
  return readJson<AiPlanResponse>(response);
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

export async function loadProjectsFromDb(): Promise<ProjectRecord[]> {
  const response = await fetch("/api/projects");
  const data = await readJson<{ projects: ProjectRecord[] }>(response);
  return data.projects || [];
}

export async function saveProjectsToDb(projects: ProjectRecord[]): Promise<void> {
  await fetch("/api/projects", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ projects }),
  });
}

export async function saveOpenAiKey(apiKey: string): Promise<{ ok: boolean; message: string }> {
  const response = await fetch("/api/settings/openai-key", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ apiKey }),
  });
  return readJson(response);
}

export async function loadSwmsLibrary(): Promise<SwmsLibraryItem[]> {
  const response = await fetch("/api/swms/library");
  const data = await readJson<{ items: SwmsLibraryItem[] }>(response);
  return data.items || [];
}

export async function uploadSwmsLibraryItem(input: { title: string; file: AttachmentPayload }): Promise<{ item: SwmsLibraryItem; duplicate: boolean }> {
  const response = await fetch("/api/swms/library", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });
  return readJson(response);
}

export async function uploadSwmsProjectDocument(input: { projectId: string; title: string; file: AttachmentPayload }): Promise<{ item: SwmsLibraryItem }> {
  const response = await fetch("/api/swms/project-docx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });
  return readJson(response);
}

export async function deleteSwmsLibraryItem(id: string): Promise<void> {
  const response = await fetch(`/api/swms/library/${encodeURIComponent(id)}`, { method: "DELETE" });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || "Failed to delete SWMS library item");
  }
}

export async function buildSwmsSm8f(input: { title: string; selectedLibraryIds: string[]; selectedProjectDocumentIds: string[] }): Promise<Blob> {
  const response = await fetch("/api/swms/build", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });

  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || "Failed to build SWMS SM8F");
  }

  return response.blob();
}

export async function generateFormStructure(): Promise<FormSpec> {
  throw new Error("Use requestFormPlan instead.");
}
