import { GeneratedForm } from "../types";

export interface ReferenceFilePayload {
  data: string;
  mimeType: string;
  name: string;
}

export async function generateFormStructure(
  prompt: string,
  type: string,
  referenceFile?: ReferenceFilePayload
): Promise<GeneratedForm> {
  const response = await fetch("/api/generate-form", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, type, referenceFile }),
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.error || "Codex generation is handled outside the web app.");
  }

  return data;
}
