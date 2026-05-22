export type ToolMode = "create_sm8f" | "update_sm8f" | "swms_builder";

export type GenerationState = "idle" | "draft" | "needsClarification" | "approved" | "building" | "ready" | "error";

export interface AttachmentPayload {
  id: string;
  name: string;
  mimeType: string;
  size: number;
  data: string;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  timestamp: string;
  attachments?: AttachmentPayload[];
}

export interface FieldCondition {
  questionLabel: string;
  operator: "EQ" | "NEQ" | "CON" | "NCON" | "LT" | "GT" | "LTE" | "GTE";
  value: string;
}

export interface ServiceM8Field {
  label: string;
  type: "text" | "textarea" | "number" | "date" | "time" | "checkbox" | "multi_answer" | "select" | "signature" | "photo";
  options?: string[];
  required: boolean;
  additionalDetails?: string;
  conditions?: FieldCondition[];
  sortOrder?: string | number;
  docxOutput?: "raw" | "checkboxList";
}

export interface DocxSection {
  title: string;
  content: string;
  isStandardHeader?: boolean;
  isStandardFooter?: boolean;
  layout?: "text" | "table" | "grid";
  displayWhen?: FieldCondition | FieldCondition[];
}

export interface DesignSettings {
  companyName: string;
  logo?: AttachmentPayload | null;
  primaryColor: string;
  accentColor: string;
  headerColor: string;
  footerColor: string;
  tableHeaderColor: string;
  tableBorderColor: string;
  fontFamily: "Arial" | "Calibri" | "Inter" | "Aptos";
  bodyFontSize: number;
  headingStyle: "bar" | "underline" | "boxed";
  tableStyle: "source" | "minimal" | "grid";
  logoPlacement: "left" | "center" | "right";
  logoWidth: number;
  pageMargin: number;
  headerText: string;
  footerText: string;
}

export interface FormSpec {
  title: string;
  badgeName: string;
  description: string;
  fields: ServiceM8Field[];
  docxContent: {
    style?: string;
    logoPath?: string;
    sections: DocxSection[];
  };
  designSettings?: DesignSettings;
}

export interface AiPlanResponse {
  summary: string;
  assumptions: string[];
  clarificationQuestions: string[];
  designNotes: string[];
  nextSteps: string[];
  spec: FormSpec;
}

export interface ProjectRecord {
  id: string;
  title: string;
  mode: ToolMode;
  updatedAt: string;
  messages: ChatMessage[];
  spec: FormSpec | null;
  state: GenerationState;
  designSettings: DesignSettings;
  codexRequestId?: string;
  codexRequestPath?: string;
  codexResponsePath?: string;
}

export type GeneratedForm = FormSpec;

export interface AppSettings {
  headerText: string;
  footerText: string;
  companyName: string;
  showTable: boolean;
  servicem8ApiKey?: string;
}
