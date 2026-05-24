#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import JSZip from "jszip";
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  Packer,
  Paragraph,
  SimpleField,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  XmlAttributeComponent,
  XmlComponent,
} from "docx";

const DEFAULT_CONDITIONS = [
  { question: "", operator: "", value: "" },
  { question: "", operator: "", value: "" },
  { question: "", operator: "", value: "" },
];
const CHECKED_BOX = "\u2611";
const UNCHECKED_BOX = "\u2610";
const SYMBOL_FONT = "Segoe UI Symbol";
const TABLE_VALUE_RUN = { size: 20 };
const REPLICA_VALUE_RUN = { size: 18, font: "Arial", color: "404040" };
const REPLICA_LABEL_RUN = { size: 18, font: "Arial", color: "404040" };

function hexColor(value, fallback) {
  const cleaned = String(value || "").replace(/[^a-fA-F0-9]/g, "").slice(0, 6);
  return cleaned.length === 6 ? cleaned.toUpperCase() : fallback;
}

function designSettings(spec) {
  const design = spec.designSettings || {};
  return {
    companyName: design.companyName || "Autoform",
    designBrief: design.designBrief || "",
    primaryColor: hexColor(design.primaryColor, "111827"),
    accentColor: hexColor(design.accentColor, "9FDB1D"),
    headerColor: hexColor(design.headerColor || design.primaryColor, "3F3F46"),
    footerColor: hexColor(design.footerColor || design.primaryColor, "111827"),
    tableHeaderColor: hexColor(design.tableHeaderColor || design.headerColor || design.primaryColor, "3F3F46"),
    tableBorderColor: hexColor(design.tableBorderColor, "D4D4D8"),
    fontFamily: design.fontFamily || "Arial",
    bodyFontSize: Number(design.bodyFontSize || 9),
    headingStyle: design.headingStyle || "bar",
    tableStyle: design.tableStyle || "source",
    logoPlacement: design.logoPlacement || "left",
    logoWidth: Number(design.logoWidth || 150),
    pageMargin: Number(design.pageMargin || 720),
    headerText: design.headerText || spec.title || "ServiceM8 Form",
    footerText: design.footerText || "Commercial-in-confidence",
  };
}

class FieldCharAttributes extends XmlAttributeComponent {
  constructor(type) {
    super({ type });
    this.xmlKeys = { type: "w:fldCharType" };
  }
}

class FieldChar extends XmlComponent {
  constructor(type) {
    super("w:fldChar");
    this.root.push(new FieldCharAttributes(type));
  }
}

class InstrTextAttributes extends XmlAttributeComponent {
  constructor() {
    super({ space: "preserve" });
    this.xmlKeys = { space: "xml:space" };
  }
}

class InstrText extends XmlComponent {
  constructor(text) {
    super("w:instrText");
    this.root.push(new InstrTextAttributes());
    this.root.push(text);
  }
}

function fieldCharRun(type, options = {}) {
  return new TextRun({ children: [new FieldChar(type)], ...options });
}

function instructionRun(text, options = {}) {
  return new TextRun({ children: [new InstrText(text)], ...options });
}

function nowStamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
  ].join("-") + " " + [
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds()),
  ].join(":");
}

function cleanLabel(value, fallback = "Field") {
  const cleaned = String(value || fallback)
    .replace(/[^a-zA-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || fallback;
}

function humanizeFieldLabel(value, fallback = "Field") {
  const raw = String(value || fallback).trim();
  if (!raw) return fallback;
  if (/\s/.test(raw) && !raw.includes("_")) return raw;

  const smallWords = new Set(["a", "an", "and", "as", "at", "by", "for", "from", "if", "in", "of", "on", "or", "the", "to", "with"]);
  const acronyms = new Set(["AC", "BMS", "DB", "ETA", "JSA", "LOTO", "RCD", "SCADA", "SM8F", "SWMS"]);
  return raw
    .replace(/[_-]+/g, " ")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .map((word, index) => {
      const upper = word.toUpperCase();
      const lower = word.toLowerCase();
      if (index === 0 && upper === "NO") return "No.";
      if (acronyms.has(upper)) return upper;
      if (index > 0 && smallWords.has(lower)) return lower;
      return `${word.charAt(0).toUpperCase()}${word.slice(1).toLowerCase()}`;
    })
    .join(" ");
}

function slug(value) {
  return cleanLabel(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function optionCode(value) {
  return cleanLabel(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function badgeName(value, fallback) {
  return cleanLabel(value || fallback || "Form").substring(0, 11);
}

function formFieldType(field) {
  const type = String(field.type || "text").toLowerCase();
  if (type === "textarea" || type === "multiline" || type === "multi-line") return "Text (Multi-Line)";
  if (type === "number") return "Number";
  if (type === "date") return "Date";
  if (type === "signature") return "Signature";
  if (type === "photo" || type === "image") return "Photo";
  if (type === "checkbox" || type === "multi_answer" || type === "multi-answer") return "Multiple Choice (Multi-Answer)";
  if (type === "select" || type === "dropdown" || type === "multiple_choice" || type === "multiple-choice") return "Multiple Choice";
  return field.multiLine ? "Text (Multi-Line)" : "Text";
}

function mergeFieldName(field) {
  const base = slug(field.label || field.name);
  return formFieldType(field) === "Photo"
    ? `image_form_${base}_medium`
    : `form_${base}`;
}

function optionMergeFieldName(field, option) {
  return `form_${slug(field.label || field.name)}_${optionCode(option)}`;
}

function mergeField(fieldName) {
  return `{MERGEFIELD ${fieldName} \\* MERGEFORMAT}`;
}

function complexMergeFieldRuns(fieldName, placeholder = "", runOptions = {}) {
  const normalizedFieldName = String(fieldName || "").trim();
  return [
    fieldCharRun("begin", runOptions),
    instructionRun(` MERGEFIELD ${normalizedFieldName} \\* MERGEFORMAT `, runOptions),
    fieldCharRun("separate", runOptions),
    new TextRun({ text: placeholder, ...runOptions }),
    fieldCharRun("end", runOptions),
  ];
}

function conditionalTextRuns({ fieldName, operator = "=", value = "Yes", whenTrue = CHECKED_BOX, whenFalse = UNCHECKED_BOX }, runOptions = TABLE_VALUE_RUN) {
  const symbolOptions = { ...runOptions, font: SYMBOL_FONT };
  return [
    fieldCharRun("begin", runOptions),
    instructionRun(' IF "', runOptions),
    ...complexMergeFieldRuns(fieldName, "", runOptions),
    instructionRun(`" ${operator} "${value}" "`, symbolOptions),
    instructionRun(whenTrue, symbolOptions),
    instructionRun('" "', symbolOptions),
    instructionRun(whenFalse, symbolOptions),
    instructionRun('" ', runOptions),
    fieldCharRun("separate", runOptions),
    new TextRun({ text: whenFalse, ...symbolOptions }),
    fieldCharRun("end", runOptions),
  ];
}

function checkboxOptionRuns(field, option, runOptions = TABLE_VALUE_RUN) {
  return [
    ...conditionalTextRuns({
      fieldName: optionMergeFieldName(field, option),
      value: "Yes",
      whenTrue: CHECKED_BOX,
      whenFalse: UNCHECKED_BOX,
    }, runOptions),
    new TextRun({ text: " ", ...runOptions }),
    new TextRun({ text: option, ...runOptions }),
  ];
}

function toConditionArray(field, uuidByLabel) {
  const raw = Array.isArray(field.conditions) ? field.conditions : [];
  const mapped = raw.slice(0, 3).map((condition) => {
    const target =
      condition.question ||
      condition.questionUuid ||
      uuidByLabel.get(cleanLabel(condition.questionLabel || condition.field || ""));

    return {
      question: target || "",
      operator: condition.operator || "",
      value: condition.value || "",
    };
  });

  while (mapped.length < 3) mapped.push({ question: "", operator: "", value: "" });
  return mapped;
}

function invertOperator(operator, targetField) {
  const normalized = String(operator || "").toUpperCase();
  const targetFieldType = targetField ? formFieldType(targetField) : "";
  if (targetFieldType === "Multiple Choice (Multi-Answer)") {
    if (normalized === "EQ") return "NCON";
    if (normalized === "NEQ") return "CON";
  }

  const map = {
    EQ: "NEQ",
    NEQ: "EQ",
    CON: "NCON",
    NCON: "CON",
    LT: "GTE",
    GT: "LTE",
    LTE: "GT",
    GTE: "LT",
  };
  return map[normalized] || normalized;
}

function conditionTarget(condition, uuidByLabel) {
  return (
    condition.question ||
    condition.questionUuid ||
    uuidByLabel.get(cleanLabel(condition.questionLabel || condition.field || ""))
  );
}

function conditionTargetField(condition, fieldsByLabel, fieldsByUuid) {
  const label = cleanLabel(condition.questionLabel || condition.field || "");
  return (
    fieldsByLabel.get(label) ||
    fieldsByUuid.get(condition.question || condition.questionUuid || "")
  );
}

function toSkipConditionArray(field, uuidByLabel, fieldsByLabel, fieldsByUuid) {
  const displayConditions = Array.isArray(field.conditions) ? field.conditions : [];
  const inverted = displayConditions.slice(0, 3).map((condition) => {
    const target = conditionTarget(condition, uuidByLabel);
    const targetField = conditionTargetField(condition, fieldsByLabel, fieldsByUuid);

    return {
      question: target || "",
      operator: invertOperator(condition.operator, targetField),
      value: condition.value || "",
    };
  });

  while (inverted.length < 3) inverted.push({ question: "", operator: "", value: "" });
  return inverted;
}

function serviceM8ConditionMethod(field) {
  const hasDisplayConditions = Array.isArray(field.conditions) && field.conditions.length > 0;
  if (!hasDisplayConditions) return field.conditionMethod || "AND";

  const displayMethod = String(field.conditionMethod || "AND").toUpperCase();
  return displayMethod === "OR" ? "AND" : "OR";
}

function fieldLookup(fields) {
  return new Map(fields.map((field) => [cleanLabel(field.label || field.name), field]));
}

function docxOutputType(field) {
  if (!field) return "";
  if (typeof field.docxOutput === "string") return field.docxOutput;
  if (field.docxOutput && typeof field.docxOutput === "object") return field.docxOutput.type || "";
  return "";
}

function shouldRenderCheckboxList(field) {
  if (!field || !formFieldType(field).startsWith("Multiple Choice")) return false;
  if (["multi_answer", "multi-answer", "checkbox"].includes(String(field.type || "").toLowerCase())) {
    return Array.isArray(field.options) && field.options.length > 0;
  }
  const output = String(docxOutputType(field)).toLowerCase();
  if (output === "raw") return false;
  return output === "checkboxlist" || (Array.isArray(field.options) && field.options.length > 0);
}

function reportValueForField(field) {
  return shouldRenderCheckboxList(field)
    ? `{CHECKLIST ${field.label}}`
    : mergeField(mergeFieldName(field));
}

function fieldReportRow(field) {
  return `${field.label}: ${reportValueForField(field)}`;
}

function textTokens(value) {
  return new Set(String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .split(/\s+/)
    .filter((token) => token.length > 2)
    .map((token) => {
      const aliases = {
        est: "estimated",
        eta: "quote",
        tech: "technician",
        tag: "tag",
        photos: "photo",
        before: "before",
        after: "after",
        completion: "completion",
        completed: "completed",
      };
      return aliases[token] || token;
    }));
}

function tokenScore(source, target) {
  const sourceTokens = textTokens(source);
  const targetTokens = textTokens(target);
  let score = 0;
  for (const token of sourceTokens) {
    if (targetTokens.has(token)) score += 1;
  }
  return score;
}

function bestFieldForText(text, fieldsByLabel, fields = []) {
  const exact = fieldsByLabel.get(cleanLabel(text));
  if (exact) return exact;

  let best;
  let bestScore = 0;
  for (const field of fields) {
    const score = tokenScore(text, `${field.label} ${field.additionalDetails || ""}`);
    if (score > bestScore) {
      best = field;
      bestScore = score;
    }
  }

  return bestScore >= 2 ? best : undefined;
}

function sectionFieldScore(section, field) {
  const sectionText = `${section.title || ""} ${section.content || ""}`;
  const conditionText = Array.isArray(field.conditions)
    ? field.conditions.map((condition) => `${condition.questionLabel || condition.field || ""} ${condition.value || ""}`).join(" ")
    : "";
  const fieldText = `${field.label || ""} ${field.additionalDetails || ""} ${conditionText}`;
  let score = tokenScore(sectionText, fieldText);
  const title = String(section.title || "").toLowerCase();
  const label = String(field.label || "").toLowerCase();
  const fieldFullText = fieldText.toLowerCase();
  const details = String(field.additionalDetails || "").toLowerCase();

  if (title.includes("arrival") && details.includes("before")) score += 4;
  if (title.includes("job detail") && /service|category|technician|attendance|start/.test(label)) score += 3;
  if (title.includes("time") && /time|variation|notified|ack|estimated|est/.test(label)) score += 4;
  if (title.includes("works") && /works|delay|system|status|exit/.test(label)) score += 4;
  if (title.includes("additional") && /minor|major|urgent|eta|quote/.test(label)) score += 4;
  if (title.includes("compliance") && /test tag|emergency lighting/.test(fieldFullText)) score += 5;
  if (title.includes("completion") && /completion|return|finish|after|close out/.test(label)) score += 4;
  if (title.includes("footer") || title.includes("header")) score = 0;

  return score;
}

function inferredSectionFields(section, fields = [], assignedLabels = new Set()) {
  return fields
    .filter((field) => !assignedLabels.has(cleanLabel(field.label)))
    .map((field) => ({ field, score: sectionFieldScore(section, field) }))
    .filter(({ score }) => score >= 3)
    .sort((a, b) => Number(a.field.sortOrder || 0) - Number(b.field.sortOrder || 0))
    .map(({ field }) => field);
}

function sectionContentRows(section, fieldsByLabel, fields = [], assignedLabels = new Set()) {
  const rawLines = String(section.content || "").split("\n").map((line) => line.trim()).filter(Boolean);
  const rows = [];
  let fieldRowCount = 0;

  for (const line of rawLines) {
    if (line.includes("{") || line.includes(":")) {
      rows.push(line);
      continue;
    }

    const parts = line.split(";").map((part) => part.trim()).filter(Boolean);
    if (parts.length > 1) {
      for (const part of parts) {
        const field = bestFieldForText(part, fieldsByLabel, fields);
        if (field) {
          rows.push(fieldReportRow(field));
          assignedLabels.add(cleanLabel(field.label));
          fieldRowCount += 1;
        } else {
          rows.push(part);
        }
      }
      continue;
    }

    const exactField = fieldsByLabel.get(cleanLabel(line));
    const looksLikeNarrative = /,|\band\b/i.test(line);
    const field = exactField || (looksLikeNarrative ? undefined : bestFieldForText(line, fieldsByLabel, fields));
    if (field) {
      rows.push(fieldReportRow(field));
      assignedLabels.add(cleanLabel(field.label));
      fieldRowCount += 1;
    } else {
      rows.push(line);
    }
  }

  if (fieldRowCount === 0 && !section.isStandardHeader && !section.isStandardFooter) {
    const inferred = inferredSectionFields(section, fields, assignedLabels);
    if (inferred.length) {
      for (const field of inferred) assignedLabels.add(cleanLabel(field.label));
      return inferred.map(fieldReportRow);
    }
  }

  return rows;
}

function wordOperator(operator) {
  const normalized = String(operator || "EQ").toUpperCase();
  const map = {
    EQ: "=",
    NEQ: "<>",
    LT: "<",
    GT: ">",
    LTE: "<=",
    GTE: ">=",
  };
  return map[normalized] || normalized;
}

function displayConditions(section) {
  if (Array.isArray(section.displayWhen)) return section.displayWhen;
  if (section.displayWhen) return [section.displayWhen];
  if (Array.isArray(section.conditions)) return section.conditions;
  return [];
}

function resolveDocxCondition(condition, fieldsByLabel) {
  const field =
    condition.questionLabel || condition.field
      ? fieldsByLabel.get(cleanLabel(condition.questionLabel || condition.field))
      : undefined;
  const operator = String(condition.operator || "EQ").toUpperCase();

  if (condition.mergeFieldName) {
    return {
      fieldName: condition.mergeFieldName,
      operator: wordOperator(operator),
      value: String(condition.value ?? "Yes"),
    };
  }

  if (field) {
    const isMultiAnswer = formFieldType(field) === "Multiple Choice (Multi-Answer)";
    if (isMultiAnswer && condition.value && ["EQ", "CON", "NEQ", "NCON"].includes(operator)) {
      return {
        fieldName: optionMergeFieldName(field, condition.value),
        operator: ["NEQ", "NCON"].includes(operator) ? "<>" : "=",
        value: "Yes",
      };
    }

    const useOptionMergeField = condition.useOptionMergeField ??
      (formFieldType(field).startsWith("Multiple Choice") && operator === "EQ" && condition.value);

    if (useOptionMergeField) {
      return {
        fieldName: optionMergeFieldName(field, condition.value),
        operator: "=",
        value: "Yes",
      };
    }

    return {
      fieldName: mergeFieldName(field),
      operator: wordOperator(operator),
      value: String(condition.value ?? ""),
    };
  }

  return {
    fieldName: condition.fieldName || "",
    operator: wordOperator(operator),
    value: String(condition.value ?? ""),
  };
}

function conditionalSectionStartRuns(section, fieldsByLabel, runOptions) {
  const [condition] = displayConditions(section);
  if (!condition) return [];

  const resolved = resolveDocxCondition(condition, fieldsByLabel);
  return [
    fieldCharRun("begin", runOptions),
    instructionRun(' IF "', runOptions),
    ...complexMergeFieldRuns(resolved.fieldName),
    instructionRun(`" ${resolved.operator} "${resolved.value}" "`, runOptions),
    instructionRun(section.title, runOptions),
  ];
}

function conditionalSectionEndParagraph(section, fieldsByLabel, runOptions) {
  if (!displayConditions(section).length) return undefined;
  return new Paragraph({
    children: [
      instructionRun('" "" ', runOptions),
      fieldCharRun("separate", runOptions),
      new TextRun({ text: " " }),
      fieldCharRun("end", runOptions),
    ],
    spacing: { after: 120 },
  });
}

function normalizeSpec(spec) {
  const title = cleanLabel(spec.title || spec.name || "Generated ServiceM8 Form", "Generated ServiceM8 Form");
  const fields = Array.isArray(spec.fields) ? spec.fields : [];
  const labelMap = new Map();
  const normalizedFields = fields.map((field, index) => {
    const originalLabel = String(field.label || field.name || `Field ${index + 1}`);
    const label = cleanLabel(humanizeFieldLabel(originalLabel, `Field ${index + 1}`));
    labelMap.set(originalLabel, label);
    labelMap.set(originalLabel.replace(/\s+/g, ""), label);
    return {
      ...field,
      uuid: field.uuid || crypto.randomUUID(),
      label,
      sortOrder: String(field.sortOrder || field.sort_order || index + 1),
      required: !!field.required,
      additionalDetails: String(field.additionalDetails || ""),
      options: Array.isArray(field.options) ? field.options.map((option) => String(option || "").trim()).filter(Boolean) : undefined,
    };
  });
  const normalizeCondition = (condition = {}) => ({
    ...condition,
    questionLabel: labelMap.get(condition.questionLabel) || humanizeFieldLabel(condition.questionLabel || ""),
  });
  const replaceLabels = (content = "") => {
    let next = String(content || "");
    for (const [from, to] of labelMap.entries()) {
      if (!from || from === to) continue;
      next = next.replace(new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), to);
    }
    return next;
  };
  const docxContent = spec.docxContent || {};
  const normalizedSections = Array.isArray(docxContent.sections)
    ? docxContent.sections.map((section) => ({
      ...section,
      content: replaceLabels(section.content || ""),
      displayWhen: Array.isArray(section.displayWhen)
        ? section.displayWhen.map(normalizeCondition)
        : section.displayWhen
          ? normalizeCondition(section.displayWhen)
          : section.displayWhen,
    }))
    : [];

  return {
    title,
    badgeName: badgeName(spec.badgeName || spec.badge_name, title),
    description: String(spec.description || ""),
    badgeMandatoryState: String(spec.badgeMandatoryState || spec.badge_mandatory_state || "2"),
    canBeUsedIndependently: String(spec.canBeUsedIndependently || spec.can_be_used_independently || "0"),
    templateFields: Array.isArray(spec.templateFields) ? spec.templateFields : [],
    fields: normalizedFields.map((field) => ({
      ...field,
      conditions: Array.isArray(field.conditions) ? field.conditions.map(normalizeCondition) : field.conditions,
    })),
    docxContent: { ...docxContent, sections: normalizedSections },
    designSettings: spec.designSettings || {},
  };
}

function isReplicaSpec(spec) {
  return String(spec.docxContent?.style || "").toLowerCase() === "sourcepdfreplica";
}

function buildFormJson(spec) {
  const normalized = normalizeSpec(spec);
  const formUuid = spec.formUuid || crypto.randomUUID();
  const documentTemplateUuid = spec.documentTemplateUuid || crypto.randomUUID();
  const createStaffUuid = spec.createStaffUuid || spec.staffUuid || "083e6d79-d31d-4125-8eb9-1d368bb7380b";
  const editStaffUuid = spec.editStaffUuid || spec.staffUuid || "dedb066d-302f-409a-a1a0-5204eb87b1eb";
  const vendorUuid = spec.vendorUuid || "15494128-e0cd-4c53-8774-ea1ee6b5123b";
  const timestamp = nowStamp();
  const uuidByLabel = new Map(normalized.fields.map((field) => [cleanLabel(field.label), field.uuid]));
  const fieldsByLabel = fieldLookup(normalized.fields);
  const fieldsByUuid = new Map(normalized.fields.map((field) => [field.uuid, field]));

  return {
    form: {
      uuid: formUuid,
      create_by_staff_uuid: createStaffUuid,
      edit_by_staff_uuid: editStaffUuid,
      create_date: timestamp,
      edit_date: timestamp,
      active: "1",
      vendor_uuid: vendorUuid,
      document_template_uuid: documentTemplateUuid,
      name: ` ${normalized.title}`,
      badge_name: normalized.badgeName,
      is_sample_form: "0",
      sample_form_id: "0",
      can_be_used_independently: normalized.canBeUsedIndependently,
      badge_mandatory_state: normalized.badgeMandatoryState,
      prevent_form_from_export: "0",
      network_origin_form_uuid: "",
      network_origin_form_etag: "",
      store_item_uuid: "",
      template_fields_json: JSON.stringify(normalized.templateFields),
      is_locked: "0",
    },
    fields: normalized.fields
      .slice()
      .sort((a, b) => Number(a.sortOrder) - Number(b.sortOrder))
      .map((field) => {
        const fieldType = formFieldType(field);
        const fieldData = {
          fieldType,
          additionalDetails: field.additionalDetails || "",
          mandatory: !!field.required,
          conditions: toSkipConditionArray(field, uuidByLabel, fieldsByLabel, fieldsByUuid),
          conditionMethod: serviceM8ConditionMethod(field),
        };

        if (fieldType.startsWith("Multiple Choice")) {
          fieldData.choices = Array.isArray(field.options) && field.options.length ? field.options : ["Yes", "No"];
        }

        return {
          uuid: field.uuid,
          create_by_staff_uuid: createStaffUuid,
          edit_by_staff_uuid: editStaffUuid,
          create_date: timestamp,
          edit_date: timestamp,
          active: "1",
          form_uuid: formUuid,
          name: field.label,
          field_data_json: JSON.stringify(fieldData),
          sort_order: String(field.sortOrder),
        };
      }),
  };
}

function parseContentWithFields(text, runOptions = TABLE_VALUE_RUN) {
  const parts = [];
  const regex = /\{((?:[^{}]|\{[^{}]*\})+)\}/g;
  let lastIndex = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parts.push(new TextRun({ text: text.substring(lastIndex, match.index), ...runOptions }));
    }

    const fieldContent = match[1];
    if (fieldContent.startsWith("IF")) {
      parts.push(fieldCharRun("begin", runOptions));

      const innerRegex = /\{MERGEFIELD\s+(.*?)\s+\\\*\s+MERGEFORMAT\}/g;
      let innerLastIndex = 0;
      let innerMatch;

      while ((innerMatch = innerRegex.exec(fieldContent)) !== null) {
        if (innerMatch.index > innerLastIndex) {
          parts.push(instructionRun(fieldContent.substring(innerLastIndex, innerMatch.index), runOptions));
        }

        parts.push(...complexMergeFieldRuns(innerMatch[1], "", runOptions));
        innerLastIndex = innerRegex.lastIndex;
      }

      if (innerLastIndex < fieldContent.length) {
        parts.push(instructionRun(fieldContent.substring(innerLastIndex), runOptions));
      }

      parts.push(fieldCharRun("separate", runOptions));
      parts.push(new TextRun({ text: " ", ...runOptions }));
      parts.push(fieldCharRun("end", runOptions));
    } else if (fieldContent.startsWith("MERGEFIELD")) {
      parts.push(new SimpleField(fieldContent));
    } else {
      parts.push(new SimpleField(fieldContent));
    }

    lastIndex = regex.lastIndex;
  }

  if (lastIndex < text.length) {
    parts.push(new TextRun({ text: text.substring(lastIndex), ...runOptions }));
  }

  return parts;
}

function autoSections(spec) {
  const rows = spec.fields.map((field) => {
    return fieldReportRow(field);
  });

  return [
    {
      title: "Job Details",
      isStandardHeader: true,
      layout: "table",
      content: [
        `Date: ${mergeField("calculation.todays_date")}`,
        `Job Address: ${mergeField("job.job_address_singleline")}`,
        `Technician: ${mergeField("calculation.current_user_fullname")}`,
      ].join("\n"),
    },
    {
      title: spec.title,
      layout: "table",
      content: rows.join("\n"),
    },
  ];
}

function checklistMarkerField(value, fieldsByLabel) {
  const match = String(value || "").trim().match(/^\{CHECKLIST\s+(.+?)\}$/i);
  if (!match) return undefined;
  return fieldsByLabel.get(cleanLabel(match[1]));
}

function tableValueParagraphs(label, value, fieldsByLabel, runOptions = TABLE_VALUE_RUN) {
  const field = fieldsByLabel.get(cleanLabel(label)) || checklistMarkerField(value, fieldsByLabel);

  if (shouldRenderCheckboxList(field)) {
    return field.options.map((option) => new Paragraph({
      children: checkboxOptionRuns(field, option, runOptions),
      spacing: { after: 36 },
    }));
  }

  return [new Paragraph({ children: parseContentWithFields(value, runOptions) })];
}

function tableLineWeight(line, fieldsByLabel) {
  const [label, ...rest] = line.split(":");
  const value = rest.join(":").trim() || line;
  const field = fieldsByLabel.get(cleanLabel(label.trim())) || checklistMarkerField(value, fieldsByLabel);
  const optionWeight = shouldRenderCheckboxList(field)
    ? Math.max(1.2, field.options.length * 0.72)
    : 1;
  const labelWrapWeight = label.trim().length > 34 ? 0.35 : 0;
  return optionWeight + labelWrapWeight;
}

function rowField(line, fieldsByLabel) {
  const [label, ...rest] = String(line || "").split(":");
  const value = rest.join(":").trim() || line;
  return fieldsByLabel.get(cleanLabel(label.trim())) || checklistMarkerField(value, fieldsByLabel);
}

function rowDisplayConditions(line, fieldsByLabel) {
  const field = rowField(line, fieldsByLabel);
  return Array.isArray(field?.conditions) ? field.conditions : [];
}

function conditionSignature(conditions) {
  const [condition] = Array.isArray(conditions) ? conditions : [];
  if (!condition?.questionLabel && !condition?.field && !condition?.mergeFieldName) return "";
  return JSON.stringify({
    questionLabel: cleanLabel(condition.questionLabel || condition.field || condition.mergeFieldName || ""),
    operator: String(condition.operator || "EQ").toUpperCase(),
    value: String(condition.value ?? ""),
  });
}

function splitConditionalRowGroups(rows, fieldsByLabel, sectionHasConditions) {
  if (sectionHasConditions) return [{ rows, conditions: [] }];

  const groups = [];
  for (const row of rows) {
    const conditions = rowDisplayConditions(row, fieldsByLabel);
    const key = conditionSignature(conditions);
    const previous = groups[groups.length - 1];
    if (previous && previous.key === key) {
      previous.rows.push(row);
    } else {
      groups.push({ key, rows: [row], conditions: key ? conditions : [] });
    }
  }

  return groups;
}

function chunkTableRows(rows, section, fieldsByLabel) {
  if (section.isStandardHeader || rows.length <= 8) return [rows];

  const chunks = [];
  let current = [];
  let currentWeight = 0;
  const maxWeight = 11.5;

  for (const row of rows) {
    const weight = tableLineWeight(row, fieldsByLabel);
    if (current.length && currentWeight + weight > maxWeight) {
      chunks.push(current);
      current = [];
      currentWeight = 0;
    }
    current.push(row);
    currentWeight += weight;
  }

  if (current.length) chunks.push(current);
  return chunks;
}

function buildTableRows(rows, section, fieldsByLabel, replica = false, design = designSettings({})) {
  const labelWidth = replica ? 2300 : 3100;
  const valueWidth = replica ? 7060 : 6260;
  const bodySize = Math.max(16, Math.round(design.bodyFontSize * 2));
  const labelRun = replica
    ? { ...REPLICA_LABEL_RUN, font: design.fontFamily, color: design.primaryColor, size: bodySize }
    : { color: design.primaryColor, size: bodySize, font: design.fontFamily };
  const valueRun = replica
    ? { ...REPLICA_VALUE_RUN, font: design.fontFamily, size: bodySize }
    : { ...TABLE_VALUE_RUN, size: bodySize, font: design.fontFamily };
  const cellMargins = replica
    ? { top: 34, bottom: 34, left: 90, right: 90 }
    : { top: 100, bottom: 100, left: 120, right: 120 };

  return rows.map((line, index) => {
    const [label, ...rest] = line.split(":");
    const value = rest.join(":").trim() || line;
    return new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: label.trim(), bold: true, ...labelRun })],
          })],
          width: { size: labelWidth, type: WidthType.DXA },
          shading: { fill: section.isStandardHeader ? "EDEDED" : replica ? "FFFFFF" : index % 2 === 0 ? "F8FAFC" : "FFFFFF" },
          borders: {
            top: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
            bottom: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
            left: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
            right: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
          },
          margins: cellMargins,
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({
          children: tableValueParagraphs(label.trim(), value, fieldsByLabel, valueRun),
          width: { size: valueWidth, type: WidthType.DXA },
          borders: {
            top: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
            bottom: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
            left: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
            right: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
          },
          margins: cellMargins,
          verticalAlign: VerticalAlign.CENTER,
        }),
      ],
    });
  });
}

function buildTableHeadingRow(title, replica = false, design = designSettings({})) {
  return new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({
            text: title,
            bold: true,
            color: "FFFFFF",
            size: replica ? 20 : 24,
            font: design.fontFamily,
          })],
        })],
        columnSpan: 2,
        shading: { fill: design.tableHeaderColor },
        borders: {
          top: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
          bottom: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
          left: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
          right: { style: BorderStyle.SINGLE, color: design.tableBorderColor, size: 4 },
        },
        margins: replica
          ? { top: 75, bottom: 75, left: 90, right: 90 }
          : { top: 120, bottom: 120, left: 120, right: 120 },
      }),
    ],
  });
}

async function replicaIntro(spec) {
  const children = [];
  const design = designSettings(spec);
  let logoChildren;
  if (spec.docxContent.logoPath) {
    try {
      const logoPath = path.isAbsolute(spec.docxContent.logoPath)
        ? spec.docxContent.logoPath
        : path.resolve(process.cwd(), spec.docxContent.logoPath);
      const logoData = await fs.readFile(logoPath);
      logoChildren = [
        new Paragraph({
          children: [
            new ImageRun({
              type: "png",
              data: logoData,
              transformation: { width: design.logoWidth, height: Math.round(design.logoWidth * 0.46) },
              altText: {
                title: "Ambrose Construct Group logo",
                description: "Ambrose Construct Group logo extracted from the source PDF",
                name: "Ambrose Construct Group logo",
              },
            }),
          ],
          spacing: { after: 0 },
        }),
      ];
    } catch {
      logoChildren = undefined;
    }
  }

  logoChildren ||= [
    new Paragraph({
      children: [
        new TextRun({ text: `${design.companyName} `, bold: true, color: design.primaryColor, size: 21, font: design.fontFamily }),
        new TextRun({ text: "AUTOFORM", bold: true, color: design.accentColor, size: 21, font: design.fontFamily }),
      ],
      spacing: { after: 0 },
    }),
    new Paragraph({
      children: [new TextRun({ text: design.designBrief, color: design.primaryColor, size: 13, font: design.fontFamily })],
      indent: { left: 470 },
      spacing: { before: 0, after: 0 },
    }),
  ];

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3500, 5860],
    layout: TableLayoutType.AUTOFIT,
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      children: [
        new TableCell({
          children: logoChildren,
          width: { size: 3500, type: WidthType.DXA },
          margins: { top: 30, bottom: 90, left: 0, right: 0 },
          borders: {
            top: { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE },
          },
        }),
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text: design.companyName, bold: true, font: design.fontFamily, size: 20, color: design.primaryColor })],
              alignment: AlignmentType.RIGHT,
            }),
            new Paragraph({
              children: [
                new TextRun({ text: design.headerText, font: design.fontFamily, size: 14, color: design.primaryColor }),
              ],
              alignment: AlignmentType.RIGHT,
            }),
          ],
          width: { size: 5860, type: WidthType.DXA },
          margins: { top: 180, bottom: 120, left: 0, right: 0 },
          borders: {
            top: { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE },
          },
        }),
      ],
    })],
  }));

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: spec.title, bold: true, color: "FFFFFF", font: design.fontFamily, size: 30 })],
        })],
        shading: { fill: design.headerColor },
        margins: { top: 140, bottom: 140, left: 120, right: 120 },
      })],
    })],
  }));

  children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
  return children;
}

async function buildDocx(spec) {
  const normalized = normalizeSpec(spec);
  const design = designSettings(normalized);
  const replica = isReplicaSpec(normalized);
  const sections = Array.isArray(normalized.docxContent.sections) && normalized.docxContent.sections.length
    ? normalized.docxContent.sections
    : autoSections(normalized);
  const fieldsByLabel = fieldLookup(normalized.fields);
  const assignedDocxLabels = new Set();

  const children = replica
    ? await replicaIntro(normalized)
    : [
      new Paragraph({
        children: [new TextRun({ text: normalized.title, bold: true, size: 34, color: design.primaryColor, font: design.fontFamily })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 160 },
      }),
      new Paragraph({
        children: [new TextRun({ text: normalized.description || design.designBrief, italics: true, color: design.footerColor, size: 22, font: design.fontFamily })],
        spacing: { after: 360 },
      }),
    ];

  for (const section of sections) {
    const isTableSection = section.layout === "table" || section.isStandardHeader;
    const hasSectionConditions = displayConditions(section).length > 0;
    const useTableHeading = isTableSection && !section.isStandardHeader && !section.isStandardFooter && !hasSectionConditions;

    if (replica && section.isStandardHeader) {
      children.push(new Paragraph({
        children: [new TextRun({ text: String(section.title || "").toUpperCase(), bold: true, font: design.fontFamily, size: 22, color: design.primaryColor })],
        keepNext: true,
        spacing: { before: 90, after: 55 },
        border: { bottom: { color: design.tableBorderColor, space: 1, style: BorderStyle.SINGLE, size: 10 } },
      }));
    }

    if (!section.isStandardHeader && !section.isStandardFooter && !useTableHeading) {
      const headingRunOptions = { bold: true, size: 26, color: design.primaryColor, font: design.fontFamily };
      const headingChildren = hasSectionConditions
        ? conditionalSectionStartRuns(section, fieldsByLabel, headingRunOptions)
        : [new TextRun({ text: section.title, ...headingRunOptions })];

      children.push(new Paragraph({
        children: headingChildren,
        keepNext: true,
        spacing: { before: replica ? 180 : 280, after: replica ? 80 : 120 },
        border: { bottom: { color: design.tableBorderColor, space: 1, style: BorderStyle.SINGLE, size: replica ? 10 : 6 } },
      }));
    }

    if (isTableSection) {
      const rows = sectionContentRows(section, fieldsByLabel, normalized.fields, assignedDocxLabels);
      const rowGroups = splitConditionalRowGroups(rows, fieldsByLabel, hasSectionConditions);
      let hasRenderedHeadingRow = false;

      for (const group of rowGroups) {
        const hasGroupConditions = group.conditions.length > 0;
        if (hasGroupConditions) {
          children.push(new Paragraph({
            children: conditionalSectionStartRuns({ ...section, displayWhen: group.conditions }, fieldsByLabel, { bold: true, size: 22, color: design.primaryColor, font: design.fontFamily }),
            keepNext: true,
            spacing: { before: replica ? 120 : 180, after: replica ? 60 : 80 },
          }));
        }

        const chunks = chunkTableRows(group.rows, section, fieldsByLabel);
        for (const [chunkIndex, chunk] of chunks.entries()) {
          const tableRows = buildTableRows(chunk, section, fieldsByLabel, replica, design);
          if (useTableHeading && !hasRenderedHeadingRow && !hasGroupConditions && chunkIndex === 0) {
            tableRows.unshift(buildTableHeadingRow(section.title, replica, design));
            hasRenderedHeadingRow = true;
          }

          children.push(new Table({
            width: { size: 9360, type: WidthType.DXA },
            columnWidths: replica ? [2300, 7060] : [3100, 6260],
            layout: TableLayoutType.AUTOFIT,
            rows: tableRows,
          }));
        }

        if (hasGroupConditions) {
          const conditionalGroupEnd = conditionalSectionEndParagraph({ ...section, displayWhen: group.conditions }, fieldsByLabel, { size: 20, color: design.primaryColor, font: design.fontFamily });
          if (conditionalGroupEnd) children.push(conditionalGroupEnd);
        }
      }
      children.push(new Paragraph({ text: "", spacing: { after: replica ? 100 : 180 } }));
    } else {
      for (const line of String(section.content || "").split("\n")) {
        children.push(new Paragraph({
          children: parseContentWithFields(line, { font: design.fontFamily, size: Math.max(16, Math.round(design.bodyFontSize * 2)), color: design.primaryColor }),
          spacing: { after: 120 },
          alignment: section.isStandardFooter ? AlignmentType.RIGHT : undefined,
        }));
      }
    }

    const conditionalEnd = conditionalSectionEndParagraph(section, fieldsByLabel, { size: 20, color: design.primaryColor, font: design.fontFamily });
    if (conditionalEnd) children.push(conditionalEnd);
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: design.fontFamily, size: Math.max(16, Math.round(design.bodyFontSize * 2)), color: design.primaryColor },
        },
      },
    },
    sections: [{
      properties: { page: { margin: { top: replica ? 420 : design.pageMargin, right: design.pageMargin, bottom: replica ? 900 : design.pageMargin, left: design.pageMargin } } },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: replica ? [] : [new TextRun({ text: design.headerText || normalized.title, bold: true, size: 20, color: design.headerColor, font: design.fontFamily })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: replica
              ? [
                new TextRun({ text: design.footerText, font: design.fontFamily, size: 16, color: design.footerColor }),
                new TextRun({ text: "\t\t\t\t\t\t\t\t\tPage", font: design.fontFamily, size: 16, color: design.footerColor }),
              ]
              : [new TextRun({ text: `${design.footerText} · Page `, color: design.footerColor, font: design.fontFamily }), new SimpleField("PAGE")],
            alignment: replica ? AlignmentType.LEFT : AlignmentType.CENTER,
          })],
        }),
      },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}

async function main() {
  const [, , inputPath, outputPath] = process.argv;
  if (!inputPath || !outputPath) {
    console.error("Usage: node scripts/generate-sm8f.mjs spec.json output.sm8f");
    process.exit(1);
  }

  const spec = JSON.parse(await fs.readFile(inputPath, "utf8"));
  const formJson = buildFormJson(spec);
  const docxBuffer = await buildDocx(spec);

  const zip = new JSZip();
  zip.file("form.json", JSON.stringify(formJson));
  zip.file("template.docx", docxBuffer);

  const outputBuffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  await fs.writeFile(outputPath, outputBuffer);
  console.log(`Wrote ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
