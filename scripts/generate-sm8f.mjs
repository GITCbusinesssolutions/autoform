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

function slug(value) {
  return cleanLabel(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function optionCode(value) {
  return cleanLabel(value)
    .replace(/[^a-zA-Z0-9]+/g, "_")
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
  return !!field &&
    formFieldType(field).startsWith("Multiple Choice") &&
    String(docxOutputType(field)).toLowerCase() === "checkboxlist";
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
  return {
    title,
    badgeName: badgeName(spec.badgeName || spec.badge_name, title),
    description: String(spec.description || ""),
    badgeMandatoryState: String(spec.badgeMandatoryState || spec.badge_mandatory_state || "2"),
    canBeUsedIndependently: String(spec.canBeUsedIndependently || spec.can_be_used_independently || "0"),
    templateFields: Array.isArray(spec.templateFields) ? spec.templateFields : [],
    fields: fields.map((field, index) => ({
      ...field,
      uuid: field.uuid || crypto.randomUUID(),
      label: cleanLabel(field.label || field.name || `Field ${index + 1}`),
      sortOrder: String(field.sortOrder || field.sort_order || index + 1),
      required: !!field.required,
      additionalDetails: String(field.additionalDetails || ""),
      options: Array.isArray(field.options) ? field.options.map((option) => String(option || "").trim()).filter(Boolean) : undefined,
    })),
    docxContent: spec.docxContent || {},
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
    return `${field.label}: ${mergeField(mergeFieldName(field))}`;
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

function buildTableRows(rows, section, fieldsByLabel, replica = false) {
  const labelWidth = replica ? 2300 : 3100;
  const valueWidth = replica ? 7060 : 6260;
  const labelRun = replica ? REPLICA_LABEL_RUN : { color: "334155", size: 20 };
  const valueRun = replica ? REPLICA_VALUE_RUN : TABLE_VALUE_RUN;
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
          margins: cellMargins,
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({
          children: tableValueParagraphs(label.trim(), value, fieldsByLabel, valueRun),
          width: { size: valueWidth, type: WidthType.DXA },
          margins: cellMargins,
          verticalAlign: VerticalAlign.CENTER,
        }),
      ],
    });
  });
}

function buildTableHeadingRow(title, replica = false) {
  return new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({
            text: title,
            bold: true,
            color: replica ? "FFFFFF" : "334155",
            size: replica ? 20 : 24,
            font: replica ? "Arial" : undefined,
          })],
        })],
        columnSpan: 2,
        shading: { fill: replica ? "3F3F3F" : "E2E8F0" },
        margins: replica
          ? { top: 75, bottom: 75, left: 90, right: 90 }
          : { top: 120, bottom: 120, left: 120, right: 120 },
      }),
    ],
  });
}

async function replicaIntro(spec) {
  const children = [];
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
              transformation: { width: 150, height: 69 },
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
        new TextRun({ text: "AMBROSE ", bold: true, color: "6D6E71", size: 21, font: "Arial" }),
        new TextRun({ text: "CONSTRUCT", bold: true, color: "A8D51E", size: 21, font: "Arial" }),
      ],
      spacing: { after: 0 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "GROUP", bold: true, color: "8A8C90", size: 13, font: "Arial" })],
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
              children: [new TextRun({ text: "Ambrose Construct Group Pty Ltd", bold: true, font: "Arial", size: 20, color: "404040" })],
              alignment: AlignmentType.RIGHT,
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "www.ambrose", font: "Arial", size: 14, color: "404040" }),
                new TextRun({ text: "construct", font: "Arial", size: 14, color: "A8D51E", bold: true }),
                new TextRun({ text: ".com.au", font: "Arial", size: 14, color: "404040" }),
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
          children: [new TextRun({ text: spec.title, bold: true, color: "FFFFFF", font: "Arial", size: 30 })],
        })],
        shading: { fill: "555555" },
        margins: { top: 140, bottom: 140, left: 120, right: 120 },
      })],
    })],
  }));

  children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
  return children;
}

async function buildDocx(spec) {
  const normalized = normalizeSpec(spec);
  const replica = isReplicaSpec(normalized);
  const sections = Array.isArray(normalized.docxContent.sections) && normalized.docxContent.sections.length
    ? normalized.docxContent.sections
    : autoSections(normalized);
  const fieldsByLabel = fieldLookup(normalized.fields);

  const children = replica
    ? await replicaIntro(normalized)
    : [
      new Paragraph({
        children: [new TextRun({ text: normalized.title, bold: true, size: 34, color: "0F172A" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 160 },
      }),
      new Paragraph({
        children: [new TextRun({ text: normalized.description, italics: true, color: "64748B", size: 22 })],
        spacing: { after: 360 },
      }),
    ];

  for (const section of sections) {
    const isTableSection = section.layout === "table" || section.isStandardHeader;
    const hasSectionConditions = displayConditions(section).length > 0;
    const useTableHeading = isTableSection && !section.isStandardHeader && !section.isStandardFooter && !hasSectionConditions;

    if (replica && section.isStandardHeader) {
      children.push(new Paragraph({
        children: [new TextRun({ text: String(section.title || "").toUpperCase(), bold: true, font: "Arial", size: 22, color: "3F3F3F" })],
        keepNext: true,
        spacing: { before: 90, after: 55 },
        border: { bottom: { color: "BDBDBD", space: 1, style: BorderStyle.SINGLE, size: 10 } },
      }));
    }

    if (!section.isStandardHeader && !section.isStandardFooter && !useTableHeading) {
      const headingRunOptions = { bold: true, size: 26, color: "334155" };
      const headingChildren = hasSectionConditions
        ? conditionalSectionStartRuns(section, fieldsByLabel, headingRunOptions)
        : [new TextRun({ text: section.title, ...headingRunOptions })];

      children.push(new Paragraph({
        children: headingChildren,
        keepNext: true,
        spacing: { before: replica ? 180 : 280, after: replica ? 80 : 120 },
        border: { bottom: { color: replica ? "BDBDBD" : "CBD5E1", space: 1, style: BorderStyle.SINGLE, size: replica ? 10 : 6 } },
      }));
    }

    if (isTableSection) {
      const rows = String(section.content || "").split("\n").filter((line) => line.trim());
      const chunks = chunkTableRows(rows, section, fieldsByLabel);
      for (const [chunkIndex, chunk] of chunks.entries()) {
        const tableRows = buildTableRows(chunk, section, fieldsByLabel, replica);
        if (useTableHeading && chunkIndex === 0) tableRows.unshift(buildTableHeadingRow(section.title, replica));

        children.push(new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: replica ? [2300, 7060] : [3100, 6260],
          layout: TableLayoutType.AUTOFIT,
          rows: tableRows,
        }));
      }
      children.push(new Paragraph({ text: "", spacing: { after: replica ? 100 : 180 } }));
    } else {
      for (const line of String(section.content || "").split("\n")) {
        children.push(new Paragraph({
          children: parseContentWithFields(line),
          spacing: { after: 120 },
          alignment: section.isStandardFooter ? AlignmentType.RIGHT : undefined,
        }));
      }
    }

    const conditionalEnd = conditionalSectionEndParagraph(section, fieldsByLabel, { size: 20, color: "334155" });
    if (conditionalEnd) children.push(conditionalEnd);
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22, color: "1F2937" },
        },
      },
    },
    sections: [{
      properties: { page: { margin: replica ? { top: 420, right: 720, bottom: 900, left: 720 } : { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: replica ? [] : [new TextRun({ text: normalized.title, bold: true, size: 20, color: "475569" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: replica
              ? [
                new TextRun({ text: "Commercial-in-confidence", font: "Arial", size: 16, color: "000000" }),
                new TextRun({ text: "\t\t\t\t\t\t\t\t\tPage", font: "Arial", size: 16, color: "000000" }),
              ]
              : [new TextRun({ text: "Page ", color: "64748B" }), new SimpleField("PAGE")],
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
