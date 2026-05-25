#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import crypto from "node:crypto";
import JSZip from "jszip";

const DEFAULT_CONDITIONS = [
  { question: "", operator: "", value: "" },
  { question: "", operator: "", value: "" },
  { question: "", operator: "", value: "" },
];

function nowStamp() {
  const date = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function cleanLabel(value, fallback = "SWMS") {
  return String(value || fallback)
    .replace(/[^a-zA-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim() || fallback;
}

function optionCode(value) {
  return cleanLabel(value)
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function swmsOptionMergeField(title) {
  return `form_swms_required_${optionCode(title)}`;
}

function escapeXml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function ifStartParagraph(fieldName) {
  const safeFieldName = escapeXml(fieldName);
  return `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText xml:space="preserve"> IF "</w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText xml:space="preserve"> MERGEFIELD ${safeFieldName} \\* MERGEFORMAT </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:noProof/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText>«${safeFieldName.slice(0, 40)}»</w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText>" = "Yes" "</w:instrText></w:r></w:p>`;
}

function ifEndParagraph() {
  return `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText>" "" </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t> </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
}

function documentBodyParts(documentXml) {
  const bodyMatch = documentXml.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);
  if (!bodyMatch) throw new Error("DOCX is missing word/document.xml body");
  const bodyContent = bodyMatch[1];
  const sectMatch = bodyContent.match(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/);
  return {
    content: sectMatch ? bodyContent.slice(0, sectMatch.index) : bodyContent,
    sectPr: sectMatch ? sectMatch[0] : "",
  };
}

function documentXmlFromParts(originalXml, bodyContent, sectPr) {
  return originalXml.replace(/<w:body[^>]*>[\s\S]*?<\/w:body>/, `<w:body>${bodyContent}${sectPr}</w:body>`);
}

function parseRelationships(relsXml = "") {
  const relationships = [];
  const regex = /<Relationship\b([^>]*?)\/>/g;
  let match;
  while ((match = regex.exec(relsXml)) !== null) {
    const attrs = {};
    for (const attr of match[1].matchAll(/\s([A-Za-z:]+)="([^"]*)"/g)) {
      attrs[attr[1]] = attr[2];
    }
    if (attrs.Id) relationships.push(attrs);
  }
  return relationships;
}

function relationshipXml(relationships) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((rel) => `<Relationship Id="${escapeXml(rel.Id)}" Type="${escapeXml(rel.Type)}" Target="${escapeXml(rel.Target)}"${rel.TargetMode ? ` TargetMode="${escapeXml(rel.TargetMode)}"` : ""}/>`).join("")}</Relationships>`;
}

function nextRid(existing) {
  let max = 0;
  for (const rel of existing) {
    const match = String(rel.Id || "").match(/^rId(\d+)$/);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return () => `rId${++max}`;
}

function contentTypeOverrideXml(partName, contentType) {
  return `<Override PartName="/${escapeXml(partName)}" ContentType="${escapeXml(contentType)}"/>`;
}

function contentTypeDefaultXml(extension, contentType) {
  return `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/>`;
}

function ensureContentType(contentTypesXml, partName, contentType) {
  if (!contentType || contentTypesXml.includes(`PartName="/${partName}"`)) return contentTypesXml;
  return contentTypesXml.replace("</Types>", `${contentTypeOverrideXml(partName, contentType)}</Types>`);
}

function ensureDefaultContentType(contentTypesXml, extension, contentType) {
  if (!extension || !contentType || contentTypesXml.includes(`Extension="${extension}"`)) return contentTypesXml;
  return contentTypesXml.replace("</Types>", `${contentTypeDefaultXml(extension, contentType)}</Types>`);
}

function contentTypeForPart(contentTypesXml, partName) {
  const escaped = partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const override = contentTypesXml.match(new RegExp(`<Override[^>]+PartName="/${escaped}"[^>]+ContentType="([^"]+)"`));
  if (override) return override[1];
  const extension = path.extname(partName).replace(".", "");
  const def = contentTypesXml.match(new RegExp(`<Default[^>]+Extension="${extension}"[^>]+ContentType="([^"]+)"`));
  return def?.[1] || "";
}

function replaceAll(value, replacements) {
  let next = value;
  for (const [from, to] of replacements) {
    next = next.split(from).join(to);
  }
  return next;
}

async function copyRelatedParts({ sourceZip, sourceRels, targetZip, targetRels, contentTypesXml, docIndex, bodyXml }) {
  const allocateRid = nextRid(targetRels);
  const ridReplacements = [];
  let nextContentTypes = contentTypesXml;

  for (const rel of sourceRels) {
    if (!bodyXml.includes(rel.Id)) continue;
    const target = rel.Target || "";
    if (rel.TargetMode === "External") {
      const newId = allocateRid();
      targetRels.push({ ...rel, Id: newId });
      ridReplacements.push([rel.Id, newId]);
      continue;
    }

    if (target.startsWith("../")) continue;
    const sourcePart = path.posix.normalize(`word/${target}`);
    const sourceFile = sourceZip.file(sourcePart);
    if (!sourceFile) continue;

    const ext = path.posix.extname(sourcePart);
    const base = path.posix.basename(sourcePart, ext);
    const dir = path.posix.dirname(sourcePart);
    const targetPart = `${dir}/swms${docIndex}_${base}${ext}`;
    const newTarget = path.posix.relative("word", targetPart);
    const newId = allocateRid();

    targetZip.file(targetPart, await sourceFile.async("nodebuffer"));
    targetRels.push({ ...rel, Id: newId, Target: newTarget });
    ridReplacements.push([rel.Id, newId]);

    const contentType = contentTypeForPart(contentTypesXml, sourcePart);
    if (contentType) {
      if (sourcePart.startsWith("word/media/")) {
        nextContentTypes = ensureDefaultContentType(nextContentTypes, ext.replace(".", ""), contentType);
      } else {
        nextContentTypes = ensureContentType(nextContentTypes, targetPart, contentType);
      }
    }
  }

  return { bodyXml: replaceAll(bodyXml, ridReplacements), contentTypesXml: nextContentTypes };
}

async function buildTemplateDocx(documents) {
  if (!documents.length) throw new Error("Select at least one SWMS document");

  const firstZip = await JSZip.loadAsync(await fs.readFile(documents[0].path));
  const targetZip = await JSZip.loadAsync(await fs.readFile(documents[0].path));
  const firstDocumentXml = await firstZip.file("word/document.xml").async("string");
  const firstParts = documentBodyParts(firstDocumentXml);
  const relsPath = "word/_rels/document.xml.rels";
  const firstRelsXml = await targetZip.file(relsPath).async("string");
  const targetRels = parseRelationships(firstRelsXml);
  let contentTypesXml = await targetZip.file("[Content_Types].xml").async("string");
  const blocks = [];

  for (const [index, document] of documents.entries()) {
    const sourceZip = await JSZip.loadAsync(await fs.readFile(document.path));
    const sourceDocumentXml = await sourceZip.file("word/document.xml").async("string");
    const sourceRelsXml = sourceZip.file(relsPath) ? await sourceZip.file(relsPath).async("string") : "";
    const sourceRels = parseRelationships(sourceRelsXml);
    const sourceParts = documentBodyParts(sourceDocumentXml);
    const copied = await copyRelatedParts({
      sourceZip,
      sourceRels,
      targetZip,
      targetRels,
      contentTypesXml,
      docIndex: index + 1,
      bodyXml: sourceParts.content,
    });
    contentTypesXml = copied.contentTypesXml;
    blocks.push(`${ifStartParagraph(swmsOptionMergeField(document.title))}${copied.bodyXml}${ifEndParagraph()}`);
  }

  targetZip.file(relsPath, relationshipXml(targetRels));
  targetZip.file("[Content_Types].xml", contentTypesXml);
  targetZip.file("word/document.xml", documentXmlFromParts(firstDocumentXml, blocks.join(""), firstParts.sectPr));
  return targetZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

function updateConditions(conditions, uuidMap) {
  const items = Array.isArray(conditions) ? conditions : DEFAULT_CONDITIONS;
  const mapped = items.slice(0, 3).map((condition) => ({
    ...condition,
    question: condition.question && uuidMap.has(condition.question)
      ? uuidMap.get(condition.question)
      : condition.question || "",
  }));
  while (mapped.length < 3) mapped.push({ question: "", operator: "", value: "" });
  return mapped;
}

function buildFormJson(baseForm, title, swmsTitles) {
  const formUuid = crypto.randomUUID();
  const documentTemplateUuid = crypto.randomUUID();
  const timestamp = nowStamp();
  const uuidMap = new Map();
  const sortedBaseFields = [...(baseForm.fields || [])].sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
  for (const field of sortedBaseFields) uuidMap.set(field.uuid, crypto.randomUUID());

  const fields = sortedBaseFields.map((field) => {
    const fieldData = JSON.parse(field.field_data_json || "{}");
    const nextData = {
      ...fieldData,
      conditions: updateConditions(fieldData.conditions, uuidMap),
    };

    if (field.name === "SWMS Required") {
      nextData.choices = swmsTitles;
      nextData.fieldType = "Multiple Choice (Multi-Answer)";
    }

    return {
      ...field,
      uuid: uuidMap.get(field.uuid),
      form_uuid: formUuid,
      create_date: timestamp,
      edit_date: timestamp,
      field_data_json: JSON.stringify(nextData),
    };
  });

  return {
    form: {
      ...baseForm.form,
      uuid: formUuid,
      create_date: timestamp,
      edit_date: timestamp,
      document_template_uuid: documentTemplateUuid,
      name: String(title || "SWMS").trim() || "SWMS",
      badge_name: "SWMS",
      network_origin_form_uuid: "",
      network_origin_form_etag: "",
      store_item_uuid: "",
      is_locked: "0",
    },
    fields,
  };
}

async function main() {
  const [inputPath, outputPath] = process.argv.slice(2);
  if (!inputPath || !outputPath) {
    console.error("Usage: node scripts/build-swms-sm8f.mjs swms-build.json output.sm8f");
    process.exit(1);
  }

  const input = JSON.parse(await fs.readFile(inputPath, "utf8"));
  const documents = Array.isArray(input.documents) ? input.documents : [];
  if (!documents.length) throw new Error("No SWMS documents selected");

  const baseFormPath = input.baseFormPath || path.join(process.cwd(), "assets", "swms-base-form.json");
  const baseForm = JSON.parse(await fs.readFile(baseFormPath, "utf8"));
  const formJson = buildFormJson(baseForm, input.title || "Electrical & Solar Installation SWMS", documents.map((document) => document.title));
  const templateDocx = await buildTemplateDocx(documents);

  const zip = new JSZip();
  zip.file("form.json", JSON.stringify(formJson));
  zip.file("template.docx", templateDocx);
  await fs.writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log(`Wrote ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
