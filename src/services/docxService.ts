import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel, 
  AlignmentType, 
  SimpleField, 
  Header, 
  Footer, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  BorderStyle, 
  VerticalAlign,
  HeightRule,
  XmlComponent,
  XmlAttributeComponent
} from "docx";
import { saveAs } from "file-saver";
import { GeneratedForm, AppSettings } from "../types";

class FieldCharAttributes extends XmlAttributeComponent<any> {
  constructor(type: string) {
    super({ type });
    // @ts-ignore
    this.xmlKeys = { type: "w:fldCharType" };
  }
}

class FieldChar extends XmlComponent {
  constructor(type: "begin" | "separate" | "end") {
    super("w:fldChar");
    this.root.push(new FieldCharAttributes(type));
  }
}

function parseContentWithMergeFields(text: string): (TextRun | SimpleField)[] {
  const parts: (TextRun | SimpleField)[] = [];
  
  // Matches the outermost { ... } blocks, supporting one level of nesting
  const regex = /\{((?:[^{}]|\{[^{}]*\})+)\}/g;
  let lastIndex = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      const plainText = text.substring(lastIndex, match.index);
      if (plainText) {
        parts.push(new TextRun({ text: plainText }));
      }
    }
    
    const fieldContent = match[1];
    
    if (fieldContent.startsWith('IF')) {
      // Complex IF field with potential nested MERGEFIELD
      // We use TextRun with FieldChar for complex fields
      parts.push(new TextRun({ children: [new FieldChar("begin")] }) as any);
      
      // Parse the IF instruction, looking for nested {MERGEFIELD ...}
      const innerRegex = /\{MERGEFIELD\s+(.*?)\s+\\\*\s+MERGEFORMAT\}/g;
      let innerLastIndex = 0;
      let innerMatch;
      
      while ((innerMatch = innerRegex.exec(fieldContent)) !== null) {
        if (innerMatch.index > innerLastIndex) {
          parts.push(new TextRun({ text: fieldContent.substring(innerLastIndex, innerMatch.index) }) as any);
        }
        
        // Nested MERGEFIELD
        parts.push(new TextRun({ children: [new FieldChar("begin")] }) as any);
        parts.push(new TextRun({ text: ` MERGEFIELD ${innerMatch[1]} \\* MERGEFORMAT ` }) as any);
        parts.push(new TextRun({ children: [new FieldChar("separate")] }) as any);
        parts.push(new TextRun({ text: `«${innerMatch[1]}»` }) as any);
        parts.push(new TextRun({ children: [new FieldChar("end")] }) as any);
        
        innerLastIndex = innerRegex.lastIndex;
      }
      
      if (innerLastIndex < fieldContent.length) {
        parts.push(new TextRun({ text: fieldContent.substring(innerLastIndex) }) as any);
      }
      
      parts.push(new TextRun({ children: [new FieldChar("separate")] }) as any);
      parts.push(new TextRun({ text: " " }) as any);
      parts.push(new TextRun({ children: [new FieldChar("end")] }) as any);
    } else if (fieldContent.startsWith('MERGEFIELD')) {
      // Standard merge field
      parts.push(new SimpleField(fieldContent));
    } else {
      // Fallback for other fields
      parts.push(new SimpleField(fieldContent));
    }
    
    lastIndex = regex.lastIndex;
  }

  if (lastIndex < text.length) {
    const remainingText = text.substring(lastIndex);
    if (remainingText) {
      parts.push(new TextRun({ text: remainingText }));
    }
  }

  return parts;
}

export async function downloadDocx(form: GeneratedForm, settings: AppSettings) {
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Calibri",
            size: 22,
            color: "333333",
          },
        },
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            size: 36,
            bold: true,
            color: "1e293b",
            font: "Inter",
          },
          paragraph: {
            spacing: { before: 240, after: 120 },
            alignment: AlignmentType.CENTER,
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            size: 28,
            bold: true,
            color: "334155",
            font: "Inter",
          },
          paragraph: {
            spacing: { before: 240, after: 120 },
            border: {
              bottom: {
                color: "CBD5E1",
                space: 1,
                style: BorderStyle.SINGLE,
                size: 6,
              },
            },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: settings.companyName || "FormAI Builder",
                    bold: true,
                    size: 24,
                    color: "4f46e5",
                  }),
                ],
                alignment: AlignmentType.LEFT,
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: settings.headerText || "Generated ServiceM8 Template",
                    size: 18,
                    color: "64748b",
                  }),
                ],
                alignment: AlignmentType.LEFT,
                spacing: { after: 200 },
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: settings.footerText || "Page ",
                    size: 18,
                    color: "94a3b8",
                  }),
                  new SimpleField("PAGE"),
                  new TextRun({
                    text: " of ",
                    size: 18,
                    color: "94a3b8",
                  }),
                  new SimpleField("NUMPAGES"),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        },
        children: [
          new Paragraph({
            text: form.title,
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: form.description,
                italics: true,
                color: "64748b",
              }),
            ],
            spacing: { after: 400 },
          }),

          // Render AI-generated sections
          ...form.docxContent.sections.flatMap((section) => {
            const sectionElements: (Paragraph | Table)[] = [];
            
            // Add Section Title as Heading 2 (unless it's a standard header/footer)
            if (!section.isStandardHeader && !section.isStandardFooter) {
              sectionElements.push(new Paragraph({
                text: section.title,
                heading: HeadingLevel.HEADING_2,
              }));
            }

            if (section.layout === 'table' || section.isStandardHeader) {
              // Standard headers often look good as tables
              const lines = section.content.split('\n').filter(l => l.trim());
              sectionElements.push(new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: lines.map(line => {
                  const [label, ...rest] = line.split(':');
                  const value = rest.join(':').trim();
                  
                  return new TableRow({
                    children: [
                      new TableCell({
                        children: [new Paragraph({ 
                          children: [new TextRun({ 
                            text: label.trim(), 
                            bold: true,
                            size: section.isStandardHeader ? 18 : 22,
                            color: section.isStandardHeader ? "64748b" : "333333"
                          })] 
                        })],
                        width: { size: 35, type: WidthType.PERCENTAGE },
                        shading: { fill: section.isStandardHeader ? "F1F5F9" : "F8FAFC" },
                        margins: { top: 80, bottom: 80, left: 100, right: 100 },
                        borders: section.isStandardHeader ? {
                          top: { style: BorderStyle.NONE },
                          bottom: { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" },
                          left: { style: BorderStyle.NONE },
                          right: { style: BorderStyle.NONE },
                        } : undefined,
                      }),
                      new TableCell({
                        children: [new Paragraph({ 
                          children: parseContentWithMergeFields(value || line),
                          alignment: section.isStandardHeader ? AlignmentType.LEFT : undefined
                        })],
                        width: { size: 65, type: WidthType.PERCENTAGE },
                        margins: { top: 80, bottom: 80, left: 100, right: 100 },
                        borders: section.isStandardHeader ? {
                          top: { style: BorderStyle.NONE },
                          bottom: { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" },
                          left: { style: BorderStyle.NONE },
                          right: { style: BorderStyle.NONE },
                        } : undefined,
                      }),
                    ],
                  });
                }),
              }));
            } else {
              // Default text layout
              const lines = section.content.split('\n');
              lines.forEach(line => {
                sectionElements.push(new Paragraph({
                  children: parseContentWithMergeFields(line),
                  spacing: { after: 120 },
                  alignment: section.isStandardFooter ? AlignmentType.RIGHT : undefined,
                }));
              });
            }

            // Add some spacing after standard header
            if (section.isStandardHeader) {
              sectionElements.push(new Paragraph({ text: "", spacing: { after: 240 } }));
            }

            return sectionElements;
          }),

          // Optional: Summary table of all form fields if requested
          ...(settings.showTable ? [
            new Paragraph({
              text: "Form Field Summary",
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 400 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  tableHeader: true,
                  children: [
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: "Question / Label", bold: true, color: "FFFFFF" })] })],
                      shading: { fill: "4f46e5" },
                      verticalAlign: VerticalAlign.CENTER,
                      margins: { top: 100, bottom: 100, left: 100, right: 100 },
                    }),
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: "Merge Field", bold: true, color: "FFFFFF" })] })],
                      shading: { fill: "4f46e5" },
                      verticalAlign: VerticalAlign.CENTER,
                      margins: { top: 100, bottom: 100, left: 100, right: 100 },
                    }),
                  ],
                }),
                ...form.fields.map((field) => {
                  const sanitizedLabel = field.label
                    .toLowerCase()
                    .replace(/[^a-z0-9]/g, "_") // Replace all non-alphanumeric with underscore
                    .replace(/_+/g, "_")        // Collapse multiple underscores
                    .replace(/^_|_$/g, "");     // Trim leading/trailing underscores

                  const fieldName = field.type === 'photo' 
                    ? `image_form_${sanitizedLabel}_medium`
                    : `form_${sanitizedLabel}`;

                  return new TableRow({
                    children: [
                      new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: field.label, bold: true })] })],
                        verticalAlign: VerticalAlign.CENTER,
                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                      }),
                      new TableCell({
                        children: [new Paragraph({ children: [new SimpleField(`MERGEFIELD ${fieldName} \\* MERGEFORMAT`)] })],
                        verticalAlign: VerticalAlign.CENTER,
                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                      }),
                    ],
                  });
                }),
              ],
            })
          ] : []),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${form.title.replace(/\s+/g, "_")}_template.docx`);
}
