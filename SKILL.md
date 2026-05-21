# Skill: PDF to ServiceM8 Form & DOCX Replication

Step-by-step technical blueprint and operational pipeline for parsing PDF forms and reproducing them as ServiceM8-compatible HTML electronic forms alongside a styling-replicated Microsoft Word (.docx) document template featuring nested conditional logic fields.

---

## 1. System Overview & Core Objectives

The goal is to automatically ingest a PDF form, analyze its questions, structure, and design, and generate two distinct assets:
1. **ServiceM8 HTML Form Fields Json Schema**: A list of structured fields (Text, Number, Date, Multiple Choice, Dropdown, Signature, Photo) to build the digital form inside ServiceM8.
2. **Dynamic DOCX Word Template (.docx)**: A highly-polished, styled Word layout replicating the original PDF's visual design. It uses native merge fields to bind directly to the ServiceM8 form data, employing advanced nested `IF` condition blocks for multiple-choice checkbox indicators (e.g., checking/unchecking boxes dynamically).

---

## 2. Step-by-Step Pipeline

```
[ Upload PDF ] ➔ [ Analyze Structure & Core Info ] ➔ [ Filter Global Merge Fields ] 
                       │
                       ▼
         ┌─────────────┴─────────────┐
         ▼                           ▼
[ Gen ServiceM8 Form JSON ]   [ Gen DOCX Structural Sections ]
(Clean Alphanumeric Names)    (Define headers, tables, layouts)
         │                           │
         ▼                           ▼
[ Push/Export ServiceM8 Form ] [ Compile Complex Nested IF Field DOCX ]
```

### Phase 1: PDF Question Extraction & Sanitization
When reviewing a PDF, the extraction engine must:
1. **Identify the Visual Layout**: Categorize the PDF structure. Decide if sections are dense grids (tables), simple linear flows, or split headers.
2. **Extract with Semantic Label Sanitization**:
   - ServiceM8 is highly restrictive of field names. For every question extracted, identify a **clean, short, alphanumeric label** (e.g., `RCDs Fitted` instead of `Are all safety switches (RCDs) fitted and working properly?`).
   - Move any complex, wordy description, or instruction text directly into the `additionalDetails` metadata property (which maps to the question's guidance text in the ServiceM8 UI).

### Phase 2: Deduplication via Global Merge Fields
To reduce redundant data entry for technicians in the field, compare all extracted questions against the list of ServiceM8 Global Merge Fields. 

*If a question asks for one of the following, do NOT include it as a question in the ServiceM8 form. Instead, use the global merge field directly in the DOCX layout:*

| Concept | Target Global Merge Field |
| :--- | :--- |
| Job Creation Date | `job.date` |
| Job Completed Date | `job.date_completed` |
| Client First Name | `job.contact_first` |
| Client Last Name | `job.contact_last` |
| Client Phone | `job.phone_1` |
| Client Mobile | `job.mobile` |
| Client Email | `job.email` |
| Job Address (Single Line) | `job.job_address_singleline` |
| Billing Address | `job.billing_address` |
| Current Status of Job | `job.status` |
| Job Description | `job.description` |
| Technician Full Name | `calculation.current_user_fullname` |
| Current Time (HH:MM) | `calculation.current_time_24` |
| Todays Date (Format change) | `calculation.todays_date` |

### Phase 3: ServiceM8 Form Attribute Constraints
Ensure compliance with strict validation limits imposed by the ServiceM8 API:
* **Badge Name Limit**: Every form must have a short badge name that is **strictly 11 characters or less** (e.g. `SAH Checklist` or `Elec Checklist`).
* **Field Character Cleaning**: Strip non-alphanumeric characters from standard form question name labels before pushing or building, replacing them with whitespace and then condensing spaces.

---

## 3. DOCX Template Synthesis & Custom Styling

To replicate the original PDF aesthetics, partition the generated document structure into standard semantic zones inside the DOCX content payload:

```json
{
  "title": "Safety Checklist",
  "badgeName": "SAH CHECK",
  "description": "Electrical safety assessment form",
  "docxContent": {
    "sections": [
      {
        "title": "Client Details",
        "isStandardHeader": true,
        "layout": "table",
        "content": "Date: {MERGEFIELD calculation.todays_date \\* MERGEFORMAT}\nAddress: {MERGEFIELD job.job_address_singleline \\* MERGEFORMAT}"
      },
      {
        "title": "Safety Switch Assessment",
        "isStandardHeader": false,
        "layout": "table",
        "content": "RCDs Fitted & Tested: {IF \"{MERGEFIELD form_rcds_fitted_yes \\* MERGEFORMAT}\"=\"Yes\" \"☑\" \"☐\"} Yes  {IF \"{MERGEFIELD form_rcds_fitted_no \\* MERGEFORMAT}\"=\"Yes\" \"☑\" \"☐\"} No"
      }
    ]
  }
}
```

### DOCX Section Layout Behaviors
* **isStandardHeader**: Typically holds administrative, job, or client data. Render this as a dense table grid with colored cell shading (e.g., `#F1F5F9`) and thin border-bottoms (`#E2E8F0`).
* **isStandardFooter**: Holds signatures, totals, or remarks. Render aligned to the right or bottom of the page.
* **layout: 'table'**: Use for inline grids of questions. The parser splits lines by the `\n` character, splits headers by `:`, and organizes keys and merge values into side-by-side table cells with a subtle zebra-shading background pattern (`#F8FAFC`).

---

## 4. STRICT Merge Field Formatting Rules

All merge fields MUST format precisely to ensure native binding:
1. **Standard Form Fields**: `form_lowercase_alphanumeric_only` (e.g., `{MERGEFIELD form_rcds_fitted \* MERGEFORMAT}`).
2. **Photo/Image Fields**: `image_form_lowercase_alphanumeric_only_medium` (e.g., `{MERGEFIELD image_form_site_arrival_before_photos_medium \* MERGEFORMAT}`).
3. **No Bad characters**: Labels or merge fields containing dashed formats, raw brackets, or punctuation will fail to bind.

---

## 5. Advanced Conditional Checkbox indicators (Nested Fields)

For checklists, checkboxes, or options, Word requires a **Nested Complex Field Object** structures to perform evaluation dynamically.

### Concept Output
```
{IF "form_field_option_yes"="Yes" "☑" "☐"}
```
Word must evaluate the merge value of `form_field_option_yes`. If it returns `"Yes"`, it renders the checked box `☑`, else it renders the unchecked box `☐`.

### Under-the-Hood XML Generation Pattern
A SimpleField (`w:fldSimple`) **cannot** be nested inside text. Instead, use Complex Fields via `w:fldChar` tags to build hierarchical trees.

The parser must dynamically compile an incoming raw string literal like:
```text
{IF "{MERGEFIELD form_has_safety_hazards_yes \\* MERGEFORMAT}"="Yes" "☑" "☐"}
```

Into this sequence of DOCX AST runs:
1. **Outer conditional field initiation**:
   - `Run` containing `FieldCharacter` with type `"begin"`.
   - `Run` with the conditional operation prefix instruction: `" IF "` and trailing quotes wrapping the nested field value context: `IF "`
2. **Inner Mergefield evaluation block**:
   - `Run` containing `FieldCharacter` with type `"begin"`.
   - `Run` with the instruction string containing the sanitized merge field target: `" MERGEFIELD form_has_safety_hazards_yes \\* MERGEFORMAT "`
   - `Run` containing `FieldCharacter` with type `"separate"`.
   - `Run` with visual placeholder fallback text: `"«form_has_safety_hazards_yes»"`
   - `Run` containing `FieldCharacter` with type `"end"`.
3. **Outer conditional field completion**:
   - `Run` with instruction close quote, comparison value, and conditional true/false values: `"\" = \"Yes\" \"☑\" \"☐\""`
   - `Run` containing `FieldCharacter` with type `"separate"`.
   - `Run` with blank separator visual output: `" "`
   - `Run` containing `FieldCharacter` with type `"end"`.

This XML structure guarantees that MS Word and the ServiceM8 conversion processor treat the conditional checkbox indicators as interactive, live field layers rather than literal text.

---

## 6. Supported Box Indicators & Symbol Library
The following standard symbols are supported inside conditional `IF` blocks and must be parsed out:
* **Unchecked Box**: `☐`
* **Checked Box**: `☑`
* **Checked Box (Bold / Emoji style)**: `✅`
* **Cross (X) Box**: `☒`
* **Standard Checkmark**: `✓` or `✔`
* **Standard Cross (X) Mark**: `✕` or `✖`
---

## 7. ServiceM8 Option Merge Fields and Conditional Report Blocks

For ServiceM8 option-specific merge fields, use:

```text
form_<lowercase_field_label>_<Choice_Text_With_Underscores>
```

The field label is lower-case/sanitized. The choice suffix preserves the original choice casing after punctuation is stripped. Example:

```text
form_site_close_out_confirmed_SWMS_Actioned
```

For multi-choice multi-answer fields, the option-specific merge field returns `Yes` when selected. Checklist report output must therefore be generated as one nested IF field per option:

```text
IF "{MERGEFIELD form_site_close_out_confirmed_SWMS_Actioned \* MERGEFORMAT}" = "Yes" "☑" "☐"
```

When building from a brief or PDF, ask the user whether each multi-answer question should output as:

* raw comma-separated ServiceM8 output, or
* a checklist with dynamic checked/unchecked boxes.

For conditional report sections, ask whether the section should be hidden from the PDF when skipped or still appear as a blank table. If hidden, add `displayWhen` to the DOCX section and generate a block-level Word IF wrapper around the title and table.

### ServiceM8 Multi-Answer Conditional Logic

ServiceM8 form conditions are skip conditions: a condition describes when the dependent field is hidden, not when it is shown.

For single-choice fields, invert a display rule like `Service Type = Other` to a skip rule of `NEQ Other`.

For multi-choice multi-answer fields, the selected answers are stored as comma-separated text such as:

```text
Bathroom,Ensuite,Laundry
```

Therefore, when a field or section should appear only if one option is selected, invert the display rule to "does not contain":

```json
{
  "question": "<areas_tested_uuid>",
  "operator": "NCON",
  "value": "Bathroom"
}
```

Do not use `NEQ Bathroom` for multi-answer targets. `NEQ` fails when multiple values are selected because the combined output is not exactly equal to a single option.
