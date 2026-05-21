# Codex Instructions: ServiceM8 Form Generator

This repository builds importable ServiceM8 `.sm8f` form packages from user-provided scopes, PDFs, DOCX examples, and existing `.sm8f` samples.

## Primary Goal

When a user requests a form through Slack, GitHub, or Codex chat, Codex should work conversationally first, then generate the asset only after the user approves the plan.

The final deliverable is usually:

- `form.json`
- `template.docx`
- packaged together as an importable `.sm8f` ZIP archive with the `.sm8f` extension

Do not use the live ServiceM8 API unless the user explicitly asks. Prefer `.sm8f` export generation.

## Slack Intake Workflow

When a user mentions Codex from Slack with a scope, PDF, DOCX, image, or `.sm8f` sample:

1. Reply in the Slack thread acknowledging the request and listing the files/scope received.
2. Inspect all supplied source files before designing the form.
3. Extract a proposed form plan:
   - form title and badge name
   - sections in order
   - questions in order
   - field type for each question
   - required/optional status
   - conditional logic
   - expected DOCX/PDF report output behavior
4. Ask concise clarification questions for anything that materially affects the form, especially:
   - whether multi-answer fields should output as raw comma-separated text or checklist-style checked boxes
   - whether conditional report sections should be hidden from the generated PDF when skipped
   - whether any source PDF administrative fields should map to existing ServiceM8/job/custom fields instead of new form fields
   - which fields are mandatory
5. Post a natural-language summary of what will be built and wait for user confirmation.
6. Iterate in-thread when the user asks for changes.
7. Only after approval, generate or update the JSON spec and build the `.sm8f`.
8. Reply with the final `.sm8f` artifact path or uploaded asset, plus a short summary of what changed.

## ServiceM8 Conditional Logic Rules

ServiceM8 form conditions in `.sm8f` files are skip conditions: they describe when a dependent field is hidden.

For a single-choice/dropdown source condition:

- display rule: `Service Type = Other`
- ServiceM8 skip rule: `Service Type NEQ Other`

For a multi-choice multi-answer source condition:

- ServiceM8 stores selected values as comma-separated text, for example `Bathroom,Ensuite,Laundry`
- display rule: `Areas Tested contains Bathroom`
- ServiceM8 skip rule: `Areas Tested NCON Bathroom`

Never use `NEQ <Option>` when the source question is `Multiple Choice (Multi-Answer)`. Use `NCON <Option>` / "does not contain".

## DOCX Report Output Rules

Use editable Word tables where the user asks for a report/PDF replica. Avoid fixed-position layouts unless there is no practical table equivalent.

For option checkboxes in DOCX templates:

- use option-specific merge fields such as `form_site_close_out_confirmed_SWMS_Actioned`
- render one nested Word `IF` field per option
- show checked/unchecked symbols rather than raw `Yes` values when the user asks for checklist output

For conditional DOCX/PDF sections:

- ask whether skipped sections should be hidden or appear blank
- if hidden, wrap the relevant block/table in a Word `IF` field

## Build And Verify

Generate an `.sm8f` from a spec:

```bash
node scripts/generate-sm8f.mjs examples/leak-location-report.json output/leak-location-report-replica.sm8f
```

Run type checking:

```bash
npm run lint
```

When layout matters, unpack the `.sm8f`, render `template.docx`, and inspect the generated page images before delivery.
