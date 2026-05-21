<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://ai.google.dev/static/site-assets/images/share-ais-513315318.png" />
</div>

# ServiceM8 Form & DOCX Generator

Local tooling for turning a Codex-authored form spec into an importable ServiceM8 `.sm8f` package.

View your app in AI Studio: https://ai.studio/apps/6de47f81-476c-4307-8826-27b64f6b7c09

## Run Locally

**Prerequisites:**  Node.js

1. Install dependencies:
   `npm install`
2. Copy `.env.example` to `.env.local`
3. If port 3000 is busy, set `PORT` and `APP_URL` in `.env.local`, for example:
   `PORT=3456`
   `APP_URL=http://localhost:3456`
4. Run the app:
   `npm run dev`

## Current Workflow

- Give Codex the scope/files in chat.
- Codex drafts the form plan, question list, test data, and preview output.
- After revisions and approval, Codex writes or updates a spec JSON.
- Run the `.sm8f` generator to package `form.json` and `template.docx`.

## Slack + Codex Workflow

If this repo is connected to Codex from Slack, mention Codex in a Slack thread with the scope and source files. Codex should follow [AGENTS.md](./AGENTS.md):

- inspect the uploaded files
- reply in-thread with a proposed form plan and question list
- ask clarification questions before building
- iterate from the user's replies
- generate the `.sm8f` only after approval

For reliable results, include the target form name, required fields, conditional logic requirements, preferred report/PDF behavior, and any existing ServiceM8/job/custom fields that should be reused.

## Direct `.sm8f` Generation

Codex can also generate importable ServiceM8 form packages directly. An `.sm8f` file is a ZIP archive containing:

- `form.json`
- `template.docx`

Generate the included sample:

`npm run sm8f:sample`

Generate from a custom spec:

`node scripts/generate-sm8f.mjs path/to/spec.json output/my-form.sm8f`

## Conditional Report Output

When Codex builds a spec from a brief or source file, it should confirm two report-output choices before final packaging:

- For each `Multiple Choice (Multi-Answer)` field, ask whether the DOCX should use the raw comma-separated merge output or dynamic checklist rows.
- For each conditional section, ask whether the PDF should hide the skipped section or show a blank table.

In a spec, set `"docxOutput": "checkboxList"` on a multi-answer field to generate per-option Word IF checkboxes. Add `displayWhen` to a `docxContent.sections[]` entry to wrap that section/table in a Word IF block.
