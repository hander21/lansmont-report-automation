# Lansmont Report Automation Rules

This project automates packaging test report drafts from Lansmont-style exported files.

This project may eventually process NDA-covered packaging test data, customer product photos,
customer names, and internal company information.

## Confidentiality Rules (MANDATORY)

Claude must only work with synthetic, fake, or redacted sample files.

Do NOT request:
- Real customer data or customer names
- Real product photos
- Real Lansmont exports or test results
- Confidential report templates
- Internal company information

The generated Python tool must run locally on company-approved machines and process real files
without sending them to any AI model.

The AI is used ONLY to help build and debug the code using fake data. It is NOT used at runtime
to read, summarize, classify, or interpret confidential test files.

## Hard Rules

- Never invent test values
- Never estimate missing values
- Never calculate official test results with AI
- All official report values must come from parsed local source files
- Every value inserted into the report must be logged in manifest.json
- If a required field is missing: stop and show a clear error
- If a required chart is missing: stop and show a clear error
- If required photos are missing: stop and show a warning or error depending on config
- Insert Lansmont-generated charts exactly as exported — do not alter them
- Insert pre-test and post-test photos from local files — do not alter them
- Output reports must be marked DRAFT until reviewed by a human
- Do not delete raw data files
- Preserve original exports and photos
- Prioritize reliability and traceability over flashiness

## Runtime Behavior

- The tool is deterministic: same inputs always produce the same output
- The tool never calls external APIs or sends data over the network
- The tool never uses AI models at runtime
- The tool never emails or distributes reports automatically

## Sample Data Policy

All files under `samples/` are fake/synthetic data only.
- Customer names: Customer_A, Customer_B, etc.
- Product names: Demo_Product, Sample_Package, etc.
- Project numbers: SAMPLE-123, SAMPLE-456, etc.
- Test values: clearly fake numeric samples
- Photos: placeholder images only

## Development Workflow

See README.md for setup and usage instructions.
