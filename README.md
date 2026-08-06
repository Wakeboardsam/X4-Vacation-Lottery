# X4 Vacation Lottery - Module 1 Core Foundation

This repository contains Module 1 of the X4 Vacation Lottery web application built on Google Apps Script.

## Completed Behavior in Module 1

- **Participant PIN Authentication:** Securely authenticates users exactly matching their PIN, preserving leading zeroes, without returning sensitive configuration to the client. Uses bounded, 6-hour opaque session tokens stored in Script Properties.
- **Admin Authentication:** Simple, distinct admin login verifying against a salted SHA-256 hash.
- **Durable Workbook Schema Management:** An idempotent `initializeOrUpdateWorkbook` script that safely creates required sheets, injects headers without disrupting existing columns, seeds Config data, and applies cell validation (checkboxes).
- **Durable State Service:** Read/Write capability for "Config" variables that enforces structured keys, numeric checking, serialization, and recognized phase values, with safety locks.
- **Responsive Web Shell:** A single `Index.html` app separating Participant from Admin workflows. Admins can view state and run the initializer. Participants see basic active statuses.
- **Automated Test Harness:** Node-based testing using `node:test` mocking basic Apps Script environments to verify the business logic.

## Explicitly Out of Scope in Module 1

- Annual Auto-Fill or randomization
- Active-window or serpentine queue engine
- Vacation selection / assignments
- Rules acknowledgment workflow
- General audit history or transfer mechanics
- Notification mechanics (Twilio/SMS, Emails)
- Queue advancement or phase transitions

## Repository Structure

- `Code.gs`: Entry points (`doGet`, `onOpen`) and API wrappers.
- `Schema.gs`: Safe idempotent initialization logic.
- `State.gs`: Durable config reading, writing, and validation.
- `Auth.gs`: Session management, PIN and Admin verification.
- `Utils.gs`: Shared generic functions (Row lookups, Header mappings).
- `Index.html`: Responsive frontend application shell.
- `appsscript.json`: Manifest requiring only necessary scopes.
- `test/`: Node.js based unit tests utilizing native `node:test` and local Apps Script mocks.

## Local Test Setup

We use Node's native test runner to mock Apps Script globals. No large dependency stack is needed.

1. Ensure Node.js (v18+) is installed.
2. Initialize and run:
   ```bash
   npm test
   ```
All 32 tests spanning Schema idempotency, Authentication rules, State parsing, and Util boundaries should pass.

## Deployment & Live Setup

### 1. Creating the Apps Script Project

Create a new Google Apps Script project bound to your target Google Spreadsheet, or run as a standalone script connected to your spreadsheet.

### 2. Configure Administrator Credentials

To establish the admin hash securely (do this once in the Apps Script Editor, **not** through the client!):

1. Open `Code.gs` in the Editor.
2. Add a temporary execution block, or run this line via the editor's execution environment:
   ```javascript
   setAdminAccessCode('YOUR_SUPER_SECRET_CODE');
   ```
3. Remove or comment out that line immediately afterward. The script will hash and salt this code and save it purely in `PropertiesService.getScriptProperties()`.

### 3. Initialize the Spreadsheet

1. Open the connected Google Sheet.
2. Click `X4 Lottery` in the top custom menu -> `Initialize / Update Workbook`.
3. Wait for completion. It will build the initial structure, configuration, and data validations securely.

### 4. Setup Test Participants

1. Navigate to the `Turn Management` sheet.
2. Add test rows below the headers.
3. Example:
   - Name: `Alice Test`
   - PIN: `0442` (ensure it remains text format to preserve zeroes)
   - Active for Year: `TRUE`

### 5. Deploy Web App

1. Click **Deploy > New deployment**.
2. Select **Web app**.
3. Execute as: "Me".
4. Who has access: "Anyone" (Auth is handled internally).
5. Copy the web app URL.

## Manual Acceptance Testing

For thorough live execution, the following manual checks should be performed once deployed:

1. **Idempotency Check:** Open the deployed spreadsheet, run `Initialize / Update Workbook` once. Observe new sheets. Add color/data to a cell. Run it again. Ensure data and colors were strictly preserved.
2. **Auth Verification:** Login as a participant on the web URL with PIN `0442` and verify the shell renders correctly. Open a new tab to Admin Login, use your `YOUR_SUPER_SECRET_CODE`, click Initialize. Ensure both sessions isolate safely.

*Note: Since the codebase was developed in an automated sandboxed CI environment, authorized Google Workbook manual tests against real Apps Script environments must still be completed by the maintainer.*