# Refactor & Feature Plan: MSG File Parser

## Objective

Refactor the codebase for maintainability and extensibility, and add support for exporting MSG files to HTML format with a structured header block.

---

## Step-by-Step Plan
### 7. Modularize Argument Parsing and Validation
- Refactor argument parsing into a dedicated class (`ParserArguments`).
- Refactor input/output validation into a dedicated class (`FileValidator`).
- Remove legacy methods from Program.cs.
- Ensure Program.cs only orchestrates, not implements.

### 1. Design Classes
- **MsgFileParser**: Handles parsing and validation of MSG files.
- **MsgExportService**: Handles exporting parsed data to text or HTML.
- **MsgFileInfo**: Data class to hold parsed MSG file info (subject, sender, recipients, body, etc.).
- **Program**: Entry point, handles CLI, error codes, and calls services.

### 2. Refactor Current Implementation
- Move parsing logic from `ProcessMsgFile` into `MsgFileParser`.
- Move export logic (writing to file) into `MsgExportService`.
- Use `MsgFileInfo` to pass parsed data between classes.
- Keep error handling and CLI logic in `Program`.

### 3. Test Refactored Code
- Ensure all current features and error codes work as before.
- Run the test script to validate.

### 4. Add HTML Export Functionality
- Extend `MsgExportService` to support HTML export.
- Insert message header block at the top of the HTML.
- Ensure output is well-structured and readable.

### 5. Update CLI
- Add a CLI option to choose export format (`--format text|html` or similar).

### 6. Update Tests & Documentation
- Add tests for HTML export.
- Update README.md with new usage and examples.

---

## Tracking Table

| Step | Task | Status |
|------|------|--------|
| 1 | Design classes & structure | ✅ Complete |
| 2 | Refactor current implementation | ✅ Complete |
| 3 | Test refactored code | ✅ Complete |
| 4 | Add HTML export functionality | ✅ Complete |
| 5 | Update CLI for format selection | ✅ Complete |
| 6 | Update tests & documentation | ✅ Complete |
| 7 | Modularize argument parsing & validation | ✅ Complete |

---

## Notes
- The plan will be updated as work progresses.
- Each step should be marked as complete (✅) when finished.
- This file serves as a single source of truth for refactoring and feature tracking.
