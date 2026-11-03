# Repository Cleanup, Consolidation, and Production-Readiness

Act as a senior software engineer responsible for reviewing, improving, testing, and completing every project in this repository.

The goal is to make these applications production-ready, accurate, fast, reliable, and easier to use without disrupting the workflows and data I already depend on.

Do not stop at analysis, recommendations, plans, scaffolding, or partial implementations. Complete all safe and actionable work, test it, and commit it in logical stages.

---

## Repository Context

This repository contains several separate projects with different purposes.

Keep the projects separate, except:

- Merge `HR Data Toolbox Web v1.3` into `Audit-Processor`.
- The result should be one modular web application containing all Excel audit and HR data-processing functions from both projects.
- Preserve all useful functionality from both applications.
- Remove the fact that the audit processor app, when it does an audit it leaves an audit summary page, dont include that in the audit excel file.

Use extra caution with:

### `RFP-Utility-Application`

- This application is used daily.
- An established workflow depends on its current behaviour.
- Preserve existing inputs, outputs, results, file formats, column structures, naming conventions, and workflow.
- Add regression tests before making significant changes.
- Avoid breaking changes.

### `Outlook-export`

- This application has an existing database and existing exported files.
- Preserve compatibility with the current database, records, export folders, file names, and workflow.
- Do not delete, reset, overwrite, or destructively migrate existing data.
- Back up the database before testing migrations.
- Any schema changes must be safe, tested, and reversible.

The other projects may be restructured and improved more extensively, provided their useful functionality remains available.

---

# Working Rules

1. Inspect and understand each project before changing it.
2. Check the Git status and preserve all existing uncommitted work.
3. Do not commit databases, exports, logs, credentials, HR data, personal information, temporary files, or secrets.
4. Use temporary directories, copied databases, anonymized fixtures, and test data for validation.
5. Keep changes focused and reviewable.
6. Add tests before or alongside significant refactoring.
7. Validate representative workflows, not only individual functions.
8. Do not claim something is fixed or production-ready unless it was tested.
9. Commit each completed and validated logical stage.
10. Do not rewrite shared Git history or force-push.

Continue autonomously when a change is safe, backward-compatible, testable, and reversible.

Pause and consult me only if a necessary change would:

- Break or materially alter an existing workflow.
- Change the meaning or structure of an established output.
- Remove functionality that may still be used.
- Require a destructive or difficult-to-reverse database migration.
- Require a business rule that cannot be determined from the repository.

When consultation is required, briefly explain the issue, impact, recommended change, alternatives, compatibility risks, and rollback approach.

---

# Step 1: Inspect and Organize the Repository

Review every project and identify:

- Its purpose and current workflow
- Entry points and run commands
- Inputs and outputs
- Dependencies
- Configuration
- Databases and stored data
- External integrations
- Tests
- Known risks
- Workflow-sensitive behaviour
- Dead, duplicated, or obsolete code
- Potentially destructive operations

Clean and standardize each project structure where appropriate. Separate source code, business logic, user interface code, configuration, tests, database logic, scripts, temporary files, and generated outputs.

Do not impose unnecessary complexity on small projects.

Remove only files confirmed to be generated, obsolete, duplicated, or unused. Do not remove uncertain files without investigating them.

Update `.gitignore` files to exclude:

- Secrets and environment files
- Databases
- User exports
- Logs
- Temporary files
- Cache files
- Build artifacts
- Uploaded HR data
- Generated results

Create or update:

- A concise root `CLAUDE.md` containing repository context, key commands, project cautions, testing instructions, and workflow-protection rules.
- A concise root `PLAN.md` containing the active task list, completed work, known issues, and decisions requiring my approval.
- A practical `README.md` for each project explaining its purpose, setup, configuration, run commands, inputs, outputs, tests, and important limitations.

Keep documentation brief and useful. Do not produce large reports. For each completed step, add only a short paragraph summarizing what changed, why it changed, and how it was validated.

---

# Step 2: Refactor and Improve the Applications

Review each project and implement the most valuable safe improvements.

Where appropriate:

- Simplify complex code.
- Break large files and functions into maintainable modules.
- Separate UI, business logic, file processing, and database access.
- Remove confirmed dead code and unnecessary duplication.
- Improve names, typing, and configuration.
- Remove hard-coded machine-specific paths.
- Add input and output validation.
- Improve exception handling and error messages.
- Replace silent failures and overly broad exception handling.
- Add structured logging without exposing unnecessary HR or email content.
- Use safe database transactions.
- Prevent partial or corrupted output files.
- Use temporary files and atomic writes where practical.
- Improve large-file performance.
- Avoid unnecessary repeated reads and row-by-row operations.
- Add progress indicators for long-running tasks.
- Prevent users from accidentally starting duplicate jobs.
- Improve the interface and normal user workflow.
- Add safe recovery or resume behaviour where practical.

Do not rewrite stable applications solely to use newer technologies. Make improvements based on measurable value, maintainability, reliability, or user experience.

## `RFP-Utility-Application`

Before substantial refactoring:

- Document the current workflow briefly.
- Capture representative inputs and expected outputs.
- Add regression tests for important transformations.
- Identify file names, columns, formats, paths, and behaviours used externally.

Preserve current functionality and output compatibility.

Handle malformed, incomplete, locked, missing, or invalid files gracefully without changing the normal workflow.

## Merge `HR Data Toolbox Web v1.3` into `Audit-Processor`

Before merging:

- Inventory every function in both applications.
- Identify overlapping and unique functionality.
- Create a concise feature-parity checklist in `PLAN.md`.
- Identify differences in file handling, outputs, configuration, and dependencies.

Build one modular web application with:

- A clear home page or tool-selection screen
- A separate maintainable module for each audit or processing function
- Shared Excel-reading and writing utilities
- Shared file validation
- Shared reporting and error handling
- Progress feedback for long operations
- Downloadable results and error reports
- Clear processed, matched, unmatched, skipped, warning, and failed counts where relevant
- Safe temporary-file handling
- A documented temporary-file cleanup policy
- No unnecessary retention of uploaded HR data

Preserve formulas, headers, worksheets, data types, row structures, and formatting wherever the existing workflow requires them.

One failed audit must not crash unrelated tools. The system must not silently lose rows or report success when processing is incomplete.

Do not remove the original applications until:

- Every useful feature is mapped.
- Feature parity is demonstrated.
- Representative outputs are compared.
- Tests pass.
- A rollback path exists.

---

# Step 3: Bug Audit and Remediation

Scan all projects for bugs and operational risks, including:

- Incorrect business logic
- Data loss or silent row loss
- Duplicate or missed processing
- Partial writes
- Hard-coded paths
- Incorrect Excel header assumptions
- Missing or unexpected columns
- Empty and null data
- Date, time-zone, and locale issues
- Formula and data-type corruption
- Locked, corrupt, or unreadable files
- Encoding and Unicode issues
- Database integrity and migration problems
- Duplicate records
- Concurrency problems
- Resource leaks
- Excessive memory usage
- Long-running tasks without progress
- Unsafe filenames and paths
- Sensitive data in logs
- Unhandled exceptions
- Misleading success messages
- Missing dependency controls
- Temporary-file accumulation
- Platform-specific assumptions

Record issues concisely in `PLAN.md` with:

- Project and location
- Problem
- Severity
- User or data impact
- Proposed fix
- Compatibility risk
- Resolution status

Fix issues immediately when the fix is safe, backward-compatible, testable, and reversible.

If a necessary fix could alter an established workflow, stop and consult me before implementing that specific change. Continue with unrelated safe work.

---

# Step 4: Fix and Strengthen `Outlook-export`

The current export process can fail on an unreadable email without clearly identifying which email failed. This makes it impossible to verify a large export containing thousands of emails.

Implement a complete export audit and reconciliation system.

## Required Results

After every export or validation run, the application must identify:

- Total emails discovered
- Total eligible for export
- Total successfully exported
- Total already exported and verified
- Total intentionally skipped
- Total failed
- Exact emails that failed
- Reason and processing stage for each failure
- Emails missing from the database
- Database records missing export files
- Export files missing database records
- Duplicate emails, records, or output files
- Items safe to retry
- Overall status: complete, complete with warnings, or incomplete

## Email Identity

Use the most stable available identifier after inspecting the existing implementation, such as:

- Outlook Entry ID
- Store ID
- Internet Message ID
- Existing database key
- A documented composite fallback if no stable identifier is available

Do not use subject lines as the unique identifier because subjects can be duplicated.

## Export Manifest

Create or extend a manifest or database reconciliation record containing the fields available and necessary for tracing, such as:

- Stable email identifier
- Internet Message ID
- Subject
- Sender
- Received or sent timestamp
- Mailbox and folder
- Expected and actual output path
- File name and format
- Export status
- Database status
- File-system status
- File size
- Attempt count
- Error category
- Error message
- Processing stage
- First and latest attempt timestamps
- Final status

Store only the metadata needed for reconciliation. Do not unnecessarily duplicate email content.

## Reconciliation Checks

Compare the Outlook source, database, and export folders to detect:

- Source emails missing database records
- Source emails missing files
- Database records missing files
- Files missing database records
- Empty or zero-byte files
- Corrupt or unreadable files
- Incorrect file paths
- Conflicting statuses
- Duplicate identifiers
- Duplicate output paths
- Multiple files for one email
- Multiple emails mapped to one file
- Interrupted or incomplete exports

Where practical, verify that each exported file exists, is non-empty, can be opened, and contains the required metadata. Optionally use checksums when useful for future integrity checks.

## Error Reporting

Every failure must identify the affected email using the safest available combination of:

- Stable identifier
- Subject
- Sender
- Date and time
- Source folder
- Intended output path
- Processing stage
- Exception type
- Clear error explanation
- Whether it can be retried

Produce:

- A short human-readable summary
- A machine-readable report such as CSV, JSON, or a database table

Avoid saving unnecessary message content or sensitive information in logs.

## Retry and Recovery

Add a safe workflow to:

- Retry only missing items.
- Retry only failed items.
- Resume interrupted exports.
- Avoid re-exporting or duplicating verified items.
- Preserve previous errors and attempt history.
- Distinguish retryable from non-retryable failures.
- Produce a new reconciliation summary after retrying.

Make the export process idempotent wherever practical. Re-running the same export should not create uncontrolled duplicates or corrupt existing records.

## Database Safety

Before modifying the existing database:

- Inspect the current schema.
- Create a backup.
- Test against a copy.
- Preserve every existing record.
- Use a safe migration if new fields or tables are required.
- Make the migration transactional where supported.
- Provide a rollback or recovery path.
- Do not require historical exports to be rebuilt.

## Outlook Export Tests

Add tests using mocks, test doubles, or sanitized fixtures for:

- Successful export
- Unreadable email
- Missing database record
- Missing output file
- Empty or corrupt file
- Duplicate identifier
- Duplicate subject
- Duplicate output path
- Interrupted export
- Safe resume
- Retry of failed items
- Existing file with incomplete database status
- Database record with missing file
- Invalid filename characters
- Unicode subjects
- Long paths
- Large batches
- Database migration and recovery

If Outlook integration cannot be tested in the current environment, test all independent logic using mocks and clearly list the exact validation that must be completed on the target Windows machine.

---

# Step 5: Testing and Production Hardening

Add or improve:

- Unit tests
- Regression tests
- Integration tests
- File-processing tests
- Database migration tests
- Representative workflow tests
- End-to-end smoke tests where practical

Test normal and failure scenarios, including:

- Empty inputs
- Invalid inputs
- Missing or unexpected columns
- Duplicate rows
- Null values
- Large files
- Locked or corrupt files
- Unicode content
- Long paths
- Permission failures
- Interrupted operations
- Partial failures
- Repeated execution
- Existing output files

Where appropriate, add:

- Reproducible dependency installation
- Pinned or constrained dependencies
- Environment templates without secrets
- Startup configuration validation
- Structured logging
- Graceful shutdown
- Backup and recovery instructions
- Safe launch or build scripts
- User-friendly error messages
- Diagnostic reports

For performance work:

- Identify the actual bottleneck.
- Establish a representative baseline where possible.
- Measure significant changes.
- Avoid unsupported claims about speed improvements.
- Record only a brief before-and-after summary.

---

# Validation Before Completion

Before declaring the repository complete:

- Run all available tests.
- Run linting, formatting, type checking, and builds where supported.
- Run dependency and security scans where tools are available.
- Search for secrets and hard-coded credentials.
- Search for machine-specific absolute paths.
- Search for unresolved `TODO`, `FIXME`, debugging code, and placeholders.
- Verify setup and run commands.
- Verify representative user workflows.
- Verify no databases, exports, personal information, or generated HR files are committed.
- Confirm feature parity for the merged audit application.
- Confirm Outlook reconciliation identifies exact missing and failed emails.
- Confirm failed Outlook items can be retried without duplicating successful exports.

If something cannot be tested because of missing credentials, Outlook access, unavailable dependencies, or environment restrictions:

- Complete everything that can be tested.
- Use mocks or sanitized fixtures.
- State briefly what could not be tested.
- Provide the exact command or steps required to validate it in the target environment.
- Do not claim that unexecuted validation passed.

---

# Commit Requirements

Commit each completed and validated logical stage.

Use focused commit messages such as:

- `docs: add concise repository guidance`
- `chore: standardize project structure`
- `test: add RFP workflow regression coverage`
- `refactor: improve RFP processing reliability`
- `feat: merge HR data tools into Audit Processor`
- `feat: add Outlook export reconciliation`
- `fix: identify and retry unreadable Outlook items`
- `test: cover Outlook export recovery scenarios`

Before each commit:

1. Review the diff.
2. Run the relevant tests.
3. Confirm no secrets, databases, exports, personal information, or generated results are staged.
4. Update the concise task status in `PLAN.md`.
5. Do not commit known broken work.

---

# Definition of Done

The work is complete only when:

- Each application installs and runs using documented commands.
- Existing useful functionality is preserved.
- Workflow-sensitive behaviour is regression-tested.
- Inputs and outputs are validated.
- Errors are specific and actionable.
- Important failures are logged safely.
- Representative workflows pass.
- No unresolved critical or high-severity bugs remain.
- Remaining limitations are briefly documented.
- Changes are committed in focused, working stages.

The merged audit application is complete only when all useful functionality from both original applications is available and tested in the consolidated web application.

`Outlook-export` is complete only when every discovered email receives a traceable status, exact failed items can be identified, the database and file system can be reconciled, and missing or failed items can be retried safely.

---

# Final Response

At the end, provide a concise completion summary containing:

- A brief paragraph for each project describing the changes made
- Important bugs fixed
- Compatibility status
- Tests and validation performed
- Any remaining limitations or items requiring my decision
- Commit hashes and messages

Do not create a large final report unless I specifically request one.

Begin by inspecting the repository, checking Git status, identifying project boundaries, and creating the initial concise `PLAN.md`. Then proceed through the steps autonomously, testing and committing each completed stage.