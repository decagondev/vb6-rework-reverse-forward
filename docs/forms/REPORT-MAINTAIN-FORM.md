# frmReportMaintain — Report Metadata CRUD Editor

This is the form launched from `frmReport` after the developer-password challenge passes. It's the editor for the `Report` table — the metadata-driven report definitions that the report system reads at runtime. This is where the loop closes: `frmReport` reads rows from the `Report` table and dispatches them; this form is how those rows get edited.

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template as the rest: `imgLogo` left, `lblBusinessName` (24pt, right-aligned this time — back to the convention used by Dashboard/Booking/FindCustomer), `lblCopyright` underneath, `shpCopyright` rectangle.

**Toolbar `tbrMenu`** — four buttons:
- CLOSE (Esc)
- SAVE (Ctrl+S)
- EXPERT (Ctrl+E) — toggles the visibility of the advanced editor pane
- PRINT (Ctrl+P) — *defined but `Enabled = False, Visible = False`*; placeholder for a future print-from-editor feature

`fraLogin` on the right with the live clock and `lblUserID` (idle-watchdog pattern).

**Main body — two-column master/detail layout**

*Left column — `Picture1`:*
- `fraReport` header label "Reports"
- `lvReports` ListView with three columns:
  - `ID` (zero-width, hidden)
  - `Report ID`, `Report Name`
- Color-coded same as `frmReport`: gray for active rows, pink for inactive.

*Right column — `Picture2` containing two sub-panels:*

**Top sub-panel — basic report fields** (always visible):
- `lblReportID` — read-only (looks like a label styled as a textbox), shows the Report ID
- `txtReportName`
- `txtReportTitle`
- `chkDate` — "Show Date" toggle, controls `ShowReportAsOn`
- `txtReportAsOn1` — the "as at" prefix string
- `chkActive` — Yes/No active flag

**Bottom sub-panel `fraAdvanced`** (toggleable, hidden by default behind Expert mode):
- `cboDateField` — dropdown of which date column to filter by (`CreatedDate`, `LastModifiedDate`, `CreatedDate (All 7 days)`, `(None)`)
- `cboDateType` — dropdown of date semantics (Range, Weekly, Monthly, Yearly, Since Start, Single)
- `txtReportQuery` — the actual SQL (multiline, scrollable, MaxLength 1000)
- `txtSubQuery` — appended fragment (single-line, ORDER BY etc.)
- `txtNullQuery` — fallback for empty results (multiline, MaxLength 1000)
- `txtReportFile` — the `.rpt` filename

The two combos' lists come from the form's `frx` resource, presumably hardcoded matching the values that `frmReport.PrintReport` knows how to interpret.

## Module-level state

```vb
intTick                 ' Idle-logout counter
COL_GRAY = #E0E0E0
COL_PINK = #8080FF
```

Same color scheme as `frmReport`.

## Logic

**`Form_Load`**
1. Sets the live clock and user ID label.
2. Permission gate on the EXPERT button: `MOD_REPORT_EDIT_EXPERT`. So there are *two* tiers of edit permission — basic (`MOD_REPORT_EDIT`, gates entry to this form via `frmReport`) and expert (`MOD_REPORT_EDIT_EXPERT`, gates access to the SQL editor pane).
3. Calls `LoadReports`.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog → `frmDialog`).

**`Form_Unload`**
1. Re-shows `frmReport`.
2. **Calls `frmReport.LoadReports`** explicitly — so any changes made here propagate to the parent's report list immediately. This is a clean coordination pattern.

**`Form_KeyDown`** — Esc closes, Ctrl+S saves, Ctrl+E toggles the Expert pane (only if the EXPERT button is enabled).

**`tbrMenu_ButtonClick`** — same three actions wired through the toolbar. Note Ctrl+E in the keyboard handler doesn't gate on `tbrMenu.Buttons("EXPERT").Enabled` for the toolbar branch — so a non-expert user could theoretically bypass the gate by clicking the (still-visible) toolbar button. Looking again, the toolbar button isn't disabled-and-hidden, just disabled — so MSComctlLib should swallow the click. The `Form_KeyDown` shortcut does explicitly check the enabled state. So the gating is consistent in practice.

**`SelectReport(pintReportID)`** — public entry point used by `frmReport.EditReport`:
1. Walks `lvReports.ListItems` looking for a matching ReportID.
2. If found, selects it, scrolls into view via `EnsureVisible`, and calls `lvReports_Click` to populate the editor.
3. If not found, just displays the ID in `lblReportID` without populating other fields. (This handles the "create new" case, though the Save logic doesn't actually insert new rows — see notes below.)

**`lvReports_Click`** / **`lvReports_KeyUp`** — when the user picks a report, calls `PopulateValues`. The keyboard handler is a thin wrapper that delegates to the click handler, so arrow-key navigation works the same as clicking.

**`LoadReports`** — straightforward catalog read:
```sql
SELECT * FROM Report ORDER BY ReportID
```
Adds each row to the ListView, color-codes by `Active`. Calls `lvReports_Click` at the end so the first row's data is auto-loaded into the editor.

**`PopulateValues(plngReportID)`** — re-fetches the selected row and pushes each field into its corresponding control:
- Uses helpers `SetRecord` (text fields) and `SetCheck` (booleans). These presumably handle null-coalescing safely.
- For the two combos, falls back to `ListIndex = 0` if the stored value is empty — so a new or blank report defaults to the first dropdown option.
- The commented-out `txtReportName2`, `txtReportTitle2`, `txtDateType2`, `txtReportAsOn2`, `txtTempTable` lines confirm the schema's bilingual design (the `1`/`2` suffix pattern for English/Malay) and reveal there was once a `TempTable` field — likely related to the `WeeklyBooking` helper table machinery in `frmReport`. The form was simplified to single-language editing and the secondary fields commented out.

**`SaveReport`**
1. Confirmation dialog.
2. Reads `lblReportID.Caption` to determine which row to update.
3. Re-checks the row exists (a `SELECT * WHERE ReportID = ?` followed by `If Not rst.EOF Then`) — but the UPDATE statement is built *inside* the `If Not rst.EOF` block. **This is the issue:** if the row doesn't exist, the SAVE silently does nothing, no INSERT, no error. So the form can only edit existing reports — there's no path here to create a new Report row, despite `SelectReport` being structured as if there were one.
4. Builds an UPDATE via the typed SQL helpers, including `CheckString` (presumably an SQL-escape function that wraps the typed text safely).
5. Boolean conversion for `ShowReportAsOn` and `Active` checkboxes.
6. `WHERE ReportID = mlngReportID`.
7. Closes the recordset, *then* runs the QuerySQL — important ordering, since the read recordset and the write share `gstrSQL`. The recordset has to close before the new SQL is assembled.
8. Confirmation MsgBox, then re-runs `LoadReports` to refresh the ListView (in case the name or active flag changed).

## Notable points and quirks

- **Two-tier edit permission.** `MOD_REPORT_EDIT` gates entry to this form (via `frmReport`); `MOD_REPORT_EDIT_EXPERT` gates the SQL editor pane within. So a "non-expert editor" role would be able to change a report's name, title, "as on" prefix, and active flag without being able to break it by editing the SQL. Sensible separation of concerns. Combined with the developer-password challenge in `frmReport`, that's *three* layers of gating: developer password → MOD_REPORT_EDIT → MOD_REPORT_EDIT_EXPERT for the SQL.

- **The Expert pane is hidden by default.** `fraAdvanced.Visible` is false until Ctrl+E or the EXPERT button toggles it on. So opening this form, a user sees only the basic fields — even an "expert" user has to consciously toggle to the SQL editor. This is a thoughtful UX choice for a tool where the SQL is dangerous to edit.

- **Save doesn't insert new rows.** Only existing reports can be edited. New reports presumably need to be added to the database directly (or via `frmReport`'s metadata management via some path not shown). This is consistent with the metadata-driven architecture being mostly fixed in production with rare additions, but it's worth flagging as a missing CRUD operation.

- **`MaxLength = 1000` on the SQL fields.** Both `txtReportQuery` and `txtNullQuery` cap at 1000 chars. For Access-flavor SQL with explicit column lists and joins, that's tight but workable. Complex reports could hit the limit; the `SubQuery` field (255-char cap) is presumably a safety valve for ORDER BY tails.

- **The `Active` label has a pink ForeColor** (`#8080FF`) — same pink used to color inactive rows. Visually correlates "Active = unchecked" with "row turns pink in the catalog." Subtle but a nice signal.

- **`chkDate.TabStop = False`** and **`chkActive.TabStop = False`** — both checkboxes are removed from the tab order. Tab navigation flows: List → Title → Name → ReportTitle → ReportAsOn → DateField → DateType → Query → SubQuery → NullQuery → ReportFile. The checkboxes can only be toggled by mouse or by mnemonic. Slight usability quirk.

- **`SetRecord` and `SetCheck` are app-wide helpers** that this form leans on heavily. Combined with `CheckString` for SQL-safe writes, the form is mostly a thin shell around standard plumbing.

- **The PRINT button is wired into `ImageList1` (image 4) but disabled and invisible.** The intent was probably to let editors run a quick preview from inside the editor without going back to `frmReport`. Never shipped — same pattern as `frmBooking`'s commented-out VOID button. The infrastructure is there if anyone wants to enable it.

- **Commented-out fields hint at original ambitions.** Bilingual support (`*1`/`*2` field pairs), per-report `TempTable` configuration. Reading the `frmReport.UpdateWeekDayTable` code, that hardcoded "WeeklyBooking" table name was probably *meant* to be configurable per-report via this `TempTable` field — letting different reports use different helper tables — but the feature was dropped. Probably for the better; the simpler "one weekly helper table" design is easier to reason about.

- **`gblnMaintainReport` flag is commented out in `Form_Unload`** — a leftover from when some other form (maybe `frmReport`) needed to know whether the maintenance editor was active. The pattern echoes the commented-out `gblnReportSelect` in `frmReport`. Both are dead.

- **`gblnReportSelect` in `SaveReport` is also commented out** — this was apparently meant to refresh some `frmReportSelect` form (which doesn't appear in any other file I've seen). Possibly an earlier version of the report system had a separate "select to print" form distinct from `frmReport`. Consolidation simplified the architecture.

- **Read-then-write pattern for SAVE.** The form re-reads the row before updating, even though the editor already shows its current state. This is harmless but wasteful — a single UPDATE WHERE ReportID = ? would suffice. The pre-check exists to confirm the row still exists (in case it was deleted by another user), but combined with the silent failure on missing rows it's a cautious design that doesn't quite go far enough. A "row not found, please refresh" message would close the loop properly.

- **Refresh-on-save calls `LoadReports`** which calls `lvReports_Click` at the end — meaning save jumps the selection back to the first row, not the row just saved. Slightly disorienting if you're editing several reports in a row. A small enhancement would be to remember the saved ReportID and re-select it after the refresh.

- **`fraAdvanced` is technically a `PictureBox`, not a `Frame`** — declared as `VB.PictureBox` even though it's used as a container. This is a common VB6 pattern when you want a container that can paint without the visible border of a Frame. Functionally equivalent for the role it plays.

- **Tab order through the Expert pane goes 7, 8, 9, 10, 11, 12** for the six controls (DateField, DateType, ReportQuery, SubQuery, NullQuery, ReportFile). Logical top-to-bottom flow.

- **`SQL_SET_Boolean "Active", False, False`** — the third `False` parameter is the helper's "is-not-last-in-update" flag (the convention is `SET col = val [,]` ending with `False` to not append the comma). So `Active` is always the last field updated, and the comma logic is tracked manually by the helper. Same pattern used elsewhere in the codebase.

- **No timestamp/audit columns are written.** Unlike `frmBooking.SaveBooking` and the status-change handlers in `frmDashboard`, this save doesn't update any `LastModifiedDate` or `LastModifiedBy` fields. Either the `Report` table doesn't have those columns, or the audit trail simply wasn't extended to report metadata changes. Worth flagging if there's ever a "who changed this report's SQL?" question.

- **The form is doing single-record editing in a master/detail layout.** No multi-record edits, no batch updates. Reports are edited one at a time, save-as-you-go. For an admin tool with maybe 10–20 report definitions in production, this is appropriate scale.
