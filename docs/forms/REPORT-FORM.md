# frmReport — Report Catalog & Launcher

This is the F2-from-Dashboard form. It shows a list of all available reports stored in a `Report` table, lets the user pick one, applies date-range parameters chosen from the date picker, and dispatches to `frmPrint` for the actual rendering. It's the centerpiece of a metadata-driven reporting system — reports are defined as database rows, not as code.

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template as the rest: `imgLogo` left, `lblBusinessName` (24pt, **center-aligned** like `frmPrint`), `lblCopyright` underneath, `shpCopyright` rectangle.

**Toolbar `tbrButton`** — three buttons:
- CLOSE (Esc)
- EDIT (Ctrl+E) — opens `frmReportMaintain` after a developer-password challenge
- PREVIEW (Ctrl+V) — runs the selected report

`fraLogin` on the right with the live clock and `lblUserID` (idle-watchdog pattern).

**Main body — `Picture1`**
- `Label5` "Report as at" + `dtpTransDate` (DTPicker, default = today). This single date drives whatever date-range logic the selected report calls for.
- `lvReports` ListView (Report mode = 3, i.e., Details) with 7 columns:
  - `ID` (zero-width, hidden)
  - `Report ID`, `Report Name`, `Report Title`, `Report As On`, `Date Type`, `Date Field`

The ListView is the catalog — every active report shows up as one row, color-coded by its enabled state.

## Module-level state

```vb
strDeveloperPassword  ' Hardcoded "expert" — gate for entering the report editor
intTick               ' Idle-logout counter
COL_GRAY = #E0E0E0    ' Active report rows
COL_PINK = #8080FF    ' Inactive report rows
```

## Logic

**`Form_Load`**
1. Sets `strDeveloperPassword = "expert"` (hardcoded — see notes below).
2. Sets the live clock, user ID label.
3. Defaults `dtpTransDate` to today.
4. Calls `LoadReports`.
5. Gates the EDIT button on `MOD_REPORT_EDIT` permission.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog).

**`Form_Unload`** — re-shows `frmDashboard`.

**`Form_KeyDown`** — handles Esc (close), Ctrl+E (edit, with developer-password prompt), Ctrl+V (preview the selected report). Empty-selection guard on Ctrl+V.

**`tbrButton_ButtonClick`** — same three actions wired through the toolbar.

**`lvReports_DblClick`** — double-click on a report shortcuts straight to Preview.

**`LoadReports`** — reads `Report` table:

```sql
SELECT * FROM Report ORDER BY ReportID
```

For each row, it does a permission check: `UserAccessModule(rst!ReportID + 11)`. The `+ 11` offset means report permissions live in the `ModuleAccess` table at IDs 12–17 (which matches the "Report" rows in `frmModuleAccess` — rows 12, 13, 14, 15, 16, 17 are all labeled "Report"). So Report ID 1 maps to ModuleAccess ID 12, Report ID 2 to ID 13, and so on. This is the binding that makes the access control matrix actually do something for individual reports.

If the user has access, the row is added to the ListView with all six fields. Active reports are colored gray (the standard light-on-dark text), inactive ones pink. Inactive reports are still shown — `PrintReport` will refuse to run them but the user can see they exist.

**`EditReport`** — opens `frmReportMaintain` (a separate form not in this analysis batch, but presumably a CRUD editor for the `Report` table), passing the selected `ReportID` via `SelectReport`. Calls `.ZOrder 0` to make sure it sits on top, then hides `frmReport`.

**`PrintReport(lngReportID)`** — the heart of the form. This is where the metadata-driven query assembly happens.

Step-by-step:

1. **Re-fetch the selected report row** by ID.

2. **Active check.** If `Active = False`, abort with "This report has been disabled."

3. **File existence check.** Verifies the `.rpt` file named in `rst!ReportFile` actually exists in `App.Path\Report\`. Aborts cleanly if missing.

4. **Pull the base query.** `gstrSQL = rst!ReportQuery` — every Report row carries its own SQL string.

5. **Title substitution.** `gstrReportTitle = rst!ReportTitle1`, then `Replace($UserID$, gstrUserID)` so titles can include the operator's user ID dynamically.

6. **Date filter assembly** — only if `DateField1 <> "(None)"`:
   - Decides whether to append `AND` (if the query already has a WHERE) or `WHERE`.
   - Appends the column name from `DateField1`. Special case: `"CreatedDate (All 7 days)"` triggers `UpdateWeekDayTable` — see below.
   - Then appends a date-range clause based on `DateType1`:
     - **Range / Single**: same day, midnight to 23:59:59.
     - **Weekly**: `WeekDay1(date)` to `WeekDay7(date)` — Monday through Sunday of the report date's week.
     - **Monthly**: `MonthDay1` to `MonthDay30`. Note the bug: `MonthDay30` always returns the 30th of the month — so February reports miss the last day or two of any 31-day month. Worth flagging.
     - **Yearly**: Jan 1 to Dec 31 (`YearDay365`).
     - **Since Start**: from `gdtmFiscalYearStart` (a global) up to the report date — useful for year-to-date reports.
   - Builds a human-readable `strDate` for the report header alongside.

7. **SubQuery extension.** If the row has a `SubQuery` field, that gets appended verbatim — lets a report row carry an `ORDER BY` or `GROUP BY` tail.

8. **`$UserID$` substitution** in the SQL itself (so reports can be self-filtered to the current user).

9. **`$BookingID$` parameter.** If the query contains this placeholder, an `InputBox` prompts the user for a booking number (default 100001), and the placeholder gets replaced. This is how single-booking reports prompt for which booking to print.

10. **Final report-title decoration.** If `ShowReportAsOn = True`, appends the `ReportAsOn1` prefix and the human-readable date string to the title.

11. **Empty-result fallback.** `If Not QueryHasData(gstrSQL) Then gstrSQL = rst!NullQuery`. Every report row has a fallback "null query" that returns the Company header with placeholder zeros — so even a no-data report still renders headers and footer rather than crashing or showing nothing. Same defensive pattern as `frmBooking.PrintReceipt`.

12. **Hand off to `frmPrint`.** Sets `gstrSQL`, `gstrReportFileName`, `gstrReportTitle`, instantiates `New frmPrint`, calls `.Show`. The print form then opens its own connection, runs the SQL into a recordset, pushes it into the Crystal report, and renders.

There's a redundant `OpenRS(gstrSQL)` call right before showing `rep` — the assigned `rst` variable isn't used by anything (the print form opens its own recordset). Looks like leftover scaffolding.

**`UpdateWeekDayTable(datReportDate)`** — a clever hack for weekly graph reports.

There's a `WeeklyBooking` table with 7 rows (one per weekday). Before running a weekly report, this sub UPDATEs each row's `CreatedDate` to the corresponding weekday's date relative to the report date. So if you run a "weekly bookings graph" for April 15 (a Tuesday say), the table gets stamped with Monday April 14, Tuesday April 15, … Sunday April 20.

Why? Crystal Reports cross-tabs and graphs need a date column to bucket by. Rather than building cross-tab queries on the fly, the developer pre-stages a 7-row helper table with the right week's dates, then joins against it in the report query. The graph then reliably has 7 buckets even if no bookings occurred on some days.

It's a workaround for either (a) Access SQL not having a clean way to generate a date series, or (b) Crystal Reports having limited tolerance for sparse date axes. Functional but means the table has to be writable and non-concurrent — two users running weekly reports at the same time could clash.

## Notable points and quirks

- **Reports are data, not code.** The entire reporting system is metadata-driven: queries, file names, titles, date logic, and fallbacks all live in the `Report` table. This is genuinely impressive for a 2014 VB6 app — adding a new report means inserting a database row and dropping an `.rpt` file in the Report folder, no recompile. `frmReportMaintain` (referenced but not in this batch) is presumably the CRUD editor for these rows.

- **The `ReportID + 11` offset is a magic number.** Reports are numbered 1, 2, 3, …, but their permissions live at ModuleAccess IDs 12, 13, 14, … because the first 11 permission slots are used by other modules. If anyone ever inserts a new module-level permission and shifts the numbering, the reports' permission gates silently break. A cleaner design would have `Report.PermissionModuleID` as a foreign key column rather than relying on an arithmetic offset.

- **Hardcoded developer password "expert".** This is a textbook example of a backdoor that should not exist in production code. It's literally typed in plaintext in `Form_Load`, so anyone who reads the source (or decompiles the executable) gets full edit access to all reports regardless of their database permissions. Combined with `MOD_REPORT_EDIT` it's a double gate — but the password is so trivially discoverable that it doesn't add real security. The intent was probably "make this hard for end-users to stumble into" rather than "stop a determined attacker," and that's honest about what it achieves.

- **`InputBox` for the developer password is plain text.** It doesn't even mask the password as it's typed. Reinforces that the protection is cosmetic.

- **Single date picker, multiple semantics.** The same `dtpTransDate` value means "this exact day" for Range/Single, "the week containing this day" for Weekly, "the month of this day" for Monthly, etc. Compact UI, but the user has to know what the selected report does with the date. The `Date Type` and `Report As On` columns in the ListView help by spelling it out.

- **Report-as-at convention.** `ReportAsOn1` is a prefix like "as at " or "for " that gets prepended to the date. Suffix `1` suggests there were originally `ReportAsOn2`, `ReportAsOn3` — likely for multi-language support (English/Malay), since the database uses MYR currency. Same suffix appears on `ReportName1`, `ReportTitle1`, `DateType1`, `DateField1`. So the schema supports localized report metadata, but the form only ever reads the `1` variant.

- **`MonthDay30` is wrong for ~7 months a year.** As noted — a Monthly report run for, say, March 15 will only cover March 1–30, missing March 31. Same for May, July, August, October, December. February gets either the right amount of days or 1–2 too many. This is a real data-integrity bug, though probably mitigated in practice because hotel bookings on a single day are unlikely to be the difference between a meaningful report and a misleading one.

- **No date validation.** If the user picks a future date for a "Yearly" report, you get bookings up through Dec 31 of a future year — which will return real data if any bookings span into that range. Probably benign but no guard.

- **`UpdateWeekDayTable` is a side effect of running a report.** Pulling up a weekly booking graph rewrites the contents of `WeeklyBooking` — meaning any concurrent report viewer would see inconsistent data. For a single-user hotel system this is fine. For a multi-user one it'd be a race condition.

- **`$BookingID$` and `$UserID$` are simple string substitution.** No escaping, so a UserID containing an apostrophe would break the SQL. In practice User IDs are admin-controlled and stable, so this is low-risk, but it's another spot where the dynamic SQL pattern leaks.

- **`Replace(gstrSQL, "$UserID$", gstrUserID)` runs after the date clause is appended** — so the ReportQuery can contain `$UserID$` anywhere (in the SELECT, WHERE, ORDER BY, etc.), not just at fixed positions. Same for `$BookingID$`.

- **Inactive reports are visible but not runnable.** Showing them in pink lets staff see what reports exist that they could potentially re-enable. The active gate in `PrintReport` is what enforces it.

- **Two `OpenDB`/`CloseDB` cycles in `PrintReport`** — one to read the report metadata, then a second one before showing `frmPrint`. The second one is dead (the recordset isn't used). `frmPrint` opens its own connection anyway. Cleanup opportunity.

- **`gblnReportSelect`** appears commented out in both `Form_Load` and `Form_Unload`. Probably a leftover global flag from a previous design where some other form needed to know whether the report selector was open. Cleanly disabled.

- **The whole report-metadata pattern is repeated in `frmPrint`.** The two forms together implement: "report rows in DB define what gets printed → frmReport assembles the query and parameters → frmPrint loads the .rpt and pushes the recordset." It's a clean separation: this form does parameter binding, the print form does rendering.

- **Default `dtpTransDate.Value = Date`** uses Visual Basic's `Date` keyword (today's date). Simple and right.

- **The `+ 680` taskbar offset trick from `frmPrint` is absent here.** This form just centers on owner via `StartUpPosition = 1`, which is the standard VB6 way and doesn't need the manual fix. So the `frmPrint` offset really was a workaround for a specific issue in that form's geometry.
