# modGlobalVariable — Global State & Constants Module

The system's central declaration of every global variable and constant the application depends on. Versioned 1.2.22, last touched 19/05/2018 — making this one of the more recently modified modules (alongside `frmUserLogin` and `frmUserChangePassword`, both also dated 19/05/2018, suggesting that May 2018 commit modified the security/login flow and updated the related globals).

This is the module that *names* the system's runtime state. It's small in lines but central in dependencies — every other module and form references something declared here. There are no functions or subroutines, just declarations.

## What this module provides

Three categories of declarations:

1. **The Crystal Reports application reference** — `CrApplication`
2. **Public global variables** — 18 of them, covering company info, user session state, current report, and the SQL builder buffer
3. **Public constants** — one company name + 11 module IDs + 6 report IDs

## The Crystal Reports object

```vb
Public CrApplication As CRAXDRT.Application
```

A reference to the Crystal Reports ActiveX Designer Runtime (`CRAXDRT`) Application object. Set once during `frmSplash`'s startup and used by `frmPrint` to load and render `.rpt` report templates. Living as a global means:

- Crystal Reports is initialized once at app startup (slow operation — milliseconds-to-seconds depending on the install)
- Every print operation reuses the same instance
- The reference must persist for the entire app lifetime, otherwise Crystal will silently disconnect

The "Set CrApplication = New CRAXDRT.Application" call lives in `frmSplash`. When the application terminates, VB6's reference cleanup releases it. This works because the runtime is not a singleton — each fresh instance is a fresh COM activation.

## Company / system globals

```vb
Public gstrCompanyID As String
Public gstrCompanyName As String
Public gstrCompanyName2 As String
Public gstrStreetAddress As String
Public gstrContactNo As String
Public gstrCurrency As String
Public gdtmFiscalYearStart As Date
```

Loaded from the `Company` table at startup, used everywhere for headers, footers, currency formatting, and report titles. A few notes:

- **`gstrCompanyName2`** is a second company name field — same `_2` pattern as `Company.CompanyName2` if it exists in the schema. Looking back at modDatabase's `CreateDB`, only `CompanyName` is created — there's no `CompanyName2` column. So this global is either populated from `F01` (the mystery 50-char column) or never set. Likely planned-but-never-implemented multilingual support, matching the `_1` suffix pattern noted in `Report.ReportName1` and `ModuleAccess.ModuleDesc1`.

- **`gstrCurrency`** holds "MYR" (Malaysian Ringgit) per the seed data. Used by `FormatCurrency`-related display code, though the form analyses showed currency strings are mostly hardcoded as "MYR" or as the symbol embedded in the `.rpt` report templates rather than read dynamically.

- **`gdtmFiscalYearStart`** — a Date. Probably loaded from `Company.SystemStartDate` (1/1/2018 in the seed). Used for "since start" report calculations. The fact that no fiscal-year boundary appears in any of the analyzed forms suggests this global may be set but rarely read.

- **`gstrCompanyID`** is a String even though companies in this single-tenant system don't have meaningful IDs (one row in Company). Probably a legacy from a multi-tenant template.

## User session globals

```vb
Public gstrUserID As String
Public gstrUserName As String
Public gstrUserPassword As String
Public gstrUserSalt As String
Public gintUserIdle As Integer
Public gintUserGroup As Integer
Public gblnUserChangePassword As Boolean
```

These are populated by `frmUserLogin` after successful authentication and consulted everywhere afterwards. The form analyses traced their use in detail:

- **`gstrUserID`** — every form's `lblUserID.Caption = "User ID : " & gstrUserID`, every audit-trail `CreatedBy`/`LastModifiedBy` field, every log entry's `LogUserName`. Most-referenced global in the system.

- **`gstrUserName`** — full name, displayed less prominently. Notably *not* used as the audit-trail field; that's always the UserID.

- **`gstrUserPassword`** and **`gstrUserSalt`** — the *encoded* password hash and salt are stashed at login. Used by `frmUserChangePassword.CheckPassword` to verify the old password without re-querying the database. Storing the password hash in memory throughout the session is a moderate security smell — if the process is dumped or examined, the hash leaks. For a single-machine internal system this is irrelevant; for any networked scenario it's a vulnerability.

- **`gintUserIdle`** — the auto-logout timeout in seconds, loaded from `UserData.Idle`. This is the variable that the `gintUserIdle = 0 ' Test` debug line in `frmUserLogin.cmdOK_Click` permanently overrides to zero, killing the entire idle-watchdog feature across the app. Worth flagging here because the global declaration *expects* it to be a meaningful value; the bug is that it's deliberately blanked at the assignment site.

- **`gintUserGroup`** — the user's group (1=Admin, 4=Clerk per seed data). Available globally so any code can short-circuit "is this user an admin?" without going to the database. As noted in modFunction's analysis, `UserAccessModule` could leverage this to avoid one of its two queries, but doesn't.

- **`gblnUserChangePassword`** — the flag that distinguishes forced from voluntary password changes. Set by `frmDashboard` (F8 = voluntary, sets to True). Cleared by `frmUserLogin` after a forced change-redirect (`gblnUserChangePassword = False` to enforce strict cancel-quits-app behavior). Read by `frmUserChangePassword.cmdCancel_Click` to decide whether Cancel returns to Dashboard or terminates the app. Three-form coordination via one Boolean.

**Note: `gintUserIdle` and `gintUserGroup` are Integer (16-bit signed, max 32,767).** UserGroup values are 1-99; idle is capped at 3600. Both fit. But the `UserData.UserGroup` schema column is `LONG` (32-bit) and `LoginAttempts` is also LONG — so when reading, VB6 implicitly narrows. For the values actually used this is safe, but a database value > 32,767 would overflow. Defensive code in `frmUserLogin` uses `ConvInt` which clamps to 0 on conversion failure.

## Report transfer globals

```vb
Public glngReportNum As Long
Public gstrReportFileName As String
Public gstrReportTitle As String
```

The "what report are we about to print" trio. Set by `frmReport` (or other callers like `frmBooking` for receipt printing) just before showing `frmPrint`. The print form reads these to decide which `.rpt` template to load, what title to display, and which report ID to log. Communication between forms via global state — an alternative to passing arguments through form properties.

This is the pattern that makes `frmPrint` a "universal" print viewer: the caller sets the globals, then `frmPrint.Show` reads them. The downside is that two simultaneous print operations would clash (which can't happen in this single-form-at-a-time UI) and the dependency is implicit rather than visible in any function signature.

## The SQL builder buffer

```vb
Public gstrSQL As String
```

The most consequential single declaration in the whole codebase. Every `SQL_*`, `SQLText`, `SQLData_*`, and `SQL_SET_*` helper in modCommon mutates this string. Every form's database operation reads it. The pattern is invariant: **build into `gstrSQL`, then execute via `OpenRS(gstrSQL)` or `QuerySQL gstrSQL`**.

Implications already noted in the modCommon analysis:
- Cannot build two queries in parallel
- The builder calls have to be made in sequence with no interruption
- Form code that builds a query and then calls another function that *also* uses the builder would corrupt both

The migration code in `frmSplash` works around this by using a separate `OpenData`/`QueryDataSQL` path with a different SQL string buffer (presumably also global, though not declared here — perhaps inlined into the migration code).

This single global is the codebase's defining architectural decision. It's why every form's database code follows the same shape, why parallel queries are impossible, and why a refactor to parameterized queries would be a system-wide undertaking rather than a per-form change.

## Constants

### Company name

```vb
Public Const COMPANY_PRODUCT_NAME = "Hotel Booking System"
```

The display name used in form headers and copyright lines (`lblCompanyProduct.Caption = COMPANY_PRODUCT_NAME` across multiple forms). Note: the value is *"Hotel Booking System"* — but the original branding analyzed across the forms was *"Room Booking System"* and the seed data refers to *"STAR HOTEL"*. Three different names for the same product, used in different places:

- `COMPANY_PRODUCT_NAME` constant: "Hotel Booking System"
- `lblBusinessName.Caption` (form designer default): "Room Booking System"
- `Company.CompanyName` (seed data, end-user customizable): "STAR HOTEL"
- `App.ProductName` (project property): the actual app's product name

The forms display whichever of these is most contextually appropriate. Renaming the product would require updates in at least three places. There's no single source of truth.

### Module access constants

```vb
Public Const MOD_DASHBOARD = 1
Public Const MOD_BOOKING = 2
'Public Const MOD_BOOKING_VOID = 3       ' commented out
Public Const MOD_REPORT_LIST = 3
Public Const MOD_REPORT_PRINT = 4
Public Const MOD_REPORT_EXPORT = 5
Public Const MOD_REPORT_EDIT = 6
Public Const MOD_REPORT_EDIT_EXPERT = 7
Public Const MOD_FIND_CUSTOMER = 8
Public Const MOD_MAINTAIN_ROOM = 9
Public Const MOD_MAINTAIN_USER = 10
Public Const MOD_ACCESS_CONTROL = 11
```

These are the IDs that match the `ModuleAccess.ModuleID` rows seeded by `CreateSampleData`. Every `UserAccessModule(MOD_X)` call resolves through these constants to a specific row in `ModuleAccess`. Cross-referencing with the seed data:

| Constant | ID | Seed name |
|---|---|---|
| MOD_DASHBOARD | 1 | "Dashboard" |
| MOD_BOOKING | 2 | "Booking" |
| MOD_REPORT_LIST | 3 | "List Report" |
| MOD_REPORT_PRINT | 4 | "Print Report" |
| MOD_REPORT_EXPORT | 5 | "Export Report" |
| MOD_REPORT_EDIT | 6 | "Edit Report" |
| MOD_REPORT_EDIT_EXPERT | 7 | "Edit Report (Expert)" |
| MOD_FIND_CUSTOMER | 8 | "Find Customer" |
| MOD_MAINTAIN_ROOM | 9 | "Maintain Room" |
| MOD_MAINTAIN_USER | 10 | "Maintain User" |
| MOD_ACCESS_CONTROL | 11 | "Access Control" |

The constants and the seed data agree perfectly. So adding a new module means: (1) add the constant here, (2) add a row to `ModuleAccess` in the database (either via seed data for fresh installs or via a migration for existing customers), (3) call `UserAccessModule(MOD_NEW)` in the new form's gating code.

### The commented-out `MOD_BOOKING_VOID = 3`

This is the smoking gun for the disabled Void feature noted in the form analyses. The constant was originally `MOD_BOOKING_VOID = 3` — a separate module ID gating the Void operation. When Void was disabled, two things happened:

1. The constant was commented out
2. `MOD_REPORT_LIST`'s value was changed from 4 (?) to 3, taking the freed slot

So the renumbering shifted everything downstream. **This means an old database with the original 4-row Report List → 12 modules numbering is incompatible with the current code.** A customer who upgraded from a pre-Void-disable version would have `ModuleAccess.ModuleID = 3` rows pointing at "Booking Void" while the code now treats ID 3 as "List Report." Any permission check after the renumber would either grant or deny the wrong access.

Worth noting: there's no migration code in `frmSplash`'s `Update_Database` that handles this renumber. So either:
- The renumber happened so early in development that no customer had the old IDs, or
- Customers who upgraded silently got broken permissions and nobody noticed because the system was still permissive enough

Either way, it's a fragile change in retrospect.

### Report ID constants

```vb
Public Const REP_DAILY_BOOKING = 12
Public Const REP_WEEKLY_BOOKING = 13
Public Const REP_MONTHLY_BOOKING = 14
Public Const REP_WEEKLY_DEPOSIT_GRAPH = 15
Public Const REP_SHIFT_FOR_USER = 16
Public Const REP_SHIFT_ALL_USER = 17
```

These are also `ModuleAccess.ModuleID` values (continuing the same numbering), but for *report* permissions rather than form permissions. So the access-control system treats reports as a kind of "module" with its own gate. From the seed data:

| Constant | ID | Seed name |
|---|---|---|
| REP_DAILY_BOOKING | 12 | "Daily Booking Report" |
| REP_WEEKLY_BOOKING | 13 | "Weekly Booking Report" |
| REP_MONTHLY_BOOKING | 14 | "Monthly Booking Report" |
| REP_WEEKLY_DEPOSIT_GRAPH | 15 | "Weekly Booking Graph" |
| REP_SHIFT_FOR_USER | 16 | "Shift Report for User" |
| REP_SHIFT_ALL_USER | 17 | "Shift Report (All Users)" |

But the seed data has **18 module rows** — the last is "Official Receipt (Reprint)" at ID 18. There's no `REP_OFFICIAL_RECEIPT` constant. So official receipt reprinting is gated by a hardcoded `UserAccessModule(18)` somewhere, or it's not gated at all. Looking back at the form analyses, `frmReport`'s permission checks went through the constants — so receipt reprinting either has a missing constant (bug) or is intentionally ungated.

Note also that **REP_WEEKLY_DEPOSIT_GRAPH** is the constant name but the seed data calls the module "Weekly Booking Graph". Mismatch in naming — the original may have been a "deposit graph" that was generalized to "booking graph" without renaming the constant.

## Notable points and quirks

- **All variables are `Public`.** No private state. This is the classic VB6 way — the module-as-namespace pattern. Modern VB.NET would use static class members; here, they're just module-level Publics.

- **Hungarian notation throughout.** `gstrUserID` (g=global, str=string), `gintUserIdle` (g=global, int=integer), `gblnUserChangePassword` (g=global, bln=boolean), `gdtmFiscalYearStart` (g=global, dtm=date/time), `glngReportNum` (g=global, lng=long). Consistent and self-documenting once you know the convention.

- **`gstrCompanyName2` is declared but probably never set.** The Company table doesn't have a `CompanyName2` column per modDatabase's CreateDB. Either it's populated from `F01` or it stays as the default empty string and gets concatenated harmlessly in display strings. Worth grep'ing the form code to confirm — if it's never read, it's dead state.

- **`gstrUserPassword` and `gstrUserSalt` in memory throughout the session** is a security trade-off. Stashing them avoids re-querying the database for password verification (`frmUserChangePassword.CheckPassword` doesn't actually use them — it requeries! So the stashing is unused). Worth verifying: if no code actually reads these globals after they're set, they could be cleared right after assignment to reduce exposure.

- **The MOD_/REP_ constants are flat numbers without enum grouping.** A more modern design would use `Enum ModuleID` to make the type explicit and prevent passing arbitrary integers to `UserAccessModule`. VB6 supports Enum but the codebase doesn't use it.

- **Renumbering hazard.** Because module IDs are baked into both code (constants here) and database (seed rows + customer permission overrides), changing an ID breaks any existing customer. The constants module is effectively a versioning anchor — changing values requires coordinated database migrations. The disabled `MOD_BOOKING_VOID` is the textbook case.

- **No `MOD_PRINT` for the print form.** `frmPrint` doesn't have its own gate — it's reachable only through the forms that themselves gate (Booking, Report, FindCustomer). The print operation inherits its caller's permission. Sensible — printing is an action on already-accessible data.

- **No constant for `frmAdmin`, `frmRoomMaintain`, `frmRoomTypeMaintain`, `frmReportMaintain`, `frmUserChangePassword`.** These forms either have no gate (`frmUserChangePassword` is always accessible to logged-in users), use a different mechanism (`frmRoomTypeMaintain` is gated by being only-reachable-from `frmRoomMaintain` plus Ctrl+T), or are gated by the developer password (`frmReportMaintain` plus the hardcoded "expert" string). The constants list reflects only the *cleanly module-gated* forms.

- **`COMPANY_PRODUCT_NAME` is a constant**, not a variable. So renaming the product requires recompiling the application — customers can't customize it without source code. The customer's *business name* (loaded into `gstrCompanyName` from the database) is freely editable; the *product name* is locked.

- **`Modified On : 19/05/2018`** matches the same date as `frmUserLogin` and `frmUserChangePassword`. So that May 2018 commit touched all three together — probably restructuring how `gblnUserChangePassword` flows between login and password-change forms, given that's the cross-form coordination point. The previous version likely had different (or no) coordination, and the 2018 update introduced the forced-vs-voluntary distinction.

- **`Option Explicit` is set.** Standard hygiene.

- **The module has no functions.** Pure declarations. This makes it a "header file" in the C sense — the rest of the system depends on it but never executes anything from it.

- **No initialization code.** Globals start at their default values (empty strings, 0 integers, False booleans, Nothing for object refs). They're populated by the first form that needs them — `frmSplash` for `CrApplication`, `frmSplash` again for `gstrPassword` (declared in modDatabase but set here in spirit), `frmUserLogin` for the user globals, `frmReport` or `frmBooking` for the report globals. Ad-hoc initialization throughout the app means there's no single "warm up the globals" function.

- **The split between modGlobalVariable and modDatabase is incomplete.** `gstrPassword`, `gstrDatabasePath`, `gstrDatabaseExt`, and the two ADO connections (`ACN`, `DAT`) are declared in modDatabase, not here. So globals are scattered across two modules based on a "data access globals stay with data access code" principle. Reasonable, but means a developer hunting for "where is this declared" has to check both modules.

- **`gstrSQL` is the only mutable infrastructure global.** Everything else is either a connection (set once at startup) or a session value (set at login). `gstrSQL` is the one thing that changes constantly during normal operation — every database call rebuilds it. This is the single biggest clue that the SQL builder pattern is the load-bearing architectural choice; it has its own dedicated global slot that the entire codebase reads and writes.

- **18 module-access rows in the database, but only 11 MOD_ constants and 6 REP_ constants** = 17 named constants for 18 rows. The orphan is `Official Receipt (Reprint)` (ID 18). Either an oversight (no constant for a real module) or evidence that the constants module hasn't been kept in sync with the seed data. A future developer adding ID 19 (Statement Generator, Inventory Module, etc.) would have to remember to add the constant here AND seed the row in `CreateSampleData` AND write a migration in `frmSplash.Update_Database` for existing customers. Three-place coordination, no compile-time enforcement.
