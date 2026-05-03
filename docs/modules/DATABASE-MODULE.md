# modDatabase — Database Connection & Schema Module

The system's database layer. Versioned 1.2.22, last touched 17/12/2014, attributed to Aeric Poon. This is where the application talks to Microsoft Access via ADO — connection management, recordset opening/closing, query execution, and (critically) the entire database schema definition for fresh installs. Two thirds of the file is the schema-creation routine `CreateDB` and the seed-data routine `CreateSampleData`. The rest is the connection/query plumbing every form uses.

## What this module provides

The module's responsibilities split into five distinct categories:

1. **Public state** — two ADO connections (`ACN`, `DAT`) and three connection-string globals (`gstrDatabasePath`, `gstrDatabaseExt`, `gstrPassword`)
2. **Primary connection lifecycle** — `OpenDB`, `CloseDB`, `ConnectDB`
3. **Database creation** — `CreateData` (the empty .mdb file), `CreateDB` (the schema), `CreateSampleData` (seed rows)
4. **Recordset/query helpers for the primary connection** — `OpenSQL`, `OpenQuery`, `OpenRS`, `OpenTable`, `ExecuteSelectSQL`, `CloseRS`, `QuerySQL`, `UnlockDB`
5. **Parallel "data" connection helpers** — `OpenData`, `CloseData`, `OpenDataRS`, `CloseDataRS`, `OpenDataTable`, `QueryDataSQL` — a complete second set wrapping `DAT` instead of `ACN`

The "Data" suffix family was added in this version's update — the module's header comment lists them explicitly. They exist so the migration code in `frmSplash` can read from the *old* database (via `DAT`) while writing to the *new* database (via `ACN`) simultaneously.

## Public state

```vb
Public ACN As ADODB.Connection      ' Primary connection (current DB)
Public DAT As ADODB.Connection      ' Secondary connection (migration source)
Public gstrDatabaseExt As String    ' .mdb (declared but unused in this module)
Public gstrDatabasePath As String   ' Full path to current .mdb
Public gstrPassword As String       ' Jet OLE DB password (set from GenWord)
```

The two-connection design is the architectural quirk that drives the whole "Data" parallel API. Most VB6 database apps have a single global ADO connection; this one has two, so it can do source→destination data migration in-process without temp files or external tools.

`gstrDatabaseExt` is declared but never read in this module. It's set by `frmSplash`/`frmDatabase` to `.mdb` and used by other modules for path construction.

## Connection lifecycle

### `OpenDB()`
The bread-and-butter open. Creates a new `ADODB.Connection`, sets the Jet 4.0 provider, applies the global path and password, opens. Used at the start of every database operation across every form. The pattern across the codebase is: `OpenDB` → build SQL → `OpenRS`/`OpenSQL`/`QuerySQL` → `CloseRS` → `CloseDB`.

### `ConnectDB() As Boolean`
Identical to `OpenDB` but returns Boolean (success/failure). Used by `frmDatabase` to test whether a chosen file is a valid, password-readable database before committing to it. The duplication is wasteful — `OpenDB` could just have been refactored to a Boolean-returning version and called from both places, but instead the entire body is copy-pasted.

### `CloseDB()`
Defensive close. Checks `ACN Is Nothing` first, then `ACN.State = adStateOpen`, then closes and nils. The empty `If ... Then ... Else` pattern (`If ACN Is Nothing Then ... Else ...`) is unusual VB6 idiom — most developers would write `If Not ACN Is Nothing Then`. The current style works but reads backwards.

### `OpenData(path, password)` / `CloseData()`
Same as `OpenDB`/`CloseDB` but operates on `DAT`. Crucially, `OpenData` takes the path and password as parameters rather than reading from globals. This is what enables the migration workflow: `OpenData` opens the *old* database (whose path differs from `gstrDatabasePath`), while `OpenDB` continues to point at the new one.

## Database creation

### `CreateData() As Boolean`
Creates an empty .mdb file using ADOX (`ADOX.Catalog.Create`). This is the lowest level — just the file shell with the password set. Called by `frmDatabase` when the user opts to create a new database file at startup. After this returns True, `CreateDB` is called to add tables.

### `CreateDB() As Boolean`
The schema definition. **This is where the entire database schema lives.** Twelve tables created in sequence using the `SQL_CREATE`/`SQL_COLUMN_*` builder family from modCommon:

| Table | Purpose | Notable columns |
|---|---|---|
| `Booking` | Active bookings | `Temp` (draft flag), `Active` (soft delete), full denormalized room snapshot |
| `Company` | Single-row company info | `DatabaseVersion` (1.3 baseline), `CurrencySymbol`, `F01` (mystery) |
| `LogBooking` | Booking history snapshots | Full Booking columns + `BookingID` FK |
| `LogError` | Error log | DateTime, error num, description, user, module, method, type |
| `LogRoom` | Room edit history | `BookingID` FK + full Room snapshot at time of edit |
| `ModuleAccess` | Permission matrix | `Group1`–`Group4` Yes/No columns |
| `Report` | Report metadata catalog | `ReportQuery`, `SubQuery`, `NullQuery` (memo fields), date config |
| `Room` | Master room list | ID is `LONG NOT AUTOINCREMENT` (caller-supplied) |
| `RoomType` | Room type lookup | TypeShortName/LongName |
| `UserData` | User accounts | UserGroup, encrypted password, salt, idle, attempts |
| `UserGroup` | User group definitions | GroupID PK, SecurityLevel for ordering |
| `WeeklyBooking` | Pre-aggregated weekly stats | Used by the "Weekly Booking (Chart)" report |

**Several schema observations worth flagging:**

- **`Room.ID` is `LONG PRIMARY KEY` but NOT autoincrement.** Look at `SQL_COLUMN_ID "ID", True, False, True` — third arg `False` = not autoincrement. So Room IDs are caller-supplied (the form passes the room's display number as its ID). Every other table uses autoincrement IDs. This is why `frmRoomMaintain` has logic to compute the next available ID manually. It also means deleting a room and re-adding it would conflict — but since deletion is soft, this rarely matters.

- **`Booking.Temp` defaults to `"Yes"`.** New booking rows are draft until proven otherwise. The flag gets cleared on first save. Matches the form's behavior of creating a temp row immediately on open.

- **Both `Booking` and `LogBooking` are denormalized** with full room+guest snapshots. So old receipts stay accurate even if the room is later renamed/repriced.

- **`ModuleAccess` has `Group1`–`Group4`** but only Groups 1 (Admin) and 4 (Clerk) are used in practice. The schema reserves slots for Manager/Supervisor that the UI never enables.

- **`Report` uses `MEMO` fields** for `ReportQuery` and `NullQuery` — Access's text-field cap of 255 chars wouldn't hold the actual SQL, so memo is required. `SubQuery` is plain TEXT (255 char limit) which can be tight for long ORDER BY clauses.

- **`LogError`** stores `LogModule` and `LogMethod` at 255 chars, `LogErrorNum` at 50 chars (so it can hold both numeric error codes and string codes), and `LogErrorDescription` as MEMO.

- **`WeeklyBooking`** is a pre-aggregated cache table for one specific report. It's the only table with no Active/CreatedDate audit columns — it's a working table, not a domain entity.

- **`Company.F01`** is a 50-char text column with no purpose evident from the codebase. Likely a "future field" placeholder for ad-hoc additions without schema migration. The `1` suffix on most Company columns (`ProductVersion`, `DatabaseVersion`) hints at multilingual support that was planned but never delivered — same `_1` suffix appears on `ModuleDesc1`, `ReportName1`, `ReportTitle1`, `ReportAsOn1`, `DateField1`, `DateType1` (suggests a `_2` for a second language was anticipated).

- **Every audit-trail table (`Booking`, `Room`, etc.) ends with the same five columns:** `Active`, `Temp` (where applicable), `CreatedDate DEFAULT NOW()`, `CreatedBy`, `LastModifiedDate`, `LastModifiedBy`. Consistent template.

After all twelve tables are created, `CloseDB` and return True. On error: catch, log, return False.

### `CreateSampleData() As Boolean`
Seeds the freshly-created database with starter content. Creates:

- One **Company** row: "STAR HOTEL", Kuala Lumpur address, telephone, MYR currency, DB version 1.3, system start date 1 Jan 2018
- 18 **ModuleAccess** rows defining the permission matrix for every form/report (Dashboard, Booking, List Report, Print Report, Export Report, Edit Report, Edit Report (Expert), Find Customer, Maintain Room, Maintain User, Access Control, Daily/Weekly/Monthly Booking Reports, Weekly Booking Graph, Shift Reports, Official Receipt). Pattern across rows: Group 1 (Admin) is always True; Groups 2/3 are always False; Group 4 (Clerk) gets True for the operational forms (Dashboard, Booking, List Report, Find Customer, Shift Report for User) and False for everything else.
- 7 **Report** rows with their full SQL queries (Daily/Weekly/Monthly Booking, Weekly Booking Chart, Shift by Staff, Shift by All Staff, Official Receipt). The SQL is concatenated multiline strings — this is where the report definitions live in code (and only here). Notable: each report has a `NullQuery` companion that returns the same column shape with all-null values — used to render an empty report when the main query returns no rows.
- 4 **UserGroup** rows: Administrator (level 99, active), Manager (level 98, **inactive**), Supervisor (level 20, **inactive**), Clerk (level 10, active). So the unused groups are *defined but inactive* — visible in code, hidden in UI.
- 4 **RoomType** rows: SINGLE BED ROOM, DOUBLE BED ROOM, TWIN BED ROOM, DORM. All active, all with empty `TypeLongName`.
- 1 **Room** row: ID=1, "101", Single bed, Level 1, MYR 100 with breakfast at MYR 10. The starter room.
- 2 **UserData** rows:
  - Admin / "Demo" / encrypted "admin" password / Group 1 / Idle 0 / Active
  - Clerk / "Receptionist" / encrypted "clerk" password / Group 4 / Idle 300 (5 minutes) / Active

Both seeded users have `ChangePassword = False` — they're not forced to change their password on first login. So after fresh install, `admin/admin` and `clerk/clerk` work immediately. This is consistent with the on-screen demo hint in `frmUserLogin`.

## Recordset opening helpers

Five variants exist for opening a recordset on the primary connection. The differences matter:

| Function | Cursor type | Lock type | Command type | Use case |
|---|---|---|---|---|
| `OpenSQL` | adOpenForwardOnly | adLockOptimistic | adCmdText | Fast read-once; can't `MovePrevious` or use `RecordCount` |
| `OpenQuery` | adOpenForwardOnly | adLockOptimistic | adCmdText | Identical to OpenSQL — duplicate? |
| `OpenRS` | adOpenStatic | adLockPessimistic | adCmdText | Bidirectional, scrollable, count works |
| `OpenTable` | adOpenDynamic | adLockPessimistic | adCmdTable | Whole-table open by name |
| `ExecuteSelectSQL` | configurable | adLockOptimistic | (text default) | Caller picks the cursor type |

**Why three forwards-only/optimistic variants?** `OpenSQL` and `OpenQuery` are functionally identical. The naming and the `'adCmdStoredProc` comment in `OpenQuery` suggest the original intent was for `OpenQuery` to call a stored procedure (Access query) versus `OpenSQL` for raw SQL — but Jet doesn't really distinguish, and the implementation collapsed to the same call. They both exist because forms picked one or the other arbitrarily over time.

**`OpenRS` is the right default for most form code.** `adOpenStatic` lets you scroll, count, and re-read; `adLockPessimistic` blocks concurrent writes. The form analyses showed `OpenRS` is the most-used helper (master/detail forms, report population, anything that needs `.RecordCount` or `.MoveFirst`).

**`OpenSQL` is used in authentication paths** (`frmAdmin`, `frmUserLogin`, `frmUserChangePassword`) — appropriate because they only need to read the user row once.

**`OpenTable`** opens an entire table by name with `adCmdTable` — used for bulk operations. Less common in the form code.

**`ExecuteSelectSQL`** is the most flexible — caller picks the cursor type. Probably an earlier-generation helper that the more-specialized ones replaced. Worth checking if it's still called anywhere in the codebase.

The parallel "Data" family (`OpenDataRS`, `OpenDataTable`) wraps `DAT` instead of `ACN`. Note `OpenDataRS` uses `adOpenStatic`/`adLockPessimistic` to match `OpenRS`, but there's **no `OpenDataSQL`** — only the static-cursor variant. So the migration code in `frmSplash` uses `OpenDataRS` exclusively when reading the source database.

## Recordset closing

### `CloseRS(rst)` / `CloseDataRS(rst)`
Defensive close, same pattern as `CloseDB`/`CloseData`. The reverse-conditional style (`If rst Is Nothing Then ... Else ...`) appears here too — five places in the module use this idiom.

These are called religiously across the codebase — every form's `CheckErr:` block ends with `CloseRS rst : CloseDB`. The discipline is real, which is good because VB6 doesn't have try/finally.

## Query execution

### `QuerySQL(pstrSQL, [recordsAffected])`
Wraps an `Execute` call in a transaction:
```vb
ACN.BeginTrans
ACN.Execute pstrSQL, plngRecordsAffected
ACN.CommitTrans
```

On error: `RollbackTrans`. So every UPDATE/INSERT/DELETE goes through a single-statement transaction. This is conservative — it gives atomicity per call but no batching. If you want to update 10 rows atomically, you have to call `ACN.BeginTrans` yourself before the loop and `CommitTrans` after.

The error path **logs to the database** (`LogErrorDB`) including the failed SQL — useful for debugging, since you can see exactly what query failed in the LogError table.

`QueryDataSQL` is the same wrapper for `DAT`. Used by migration code for writes to the destination database.

### `UnlockDB(Con)`
A curious utility:
```vb
OpenDB
gstrSQL = "SELECT * FROM UserData"
rst.Open gstrSQL, Con, adOpenForwardOnly, adLockOptimistic, adCmdText
rst.Close
CloseDB
```

It opens the database, reads UserData, immediately closes. The commented-out `'gstrSQL = "SELECT * FROM NOTHING"` suggests the intent was probably "perform a no-op query to refresh/unlock". In Jet, opening and closing a recordset can sometimes clear a stale lock state on the .mdb file — this looks like a defensive measure. The `Con` parameter is unused inside the function body (which uses `ACN` from `OpenDB`), suggesting an earlier signature took an arbitrary connection. The function is currently dead-ish — it works but the parameter is meaningless. Worth grep'ing to see if anyone calls it; if not, dead code.

## Notable points and quirks

- **The module is the canonical schema definition.** If the database schema needs to change, this is where it changes — and `frmSplash`'s `Update_Database` migration code needs a matching `ALTER TABLE` block to bring existing customer databases into sync. There's an implicit invariant that `CreateDB` and `Update_Database` produce the same schema for the same version — which is hard to verify because they're written separately.

- **Two ADO connections (`ACN` + `DAT`) drive the whole "Data" parallel API.** The architectural decision to support concurrent old/new connections enables the migration story. Without `DAT`, migrations would need temp files or a different process. This is the cleanest part of the design.

- **`gstrPassword` is a global that's set once at startup.** Whatever sets it (presumably `frmSplash` calling `gstrPassword = GenWord`) determines the password for every connection in the app's lifetime. Changing the database password mid-session would require updating the global and reconnecting.

- **`OpenSQL` and `OpenQuery` are functional duplicates.** Either is a candidate for removal in any cleanup pass. The form code uses both arbitrarily — `frmAdmin` uses `OpenSQL`, `frmFindCustomer` uses `OpenRS`, `frmUserMaintain` mostly uses `OpenRS` but with one `OpenSQL` call. There's no semantic distinction the calling code is honoring.

- **`adLockPessimistic` on read-mostly recordsets is conservative.** `OpenRS` uses pessimistic locking even though most callers only read. In single-user scenarios this is fine; in multi-user scenarios with concurrent reads/writes it could create unnecessary lock contention. For a single-machine Access app with one operator at a time, the choice is harmless.

- **`CreateSampleData` hardcodes the seed report SQL.** All 7 reports' queries live inline as VB6 string literals. Editing a report's *default* query (not its current configured query) means editing this function and recompiling. The `frmReportMaintain` form lets you edit the *current* query in the database, but if you ever wanted to factory-reset a report, you'd have to manually overwrite it with the literal here. The Reports Maintenance form has no "restore default" feature.

- **All commented-out `Debug.Print gstrSQL` lines** suggest the developer iteratively tested each table-create and seed-insert by uncommenting the print. Standard VB6 development pattern from the IDE-era.

- **`CompactDB` is fully commented out at the bottom.** It would have used `JRO.JetEngine.CompactDatabase` to compact a bloated .mdb file (Access databases grow with deletes/updates and need periodic compaction). The comment about `CrApplication.LogOffServer "P2smon.dll"` is interesting — `P2smon` is a Crystal Reports support DLL, suggesting an issue where Crystal kept a connection open that prevented compaction. The whole feature was disabled rather than worked around. So **the database has no maintenance/compaction routine**. Customers running this for years would see steady .mdb growth.

- **The `LogErrorDB` calls are commented out in some functions** (`OpenDB`, `ConnectDB`, `CloseDB`, `CreateData`, `CreateDB`, `CreateSampleData`, `OpenSQL`, `OpenQuery`, `OpenRS`, `OpenTable`, `CloseRS`). The pattern: log to text file via `LogErrorText` (always uncommented), but the database log call is suppressed. The reason is **circular dependency** — if `OpenDB` itself fails, you can't write to the LogError table because that requires an open connection. The text-file fallback lets you debug connection failures without needing the database to work. `QuerySQL` and `ExecuteSelectSQL` keep `LogErrorDB` enabled because they only run after a connection is open.

- **`UnlockDB`'s parameter is unused** — function takes `Con As ADODB.Connection` but uses `ACN` internally. Either dead parameter or a refactoring artifact. If the function is still called anywhere, the parameter signature is misleading.

- **No connection pooling, no retry logic.** Every `OpenDB` creates a fresh `ADODB.Connection`. For a single-user app this is fine; pooling matters more for web/multi-user scenarios.

- **`adCmdText` everywhere** means no use of stored Access queries (QueryDefs). All SQL is raw text. Combined with the no-parameterization noted in modCommon, the system makes zero use of Jet's query optimization features. For an Access app of this size, performance is fine — the overhead is negligible.

- **Manual `BeginTrans`/`CommitTrans`/`RollbackTrans` per query** in `QuerySQL` means every single write is wrapped in a transaction. The cost is small but real — for a bulk operation like the 18-row `ModuleAccess` seed in `CreateSampleData`, you'd get 18 separate transactions. A more efficient design would batch them, but for a one-time seed operation it doesn't matter.

- **The `SQLData_Text` calls in `CreateSampleData` for password fields** call `Encrypt("admin", mstrSalt)` directly, which goes through `GoldFishEncode` per modCommon. So the seeded passwords are properly hashed — they're not stored in plaintext.

- **Salt length is 6 in `CreateSampleData`** (`GenSalt(6)`) but 4 elsewhere (`GenSalt(4)` in `frmUserChangePassword` and `frmUserMaintain`). Inconsistency: the seed users get a stronger salt than user-created accounts. Probably an oversight — the seed code was written first with 6-char salts, and somewhere the figure got reduced to 4 for the Change Password flows without going back to update the seed.

- **No foreign key constraints** declared anywhere. The schema is referentially loose — `Booking.RoomID` references `Room.ID` by convention, but the database engine doesn't enforce it. So orphan booking rows are technically possible. The application code maintains integrity by always checking before delete (and using soft delete to avoid the issue entirely).

- **Module is named `modDatabase`** but doesn't include the SQL builder layer (which lives in `modCommon`). A cleaner separation would be: `modDatabase` for connection/recordset, `modSchema` for `CreateDB`/`CreateSampleData`, `modSQLBuilder` for the `SQL_*`/`SQLData_*` family. The current organization conflates "talk to the database" with "define the database" with "general utilities". For an app this size, the lumping is acceptable but it's the kind of thing you notice when navigating the codebase.

- **`Option Explicit` is set.** Standard hygiene.

- **`mstrModule` is private and consistent.** Every error-logging call uses `"modDatabase"` as the module name in the log table. Easy to filter the LogError table for connection-layer issues.
