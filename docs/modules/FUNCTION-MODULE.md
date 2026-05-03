# modFunction — Application Service Layer

The system's domain-helper module. Versioned 1.2.22, with two recorded modifications: 01/10/2014 (added `LogErrorDB`) and 04/01/2015 (extended `UserAccessModule` to support querying access for users other than the current one). This is the layer between the form code and the database — small, focused functions that answer common application-level questions ("does this user have access?", "what's the database version?", "is this user an admin?") and provide convenience wrappers for one-shot database operations.

If `modCommon` is "general utilities" and `modDatabase` is "talk to ADO," `modFunction` is "answer business questions." Every function here corresponds to a meaningful concept in the application's domain.

## What this module provides

Ten functions, grouped into four categories:

1. **System metadata** — `DB_Version`
2. **User identity & permissions** — `UserGroup`, `UserAccessModule`, `NeedChangePassword`, `AdminUser`
3. **Generic database shortcuts** — `QueryHasData`, `GetList`, `SelectField`, `UpdateField`
4. **Diagnostics** — `LogErrorDB`

The module is the highest-level database abstraction in the codebase. Form code that needs "the user's access permissions" doesn't query `ModuleAccess` directly — it calls `UserAccessModule(MOD_DASHBOARD)`. This keeps domain logic in one place.

## System metadata

### `DB_Version() As Double`

Reads `Company.DatabaseVersion` and returns it as a Double. Used by `frmSplash` to compare the file's schema version against the running application's version, and by `frmUserLogin` to display the DB version label.

```vb
strSQL = "SELECT DatabaseVersion FROM Company"
```

The `Company` table is single-row by convention (no WHERE clause needed), so this just grabs the first value. If empty (`rst.EOF Or rst.BOF`), returns 0 — which `frmSplash` uses to decide whether the database is uninitialized vs already migrated.

The local variable is declared `Dim dblVersion As String` but the function returns Double. VB6 implicit conversion handles it through `Val()`, but the type confusion is the kind of small bug that makes static analysis warn. `Val("1.3")` returns `1.3` as Double which then assigns into the String variable, gets converted back through the function return — works, but ugly.

### Commented-out `DB_Version_Update`

A simple `UPDATE Company SET DatabaseVersion = ?` writer, fully commented out. The migration code in `frmSplash.Migrate_Database` and `Update_Database` writes the version directly inline rather than going through this helper. Probably removed because it became trivial enough to inline. Dead code in the source file.

## User identity & permissions

### `UserGroup(strUserID) As Long`

Returns the numeric UserGroup for a given user ID. Returns 0 if user not found.

```vb
SELECT UserGroup FROM UserData WHERE UserID = '<strUserID>'
```

Notice the **direct string concatenation** — no `CheckInput` sanitization on `strUserID`. If a caller passed an attacker-controlled user ID with apostrophes, this would either break or be exploitable. In practice, the only callers of this function (looking at the form analyses) pass either `gstrUserID` (already validated at login) or hardcoded values, so it's safe by convention. But the function signature says "any string" while the implementation assumes "trusted string."

This function is essentially a building block — `AdminUser` and `UserAccessModule` both reimplement this same query inline rather than calling `UserGroup()`. Three places in the same module ask the database "what's this user's group" with three slightly different code paths.

### `UserAccessModule(intModuleID, [strUserID]) As Boolean`

The most-called function in the entire codebase. Every form's `Form_Load` and most toolbar button handlers consult this to decide whether to show/enable functionality. Returns True if the specified user (defaulting to the current user) has access to the specified module.

The 04/01/2015 modification added the optional `strUserID` parameter. Before that, it always checked the current user (`gstrUserID`); now you can ask "does Bob have access to Reports?" — useful for admin-management scenarios but rarely called with a non-default value in practice.

The function does **two queries**:

1. Read `ModuleAccess` row for the module → load `Group1`–`Group4` flags into a 4-element Boolean array
2. Read `UserData.UserGroup` for the user → return the matching group's flag

```vb
If rst!UserGroup = 1 Then UserAccessModule = blnGroup(1)
ElseIf rst!UserGroup = 2 Then UserAccessModule = blnGroup(2)
ElseIf rst!UserGroup = 3 Then UserAccessModule = blnGroup(3)
Else ' rst!UserGroup = 4
    UserAccessModule = blnGroup(4)
End If
```

The `Else` branch handles "anything that isn't 1, 2, or 3" — which means UserGroup values of 4, 5, 99, or anything else all map to Group4. Reasonable for a 4-group system, but **it means any unrecognized UserGroup silently inherits Clerk-level permissions** instead of being rejected. If a future migration introduced UserGroup 5 (Read-only Auditor) and forgot to update this function, those users would get Clerk access by accident.

A cleaner pattern would be `UserAccessModule = blnGroup(rst!UserGroup)` with a bounds check — but the array is `1 To 4` so out-of-range indices would crash. The if-else chain is defensive at the cost of being permissive.

**Two queries per call is also a perf concern** — every form load calls this several times (once per gated feature). For a form with 5 toolbar buttons that's 10 round-trips to the database. With local Access this is fast (microseconds), but it's a measurable contribution to form-load time. A caching layer would help; one doesn't exist.

The function returns False on any error or if the user is not found. So the security stance is **fail-closed**: if the database errors, no permissions granted. Right call.

### `NeedChangePassword(strUserID) As Boolean`

Reads the `ChangePassword` flag from `UserData`. Used by `frmUserLogin` to decide whether to redirect to the change-password screen after a successful login. Returns False if user not found (which means "if the user doesn't exist, don't force a password change" — an interesting fail-closed-ish stance, although the user shouldn't be reaching this code path if they don't exist).

Standard query pattern, same SQL-injection caveat (no `CheckInput`), same one-trip-to-database overhead.

### `AdminUser(strUserID) As Boolean`

Checks if a user is in Group 1. Just `UserGroup(user) = 1` factored into a separate function.

```vb
If rst!UserGroup = 1 Then AdminUser = True Else AdminUser = False
```

Could be a one-liner: `AdminUser = (rst!UserGroup = 1)`. The verbose form is consistent with the codebase's house style of explicit branches.

This function isn't called as much as you might expect from the form analyses — most permission checks go through `UserAccessModule(MOD_*)` rather than `AdminUser`. The exceptions would be places that want to short-circuit "admin can do anything" without checking a specific module — which the codebase mostly avoids in favor of the granular per-module checks.

## Generic database shortcuts

### `QueryHasData(pstrQuery) As Boolean`

Runs an arbitrary query and returns True if it produced any rows. Used by `frmRoomTypeMaintain.IsRoomTypeUsed` (which checks "are any rooms currently using this type?") to gate destructive operations.

```vb
If Not rst.RecordCount = 0 Then QueryHasData = True
```

**`RecordCount = 0` is unreliable on forward-only cursors.** `OpenRS` uses `adOpenStatic` which does support `RecordCount`, so this works — but if anyone changed `OpenRS` to forward-only, this function would silently return False even when rows existed. Implicit dependency on the cursor type chosen elsewhere in the codebase.

A more robust version would use `If Not rst.EOF Then QueryHasData = True` — works regardless of cursor type and is what most VB6 developers would write.

### `GetList(table, columnID, columnName, [condition], [columnOrder], [ascending]) As ADODB.Recordset`

Generates and executes a `SELECT id, name FROM table [WHERE condition] [ORDER BY column [DESC]]` query and returns the recordset. Used by combo-box population code throughout the system — `frmUserMaintain.PopulateGroupName` calls this to load user groups, `frmRoomMaintain` uses it for room types.

```vb
SELECT GroupID, GroupName, SecurityLevel FROM UserGroup WHERE Active = TRUE ORDER BY SecurityLevel
```

Notice that if `pstrColumnOrder` is non-empty, the column is added to the SELECT list **and** used in ORDER BY. This is unusual — most query builders treat ORDER BY columns separately from SELECT columns. The reason is presumably so callers can iterate through and access the order column if needed. But it adds a column to the result set the caller may not expect.

**Caller owns the recordset** and is responsible for calling `CloseRS` after use. The function doesn't close anything (which would invalidate the returned recordset). The error path closes early and returns Nothing — callers must check for Nothing before using.

This is the cleanest, most reusable function in the module. It's also the one with the most parameters (six!) — typical of the trade-off between "one helper for many cases" vs "many helpers for one case each."

### `SelectField(table, column, [condition]) As String`

Returns the first column of the first row from a query. Used for one-shot value lookups: "what's the room name for room ID 5?" → `SelectField("Room", "RoomShortName", "ID = 5")`.

```vb
SELECT <column> FROM <table> [WHERE <condition>]
```

**Two notable bugs:**

1. **Always returns String.** If the column is numeric or boolean, VB6 converts via implicit cast. For NULL values, this throws an error (caught by the error handler, which returns "").
2. **No CloseDB.** The function calls `OpenRS` (which assumes `OpenDB` was already called) — but never opens or closes the connection itself. So **callers must `OpenDB` before calling and `CloseDB` after.** This is inconsistent with `UserGroup`, `AdminUser`, etc., which manage the connection internally.

The error handler also doesn't close the connection. If a SelectField call fails mid-form, the DB connection stays open — combined with VB6's reference counting it usually gets cleaned up, but it's a leak waiting to manifest.

Looking at the function more carefully: the absence of `OpenDB`/`CloseDB` calls is **probably intentional** for performance — letting the caller batch multiple SelectField calls inside a single open connection. But the lack of documentation makes it a foot-gun. A new developer would assume it manages its own connection like the user-related functions do.

### `UpdateField(table, column, type, value, [condition]) As ADODB.Recordset`

A one-shot UPDATE helper:

```vb
UPDATE <table> SET <column> = <value> [WHERE <condition>]
```

Three notable behaviors:

**1. Type-aware quoting.** If `pstrType = "STRING"`, the value gets wrapped in apostrophes. Otherwise, it's emitted raw (suitable for numbers, booleans, expressions). This works but is less safe than the typed `SQL_SET_*` family in modCommon — there's no per-type validation.

**2. Magic string "INCREMENT_ONE".** If the value is the literal string "INCREMENT_ONE" and type is non-string, the SQL becomes `column = column + 1`:

```vb
If pstrValue = "INCREMENT_ONE" Then
    strSQL = strSQL & pstrColumn & " + 1"
```

This is what `frmUserLogin` uses to bump the LoginAttempts counter on wrong password. Functional but inelegant — a special-case string sentinel is the kind of API design that breaks when you add a second sentinel ("DECREMENT_ONE", "ADD_TEN"). A cleaner design would be a separate `IncrementField` function.

**3. Returns ADODB.Recordset but never assigns one.** The function signature is `As ADODB.Recordset` but the body never sets a return value. So all callers who treat the result as a recordset are working with Nothing. This is dead return type — should be `As Boolean` or just a Sub. Looking at the form code, callers use it as a Sub (`UpdateField "UserData", ...`) which works fine because VB6 is permissive about ignoring return values.

**4. Same connection-management asymmetry.** Like `SelectField`, this function uses `ACN` directly via `BeginTrans/Execute/CommitTrans` but doesn't `OpenDB` itself. Callers must have an open connection. Inconsistent with the read-side helpers above. The `frmUserLogin` and `frmAdmin` analyses noted that callers wrap each `UpdateField` with their own `OpenDB`/`CloseDB` — so the pattern is "inline each call in its own connection block."

## Diagnostics

### `LogErrorDB(logType, logModule, logMethod, errorNumber, [errorDescription])`

Inserts an error row into `LogError`. Called from every form's `CheckErr:` block as the last step before exiting. The pattern across the codebase:

```vb
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
```

The function builds the INSERT manually:

```sql
INSERT INTO LogError (LogDateTime, LogErrorNum, LogErrorDescription, LogUserName, LogModule, LogMethod, LogType)
VALUES (#date#, 'errnum', 'desc', 'user', 'module', 'method', 'type')
```

**Notice all values are wrapped through `CheckInput`** — the strongest sanitization function in the codebase. This is necessary because the error description may itself contain SQL fragments (failed query text, error messages mentioning column names) that would break the INSERT if not escaped. The caller doesn't have to think about escaping — `LogErrorDB` handles it.

The function **opens its own connection** (`OpenDB`) and closes it (`CloseDB`) — unlike `SelectField` and `UpdateField`, this one is self-contained. Reasonable choice — error logging shouldn't depend on caller state, since by the time you're calling it you might be in a partially-corrupted state already.

**The error handler is intentionally muted.** If `LogErrorDB` itself fails (database can't be opened, INSERT fails), it falls back to `LogErrorText` (file-based logging) without showing a MsgBox. The commented-out MsgBox makes this explicit:

```vb
'MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
LogErrorText "Error", mstrMethod, Err.Description
```

Right call. If the user has hit an error and the error logger itself fails, popping a second MsgBox would just be noise. The text-file fallback ensures the failure is at least captured somewhere.

`LogType` defaults to "Unknown" if blank. The four conventions used across the codebase: "Sub", "Function", "Property" (rare), and "Error" (used by the recursive fallback). So filtering the LogError table by LogType lets you distinguish between sub-vs-function errors and the rarer "the logger itself failed" cases.

## Notable points and quirks

- **The `_1` suffix on `ModuleAccess.ModuleDesc1`** matches the `_1` suffix on `Report.ReportName1`, `ReportTitle1`, `ReportAsOn1`, `DateField1`, `DateType1` from modDatabase. The unfulfilled multilingual story shows up here too — the function reads the row but doesn't reference `ModuleDesc1`, only the boolean group columns. So even if you populated `ModuleDesc2` (a hypothetical second language), this function wouldn't surface it.

- **`UserAccessModule` does two queries when one would do.** A JOIN between `UserData` and `ModuleAccess` could resolve the answer in a single round trip:
  ```sql
  SELECT M.Group1, M.Group2, M.Group3, M.Group4, U.UserGroup
  FROM ModuleAccess M, UserData U
  WHERE M.ModuleID = ? AND U.UserID = ?
  ```
  Doubling the query count for every permission check is the kind of inefficiency that adds up across a session. Caching the user's group at login (it's already in `gintUserGroup`) and only querying the ModuleAccess row would also work.

- **Actually `gintUserGroup` already exists** — it's set during `frmUserLogin`. So `UserAccessModule(MOD_X)` with the default user could check `gintUserGroup` directly without ever querying `UserData`. The code doesn't take this shortcut, presumably because the function was originally written to support arbitrary `strUserID` arguments (the 04/01/2015 modification). The unified path means even the common case pays for the extension.

- **`AdminUser` could be `gintUserGroup = 1`** when checking the current user. Same query-avoidance opportunity.

- **`SelectField` and `UpdateField` don't manage their own connection** while `UserGroup`, `AdminUser`, `NeedChangePassword`, `DB_Version`, `QueryHasData`, and `LogErrorDB` do. This split is the most surprising thing about the module. It means callers must remember which helpers self-manage and which require an outer `OpenDB`/`CloseDB`. There's no naming convention to indicate this.

- **No SQL injection sanitization** in any of the user-id-taking functions. `WHERE UserID = '" & strUserID & "'` is concatenated raw. Safe in practice because the only callers pass `gstrUserID` (which was sanitized at login) or hardcoded values, but the function signatures don't communicate this assumption. If a future caller passed an unsanitized input, they'd have a hole.

- **`UpdateField`'s "INCREMENT_ONE" sentinel** is a tell-tale sign of a refactor that didn't quite finish. The function was originally just "update field to value" and someone added increment-by-one as a special case. A cleaner version of the codebase would have an `IncrementField(table, column, [condition])` helper. The current pattern means anyone reading `UpdateField "UserData", "Loginattempts", "NUMBER", "INCREMENT_ONE", ...` has to know what that magic string does.

- **`UpdateField`'s `pstrType` parameter only meaningfully checks for "STRING".** Other values ("NUMBER", "BOOLEAN", "DATE") all go through the same non-quoted path. So callers passing "BOOLEAN" with value "FALSE" and callers passing "NUMBER" with value "0" are both handled identically. The parameter is documenting intent but not enforcing types — it's documentation-as-code, not validation.

- **Comments throughout the file say `'LogError "Error", mstrMethod, Err.Description`** with the `LogError` (text-file) call commented out. The pattern is "log to database (uncommented), don't double-log to text file (commented)." Different from `modDatabase` where text-file logging is *kept* for connection-layer code (because the database might be unavailable). At the `modFunction` layer, the assumption is "if you got here, the database is working, so logging to it is fine."

- **All functions return safe defaults on error.** False for booleans, 0 for numbers, "" for strings, Nothing for recordsets. The catch handlers consistently set the return value before exiting. This is the right pattern — callers can trust that "failed" looks like "no permission/no data" rather than crashing.

- **The error handlers all `CloseRS` before exiting** but `LogErrorDB`'s catch block doesn't `CloseDB` after the failed `OpenDB`. Inconsistent. Probably never causes a leak in practice because `LogErrorDB` is the last thing called before an exception bubbles up, and the connection state is moot if the form is about to unload.

- **`DB_Version` returns Double but the local variable is String.** The implicit type conversion works but makes the function harder to reason about. `Val()` does the heavy lifting — passing a non-numeric string to `Val()` returns 0, which is what we want for "uninitialized database". So there's some defensive intent in using a String temporary, but the type signature is misleading.

- **The module is named `modFunction`** but contains both Functions and one Sub (`LogErrorDB`). The naming convention across the modules — modDatabase, modCommon, modEncryption, modFunction — suggests the author thought of these as functional categories rather than strict typing. Reasonable for a single-author codebase.

- **`gstrUserID` is referenced inside `LogErrorDB`** to populate the `LogUserName` column. So error logs are automatically attributed to the current user. Useful for support — when a user reports "the system crashed," you can find their errors in the log filtered by username and time.

- **No `DeleteRecord`, `InsertRecord` equivalent** to `UpdateField`. The asymmetry suggests `UpdateField` was added because partial updates of single fields were common (login attempts, active flag); inserts and deletes are rare enough to justify writing the SQL inline each time.

- **`Modified On : 04/01/2015`** in the header is the latest modification date in any module we've seen so far — newer than `modCommon` (28/10/2014), `modDatabase` (17/12/2014), and `modEncryption` (01/10/2014). The 04/01/2015 change extended `UserAccessModule` to take a UserID parameter, presumably because some new feature needed to query permissions for users other than the current one. Looking at the form code, `frmAdmin` and `frmModuleAccess` would be candidates — but neither calls it with a non-default UserID in the analyzed code. So the extension was made for a use case that may not have been fully implemented, or that uses it in code we haven't seen yet.

- **`Option Explicit` is set.** Standard hygiene.

- **The module's responsibilities overlap with modDatabase.** `LogErrorDB` arguably belongs in modDatabase (it's about logging to the database) and `UpdateField`/`SelectField` are generic database helpers that match modDatabase's level of abstraction. The split between the modules is fuzzy — if anything, the title `modFunction` suggests "miscellaneous functions" which is what it became. A cleaner organization would put `DB_Version`, `UserGroup`, `UserAccessModule`, `NeedChangePassword`, `AdminUser` in a `modAuth` or `modAccess` module, and the generic helpers in modDatabase.
