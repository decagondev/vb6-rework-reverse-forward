# modCommon — Shared Utilities Module

The system's general-purpose helper module. This is the foundation that every form depends on — utility functions for string sanitization, type conversion, date formatting, and the SQL builder layer that wraps every database query in the application. Versioned 2.2.3, with the latest changes dated 19/12/2014.

## What this module provides

The module groups roughly six categories of helpers:

1. **Input sanitization** — `CheckString`, `CheckInput`
2. **Type conversion (null-safe)** — `ConvText`, `ConvInt`, `ConvLng`, `ConvDbl`, `ConvCur`
3. **Display formatting** — `FormatCurrency`, `FormatDate`, `FormatTime`, `FormatDateAndTime`, `FormatDateDayAndTime`, `FormatMonthYear`, `FormatYear`, `FormatDigit`, `FormatFixedID`, `FormatItem`, `GetItemID`
4. **Date arithmetic** — `WeekDay1`, `WeekDayN`, `WeekDay7`, `MonthDay1`, `MonthDay30`, `YearDay1`, `YearDay365`
5. **Control helpers** — `SetCombo`, `SetRecord`, `SetCheck`, `ResetControls`
6. **The SQL builder** — a large family of `SQL_*`, `SQLText`, and `SQLData_*` subs that mutate a global `gstrSQL` string
7. **Misc** — `GenWord`, `WriteText`, `ReadText`, `FileExists`, plus a `ShellExecute` declaration

There's also one small data-handling oddity (`GenWord`) that's actually load-bearing for security.

## Input sanitization

### `CheckString(strInput)`
Escapes single quotes (doubled) and strips potentially dangerous characters: double quotes, square brackets, asterisks, ampersands. Used when text is going into a SQL string literal where these characters would either break the query or interfere with `LIKE` searches (`*` and `[]` are wildcards in Jet SQL).

```vb
Replace(strInput, "'", "''")  ' SQL escape for apostrophes
Replace(strInput, """", "")    ' Remove double quotes
Replace(strInput, "[", "")     ' Strip Jet wildcard
Replace(strInput, "]", "")
Replace(strInput, "*", "")
Replace(strInput, "&", "")
```

The commented-out original line `Replace(strInput, "'", "")` (replacing apostrophes with nothing) was changed to the doubled-apostrophe pattern. Worth noting: legitimately useful punctuation in text fields gets stripped — a guest from "St. John's, Newfoundland" would have the apostrophes preserved (good) but the period stays (good), while a remark like "Note: see room *007*" would lose its asterisks. For a hotel app this is fine; it just means the data is sanitized to be Jet-SQL-safe.

### `CheckInput(strInput)`
Stronger sanitization for credentials and login-flow inputs. Strips reserved words ("password", "salt", "author", "code", "username", "select", "from") case-insensitively, then escapes apostrophes and removes double quotes.

```vb
Replace(strInput, "password", "", 1, -1, 1)  ' case-insensitive
Replace(strInput, "salt", "", 1, -1, 1)
Replace(strInput, "select", "", 1, -1, 1)
Replace(strInput, "from", "", 1, -1, 1)
' ...
```

The intent is anti-injection — the comment "Try to key in 1=1'or'1=1 as password to hack" makes the threat model explicit. The implementation has real problems though:

- **The reserved-word stripping is naive.** A user named "Pselectter" would have "select" stripped to become "Pter". A user from "Frome, Somerset" would lose "from" and become "e, Somerset". Real-world data with these substrings (rare but possible) gets corrupted before it ever hits the database.
- **It's defense-in-depth, not primary defense.** The actual SQL safety comes from the typed builder helpers below (where text values get wrapped in escaped quotes). The strip-words pass is hopeful — `1=1'or'1=1` doesn't contain any of the stripped keywords, so it slides through. The apostrophe doubling at the end is what actually neutralizes that specific attack.

This is consistent with what we've seen across the forms: `CheckInput` is used primarily on User IDs at login and admin authentication, where the risk is higher and the data shape is more constrained (uppercase, ≤10 chars, alphanumeric in practice).

## Type conversion

Five small wrappers (`ConvText`, `ConvInt`, `ConvLng`, `ConvDbl`, `ConvCur`) that handle the common database-read pattern: take a value that might be Null, empty string, or a malformed number, and return a safe typed default.

```vb
Public Function ConvInt(pstrInput As String) As Integer
    If Trim(pstrInput) = "" Then
        ConvInt = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvInt = CInt(pstrInput)
        Else
            ConvInt = 0
        End If
    End If
End Function
```

All five follow the same template: empty → 0 (or empty string), non-numeric → 0, otherwise convert. This is the pattern that makes `If rst!Field <> "" Then ... Else ...` blocks across the form code possible without crashing on database NULLs.

`ConvDbl` and `ConvCur` additionally check `IsNull(pstrInput)`. The other three don't, but in practice `Trim(Null) = ""` returns True for those cases so the empty-string branch catches it.

A quirk: the parameter is typed as `String` even though Variant would be more accurate — VB6 implicit conversion handles ADO field values reasonably, but if you ever passed a true `Null` Variant you'd hit an "Invalid use of Null" error before any of the function's defensive code ran. In practice the form code wraps these calls around `rst!Field.Value` which returns a Variant, and VB6's implicit String conversion handles it.

## Display formatting

### Currency
`FormatCurrency(value, [decimals])` wraps `FormatNumber` with thousands separators and 2-decimal default. The standard VB6 `FormatCurrency` function exists too, but this one is named the same and effectively shadows it within the module's scope — slight foot-gun if anyone tries to use the built-in.

### Dates and times
A consistent set of date formats used throughout the application:

| Function | Format string | Example output |
|---|---|---|
| `FormatDate` | `dd MMM yyyy` | `30 Sep 2014` |
| `FormatTime` | `hh:mm AMPM` | `03:32 PM` |
| `FormatTimeSeconds` | `hh:mm:ss AMPM` | `03:32:45 PM` |
| `FormatDateAndTime` | `dd MMM yyyy hh:mm:ss AMPM` | `30 Sep 2014 03:32:45 PM` |
| `FormatDateDayAndTime` | `DDDD, dd MMMM yyyy     hh:mm AMPM` | `Tuesday, 30 September 2014     03:32 PM` |
| `FormatMonthYear` | `MMM yyyy` | `Sep 2014` |
| `FormatYear` | `yyyy` | `2014` |

Note `FormatDateDayAndTime` uses five spaces between the date and time — the visible separator in the system clock displays. `FormatTimeSeconds` takes an optional AMPM string so callers can suppress it.

### ID formatting
`FormatDigit(lngID, intDigit)` left-pads a Long ID with zeros to a fixed digit count (e.g., 12345 → "012345"). Used across the system for booking number display ("Booking No: 100001"). The `Format$(lngTransID, "0000000")` in the comment shows the simpler equivalent that VB6's built-in supports — the manual implementation predates someone discovering that.

`FormatFixedID(strID, intLength)` right-pads with spaces to fixed length. Less commonly used.

### Item formatting
`FormatItem(strItemID, strItemName, [blnActive])` and `GetItemID(strItem)` are a paired encoder/decoder for combo-box-style "ID + Name" strings using a tab separator. Active items are formatted plain; inactive items get an "(X)" suffix on the ID. The commented `"�"` (a single character that didn't survive encoding) suggests the marker was originally a Unicode glyph that got replaced with ASCII for portability.

## Date arithmetic

Seven functions that compute boundary dates relative to a given date — used by `frmReport.PrintReport` for date-range filtering.

| Function | Returns |
|---|---|
| `WeekDay1(date)` | The Sunday at or before `date` |
| `WeekDayN(date, n)` | The Nth day (1=Sun, 7=Sat) of `date`'s week |
| `WeekDay7(date)` | The Saturday at or after `date` |
| `MonthDay1(date)` | First day of `date`'s month |
| `MonthDay30(date)` | Last day of `date`'s month (despite the name) |
| `YearDay1(date)` | January 1 of `date`'s year |
| `YearDay365(date)` | December 31 of `date`'s year |

Two important notes:

**`MonthDay30` is correctly named misleadingly.** Despite the "30" in the name, the implementation computes "1st of next month, minus 1 day" — which gives the actual last day of the month (28, 29, 30, or 31 depending). So January's last day is 31, February's is 28 or 29, and so on. The name is wrong but the implementation is right. The form analysis flagged this earlier as a potential bug — looking at the actual code, it isn't a bug, the function is just misnamed. Worth correcting the name in any refactor to avoid confusion.

**Week starts on Sunday.** `WeekDay1` returns Sunday, `WeekDay7` returns Saturday. So a "weekly" report run on a Wednesday covers the Sunday-to-Saturday span containing that Wednesday. This is the US convention; in many other regions weeks start Monday. For a Malaysian hotel system either convention works internally as long as it's consistent — and it is.

**Date arithmetic via string concatenation.** Most of these functions build a date by concatenating a day number, month name, and year into a string, then `DateAdd("D", 0, strTemp)` to coerce it to a Date. `DateAdd("D", 0, ...)` is a no-op offset — the actual conversion happens in the implicit string-to-date parsing. This works because `MonthName(n, True)` produces the locale-abbreviated name ("Jan", "Feb", etc.) and the resulting "1 Jan 2014" string is unambiguous to VB6's date parser. Cleaner alternatives like `DateSerial(year, month, day)` exist but weren't used.

## Control helpers

### `SetCombo(cbo, strText, [intItemData])`
Selects a combo item by either matching `ItemData` (numeric ID) or `List` (visible text). If neither matches, falls back to first item. Cleaner than inline loop in caller code. Apparently underused — the form code mostly does the loop manually.

### `SetRecord(txtOutput, adoField, [blnTrim])`
Reads an ADO field into a TextBox, with empty-string fallback. Used by `frmReportMaintain.PopulateValues` and elsewhere where the form pulls existing record data into the editor.

### `SetCheck(chkOutput, adoField)`
Reads a Boolean ADO field into a CheckBox.

### `ResetControls(frm)`
Iterates all controls on a form, clearing TextBoxes and resetting ComboBoxes to no selection. A reusable Form-wide reset, though most forms have their own per-field `ResetFields` instead so they can apply specific defaults.

The commented-out `SetString` and `SetNumeric` show abandoned earlier versions — `SetRecord` is what `SetString` became.

## The SQL builder layer

This is the biggest section of the module — a **fluent-style SQL builder** built around mutating a global string variable `gstrSQL`. Every form's database operations go through these helpers.

### How it works

The pattern is: form code calls a sequence of subs, each appending to `gstrSQL`. A typical SELECT:

```vb
SQL_SELECT
SQLText "ID"
SQLText "UserID"
SQLText "UserName", False    ' last column, no trailing comma
SQL_FROM "UserData"
SQL_WHERE_Long "ID", 5
```

Builds: `SELECT ID, UserID, UserName FROM UserData WHERE ID = 5`

Each helper takes a typed value and emits the right SQL syntax — text gets quoted, dates get hash-delimited (`#date#`, Jet syntax), Booleans become `TRUE`/`FALSE`, numbers go raw.

### The categories

**SELECT family:**
- `SQL_SELECT()` — starts a `SELECT` clause
- `SQL_SELECT_ALL(table)` — `SELECT * FROM table`
- `SQL_SELECT_TOP(field, table, [n])` — `SELECT TOP n field FROM table`
- `SQL_FROM(table, [alias])` — appends ` FROM table [alias]`

**Joins:**
- `SQL_INNER_JOIN(table, [alias])`
- `SQL_LEFT_JOIN(table, [alias])`
- `SQL_ON(alias1, field1, alias2, field2)` — emits the join condition

The 28/10/2014 modification noted at the top of the module simplified these — earlier versions took the table aliases and field names as one combined call. Now you build it in two steps (join + on), which is more flexible.

**WHERE clauses (typed):**
- `SQL_WHERE_Text(field, value)` — quotes the value
- `SQL_WHERE_Long(field, value)` / `SQL_WHERE_Integer(field, value)`
- `SQL_WHERE_Boolean(field, value)`
- `SQL_WHERE_BETWEEN(field, left, right)`
- `SQL_WHERE_LIKE_Text(field, value)` — adds `%...%` wildcards
- `SQL_AND_LIKE_Text(field, value)`, `SQL_OR_LIKE_Text(field, value)` — for chaining LIKE conditions

**ORDER BY:**
- `SQL_ORDER_BY(field, [ascending])` — appends DESC if false

**INSERT family:**
- `SQL_INSERT(table)` — starts `INSERT INTO table (`
- `SQL_VALUES()` — closes column list, opens values: `) VALUES (`
- `SQL_Close_Bracket()` — closes the final paren

Used as: `SQL_INSERT` → `SQLText "Col1"` × N (last with `False` for no comma) → `SQL_VALUES` → `SQLData_*` × N → `SQL_Close_Bracket`.

**UPDATE family:**
- `SQL_UPDATE(table)` — starts `UPDATE table SET`
- `SQL_SET_Text(field, value, [endComma])` — emits ` field = 'value',`
- `SQL_SET_Double`, `SQL_SET_Integer`, `SQL_SET_Long` — for numeric types
- `SQL_SET_Boolean(field, value, [endComma])` — emits `TRUE`/`FALSE`
- `SQL_SET_DateTime(field, dateTimeStr, [endComma])` — wraps in `#...#`

The comma management is **the form code's responsibility** — every helper takes a `blnEndComma` flag (default True) and the *last* call in a sequence must pass `False` to suppress the trailing comma. Get this wrong and your SQL has a syntax error.

**DELETE / DROP / ALTER:**
- `SQL_DELETE(table)` — `DELETE FROM table`
- `SQL_DROP(table)` — `DROP TABLE [table]`
- `SQL_ALTER_TABLE(table)` — for migrations

**CREATE TABLE family:**
- `SQL_CREATE(table, [prefix])` — opens the column list
- Column definitions: `SQL_COLUMN_ID`, `SQL_COLUMN_TEXT`, `SQL_COLUMN_MEMO`, `SQL_COLUMN_NUMBER` (with field-size selector for BYTE/SHORT/LONG/SINGLE/DOUBLE/CURRENCY/GUID/DECIMAL), `SQL_COLUMN_BIT`, `SQL_COLUMN_YESNO`, `SQL_COLUMN_DATETIME`

The column helpers take optional `strDefault`, `blnNullable`, `blnEndComma` — they're parameterized enough to handle most schema-creation cases. The comment "NOT yet used or tested" on `SQL_COLUMN_NUMBER` suggests the schema was built once with the simpler helpers and the more flexible Number variant was added later as future-proofing.

**Generic escape hatches:**
- `SQLText(text, [endComma], [beginSpace])` — emit literal SQL fragment
- `SQL_Comma()` — append a comma
- `SQL_Close_Bracket()` — append `)`

These are used when none of the typed helpers fit — like the `SQLText "AND R.ID = " & intRoomID, False, True` pattern seen in `frmDashboard`.

**INSERT VALUES (typed):**
- `SQLData_Text` (runs through `CheckString` for safety!)
- `SQLData_Double`, `SQLData_Long`, `SQLData_Integer`
- `SQLData_Boolean`, `SQLData_DateTime`

Note the **asymmetry**: `SQLData_Text` calls `CheckString` automatically, but `SQL_SET_Text` does **not** sanitize. So inserts are safer than updates by default — relying on the caller to pre-sanitize update values. Across the form code, callers usually do trim and sometimes call `CheckString`/`CheckInput` explicitly, but it's inconsistent. This matches what we saw in `frmRoomTypeMaintain.SaveRoomType` where INSERT used `CheckInput` but UPDATE didn't.

### Critical observations about the SQL layer

**Global state.** `gstrSQL` is a module-level public string declared elsewhere (presumably in a `globals` module). Every helper mutates it. Two consequences:
- You cannot build two queries in parallel from the same thread — they'd interleave into garbage.
- Calling `OpenDB` / `OpenRS(gstrSQL)` reads the current state — which is why every form follows the build-then-execute-then-close pattern strictly.
- The migration code in `frmSplash.Migrate_Database` opens a *second* connection (`OpenData` / `QueryDataSQL`) that uses its own SQL string, because the main `gstrSQL` is being used for the source-database read.

**Comma management is manual and fragile.** The `blnEndComma` parameter on every helper means the caller has to know which call is "last." Typical buggy patterns:
- Forgetting to set `False` on the last column → trailing comma → SQL syntax error
- Setting `False` on a middle column → missing comma → SQL syntax error
- Splitting a save into conditional blocks where the "last" column varies → harder to track

A more modern design would build a list of column-value pairs and have the executor emit the commas automatically. The current design trades that abstraction for verbosity and possible foot-guns.

**No parameterized queries.** Every value is concatenated into the SQL string. Combined with `CheckString` and `CheckInput` providing escape-by-stripping, the application is **not using prepared statements**. This means:
- SQL injection is possible if values bypass sanitization
- Performance is suboptimal (no query plan caching at the Jet level, though for Access this matters less than for SQL Server)
- Date formatting depends on locale — the system relies on `FormatDateAndTime`'s `dd MMM yyyy` format being parsed unambiguously by Jet

For an internal hotel app this is acceptable; for any system with untrusted input it's a significant gap. Migrating to ADO `Command` parameters would be a substantial refactor since every form goes through this layer.

## Misc utilities

### `GenWord()` — the database password generator

A small but consequential function:

```vb
intArray(0) = 67   ' C
intArray(1) = 111  ' o
intArray(2) = 109  ' m
intArray(3) = 112  ' p
intArray(4) = 117  ' u
intArray(5) = 116  ' t
intArray(6) = 101  ' e
intArray(7) = 114  ' r
intArray(8) = 105  ' i
intArray(9) = 115  ' s
intArray(10) = 101 ' e
```

It builds the string "Computerise" (the company name from "Computerise System Solutions") character-by-character via ASCII codes. This is what's used as the **Jet OLE DB database password** in `frmSplash` (`gstrPassword = GenWord`) and `frmPrint.OpenRDB`.

The obfuscation goal is to keep the literal string "Computerise" out of the executable's strings table, where a casual peek with strings(1) or a hex editor would find it. Building it from int values means the password as a string never appears in the binary — only the integers do, and an attacker has to either run the code or recognize the pattern.

This is **security through obscurity** in the textbook sense. A determined attacker decompiling the VB6 binary will find this function, recognize the int-to-char pattern, and recover the password in seconds. But it does block casual inspection. For an .mdb password protecting against opportunistic snooping, this is roughly the right level of effort. Notably, the password is the **company name** — anyone who knows the developer could probably guess it on a few tries.

### `WriteText(filename, [note])`
Appends an error log line to `App.Path\<filename>.txt`. Format: timestamp, error, optional note. Used by `LogErrorText` callers when database logging is unavailable (notably from `frmDatabase`'s pre-DB error paths).

### `ReadText(filename, lineNo, sOutput)`
Reads a specific line from `App.Path\<filename>.txt`. The interface is unusual — `sOutput` is `ByRef` for the return value, lineNo can be -1 (last line), 0 (first line), or N (Nth line). Used by `frmSplash` to read `Config.txt` (path on line 0, filename on line 1).

The error handler `NewFile:` writes to the error log if the file doesn't exist — nominally creating evidence of the missing file rather than the read itself succeeding.

### `FileExists(strPath)`
Tiny helper using `Dir$()` and `Len()`. Returns False on any error or empty result. The standard idiom in VB6 — there's no built-in `File.Exists` on this version.

### `ShellExecute` declaration
At the top of the module, the Win32 `ShellExecute` API is declared. Only needed if the code launches external programs (like opening a help file or URL in the default browser) — not used in any of the forms we've seen but available globally if needed.

## Notable points and quirks

- **Module-level globals as the architecture.** `gstrSQL` is the central piece of shared state, and nearly every database operation in the application reads/writes it. This is the VB6-of-its-era design pattern; it's fragile but consistent.

- **`CheckInput`'s reserved-word stripping is misguided.** Stripping "select", "from", "password", "salt" is well-intended but corrupts legitimate data and provides no real injection protection on its own. The actual safety comes from the apostrophe doubling at the end. A future refactor should remove the substring stripping and rely on parameterized queries instead.

- **`MonthDay30` is correctly implemented despite the name.** The form-analysis pass flagged this as a possible "missing last day of month" bug in `frmReport`. Looking at the actual implementation here, it correctly returns the last day of any month (28/29/30/31). The name is misleading but the math is right. Earlier analysis was wrong on this point.

- **`CheckString` strips ampersands.** For label captions in forms (where `&` is the accelerator marker) this would be problematic, but the function is for SQL values, not display strings. Ampersands in guest names would be lost — "Lee & Associates" becomes "Lee  Associates" with two spaces. Minor data quality issue.

- **Date format is locale-dependent.** `FormatDate` produces `dd MMM yyyy` like `30 Sep 2014`. Jet parses this in any locale because the month is text. But `FormatDateAndTime` uses the same format — so date arithmetic in Jet via `#30 Sep 2014#` works reliably. In a Malaysian locale specifically, this should always produce English month names because `MonthName(n, True)` returns abbreviated English names regardless of locale unless the system locale has overridden them. Worth verifying on a fresh install.

- **`SetRecord`'s default `blnTrim = True`** but the form code sometimes passes `False` explicitly when the field needs preserved whitespace (rare).

- **Five conversion functions, four are similar.** `ConvText`/`ConvInt`/`ConvLng`/`ConvDbl`/`ConvCur` all share the same template. A more modern codebase would generic them; in VB6 the explicit per-type functions are the idiomatic choice.

- **The build-by-mutation SQL pattern is the codebase's defining characteristic.** Every form's database code goes through `gstrSQL` and the helpers here. The cleanliness of individual form code depends entirely on whether the developer remembered to set `blnEndComma = False` in the right places. This is the single biggest source of fragility in the codebase, and the single thing a modern refactor would replace first.

- **`SQLData_Text` sanitizes; `SQL_SET_Text` doesn't.** This asymmetry surfaces across forms — INSERT paths are inherently safer than UPDATE paths because INSERT calls automatically run text through `CheckString`. Patching this would mean adding `CheckString` to `SQL_SET_Text` (with the consequence that any caller already calling `CheckString` would double-escape).

- **`GenWord` proves the security model is "make casual snooping harder."** Building "Computerise" from ASCII integers is exactly the kind of obfuscation that meets its design goal — a quick `strings | grep -i password` won't find it — without pretending to be cryptographic. Combined with the hand-rolled `GoldFishEncode` for user passwords and the on-screen advertisement of `admin/admin` defaults, the system is built for a low-threat internal environment. Honest about its level.

- **`ReadText` and `WriteText`'s file numbers (`#1`, `#2`) are hardcoded.** If two callers are reading and writing simultaneously they'd clash — but VB6 single-threaded apps don't hit this. Still, it's the kind of code that breaks under any kind of concurrency change.

- **No corresponding `WriteTextFile` here.** `frmDatabase` uses `WriteTextFile` (different name) to write `Config.txt`, suggesting that helper is in another module. So file I/O isn't fully consolidated in this module.

- **Module is named `modCommon`** but contains roughly equal parts general utilities and SQL builder. A more idiomatic split would be `modUtils` (conversions, formatting, dates) plus `modSQL` (the SQL builder) — the SQL section is ~60% of the file by line count. Combined-module is fine for an app this size.

- **`Option Explicit` is set.** Variable declarations are required throughout. Good hygiene, prevents typo-as-new-variable bugs.

- **Most functions don't have error handlers.** `On Error GoTo CheckErr` is conspicuously absent here, unlike form code. The reasoning: these helpers are pure/deterministic and shouldn't fail. `WriteText` and `ReadText` are the exceptions because file I/O can. Reasonable choice.

- **The version comment block at the top documents three iterations** (28/10, 02/11, 19/12 of 2014). The module stabilized fast — most core functionality was right by November 2014, with `ConvText` being the last addition before the codebase locked in.
