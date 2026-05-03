# modTextFile — Text File I/O Module

The system's text-file read/write helpers. Versioned 1.2.22 with two recorded modifications: 01/10/2014 (the original module creation, version 2.0 internally — already at v2.0 by then, suggesting an even older v1.0 existed) and 17/12/2014 (adding `Log2File`). Like `modEncryption`, attributed to Aeric Poon. This is the file-I/O sibling to `modDatabase`'s database-I/O — same author, same era, same role of "abstract a Win32-era data layer behind named functions."

This module owns every text file the application touches: `Config.txt` (database location), `Error.txt` (text-fallback error log), and any other ad-hoc log files. The functions here are called by `frmDatabase` (writes Config), `frmSplash` and login flows (read Config), and every error handler in every module (write Error).

## What this module provides

Five functions/subs:

1. **`Log2File(path, text, [timestamp], [append])`** — generic logging with options
2. **`LogErrorText(filename, note, [error])`** — error logger to `App.Path\<filename>.txt`
3. **`WriteTextFile(fullpath, text)`** — overwrite a file (full path)
4. **`WriteFile(path, text)`** — overwrite a file (alternate version)
5. **`ReadTextFile(filename, lineNo, output)`** — read a specific line from `App.Path\<filename>.txt`
6. **`ReadFile(path, lineNo, output)`** — read a specific line from a full path

There's significant duplication — `WriteTextFile` and `WriteFile` do the same thing with different parameter names; `ReadTextFile` and `ReadFile` do the same thing with one taking a name (and prefixing `App.Path`) versus one taking a full path. The split represents two generations of the same API that coexisted without consolidation.

## Function-by-function

### `Log2File(strPath, pstrText, [blnAddTimeStamp], [blnAppend])`

The most flexible writer. Two optional flags:

```vb
FF = FreeFile
If blnAppend Then
    Open strPath For Append As #FF
Else
    Open strPath For Output As #FF
End If
If blnAddTimeStamp Then
    Print #FF, vbCrLf & FormatDateAndTime(Now) & vbCrLf & pstrText
Else
    Print #FF, vbCrLf & pstrText
End If
Close #FF
```

Notable design choices:

- **`FreeFile` returns the next available file handle number.** The standard VB6 idiom — instead of hardcoding `#1`, `#2`, etc., this asks the runtime for an unused number. Safer if the function is called from a context where another file is already open.
- **Append is the default.** `blnAppend As Boolean = True`. So calling `Log2File "log.txt", "message"` adds to the existing file. To overwrite, you pass `False` explicitly. Sensible default for a logger.
- **Timestamp is opt-in.** Default False. Callers must pass True to get the time prefix.
- **Always prepended with `vbCrLf`** — so every log line starts on a new line, even the first. This means a freshly created log file starts with a blank line before any content. Cosmetic quirk.
- **Error handler shows a MsgBox.** `MsgBox "Error #" & Err.Number, ..., App.Title`. So a write failure (disk full, file locked, permissions) interrupts the user with an alert. For a logger called from an error handler, this can chain — the logger fails, shows a MsgBox, the original error handler then shows its own MsgBox. Two dialogs from one root cause.

The 17/12/2014 modification note says this was added — probably to give callers a more flexible logger than the inflexible `LogErrorText`. In practice, looking across the form code, `Log2File` doesn't appear to be called anywhere — `LogErrorText` is the workhorse. So `Log2File` may be a "future" function that was added but never adopted.

### `LogErrorText(FileName, pstrNote, [pstrError])`

The actual error logger used everywhere. Three parameters: the base filename (without path or extension), a note describing the context, and an optional error description.

```vb
Open App.Path & "\" & FileName & ".txt" For Append As #1
If pstrError = "" Then
    Print #1, vbCrLf & FormatDateAndTime(Now) & vbCrLf & pstrNote
Else
    Print #1, vbCrLf & FormatDateAndTime(Now) & vbCrLf & pstrNote & vbCrLf & pstrError
End If
Close #1
```

Always opens in `App.Path` (the directory containing the .exe), always appends, always with timestamp, always to a `.txt` extension. So `LogErrorText "Error", "OpenDB", "Permission denied"` creates `App.Path\Error.txt` and writes:

```
[blank line]
30 Sep 2014 03:32:45 PM
OpenDB
Permission denied
```

Three or four lines per error event (depending on whether `pstrError` is provided). The blank line between entries makes the file scannable by humans.

**Hardcoded `As #1`.** Unlike `Log2File`, this one uses file handle 1 directly. If another file happens to also be using handle 1 (which would be unusual but possible), this would fail. The `FreeFile` idiom would be safer. Modest issue.

**Used as the universal text-fallback logger.** Every module's commented-out `'LogErrorText "Error", mstrMethod, Err.Description` and its uncommented `LogErrorDB ...` siblings show that the codebase consistently chose database logging when possible and falls back to this when database isn't available (notably from `modDatabase`'s connection helpers and from `modEncryption`'s low-level functions).

The error handler is the same MsgBox pattern as `Log2File`. Same chained-dialog risk.

### `WriteTextFile(pstrFileFullPath, pstrText)`

Overwrites a file with a single string of text. Used by `frmDatabase.cmdOK_Click` to write `Config.txt`:

```vb
WriteTextFile App.Path & "\Config.txt", strPath & vbCrLf & strFile
```

Implementation:

```vb
FF = FreeFile
Open pstrFileFullPath For Output As #FF
    Print #FF, CStr(pstrText)
Close #FF
```

**`Print #FF` adds a trailing newline.** So the file ends with an empty line. For Config.txt (read line-by-line by `ReadTextFile`), this is harmless — the trailing empty line is past the data.

**`CStr(pstrText)`** — the text is already a String, but the explicit conversion is defensive against any caller passing a Variant. No-op for String inputs.

**Error handler delegates to `LogErrorText`** rather than showing a MsgBox. Compare to `Log2File` (which MsgBoxes). The choice probably reflects "writes are more common, don't want to show an error dialog for every failed write" — with the assumption that `LogErrorText` itself won't fail. Reasonable if the failure is "Config.txt couldn't be written"; chained-failure-prone if the disk is genuinely broken.

### `WriteFile(strPath, strText)`

Functionally **identical** to `WriteTextFile`. Same body, different parameter names:

```vb
FF = FreeFile
Open strPath For Output As #FF
    Print #FF, CStr(strText)
Close #FF
```

The only difference is `pstrFileFullPath` → `strPath` and `pstrText` → `strText`. The `p` prefix in the older version follows a parameter-naming convention; the newer version dropped it. Same Hungarian for type but different conventions for purpose.

This is the kind of duplication that suggests the codebase was developed in stages: someone wrote `WriteTextFile`, later someone wrote `WriteFile` either not knowing the original existed or wanting a "cleaner" version. Both shipped.

### `ReadTextFile(FileName, LineNo, sOutput)`

Reads a specific line from `App.Path\<FileName>.txt` into the byref output parameter. Same logic as `ReadText` from modCommon — actually identical to it, suggesting two copies of the same function in two modules. Logic:

```vb
If LineNo < 0 Then
    Do Until EOF(FF) = True
        Input #FF, sOutput      ' read until end, last line wins
    Loop
ElseIf LineNo > 0 Then
    For i = 0 To LineNo         ' read up to LineNo, last read wins
        If Not EOF(FF) Then
            Input #FF, sOutput
        Else
            Exit For
        End If
    Next
Else                            ' LineNo = 0
    Input #FF, sOutput          ' read first line
End If
```

Three modes via overloaded LineNo:
- **Negative** → read until EOF, return last line
- **Zero** → return first line
- **Positive N** → return the (N+1)th line (note: `For i = 0 To N` runs N+1 times, so `LineNo = 1` returns the second line)

The off-by-one is **counterintuitive**. `LineNo = 0` returns line 1; `LineNo = 1` returns line 2. So `LineNo` is a zero-based "skip N lines and return the next one" rather than a one-based line number. Caller has to know this. In `frmSplash`, the calls are `ReadText "Config", 0, strPath` (line 1) and `ReadText "Config", 1, strFile` (line 2), which lines up with the off-by-one but isn't obvious from the call site.

**`Close` without a file handle.** The line `Close` (no `#FF`) closes *all* open files, not just the one this function opened. Slightly bug-prone if this function were ever called from a context with other open files. In practice the calling convention is always "open, read, close" with no nesting, so it works.

### `ReadFile(strPath, LineNo, strOutput)`

The full-path version of `ReadTextFile`. Identical logic, but takes a full file path instead of a name-relative-to-App.Path.

```vb
Open strPath For Input As #FF
```

Versus `ReadTextFile`'s:

```vb
Open App.Path & "\" & FileName & ".txt" For Input As #FF
```

So the function gives callers control over the path when they have files outside `App.Path` or with non-`.txt` extensions. Same off-by-one LineNo convention, same `Close`-without-handle bug.

## Notable points and quirks

- **The duplication of `Read`/`Write` between this module and modCommon's `ReadText`/`WriteText` is the most striking thing.** modCommon's `WriteText` writes to `App.Path\<filename>.txt` (always-append, always-with-timestamp, `Write` syntax with comma-separated values), and `ReadText` reads with the same off-by-one LineNo scheme. The two modules implement the same conceptual API differently. Most form code calls modCommon's versions for diagnostic logging from common functions and modTextFile's `LogErrorText` for error logging from form code. The split is "modCommon writes WHAT happened, modTextFile writes ERROR happened" — but the technical implementations are nearly identical.

- **`Print #FF` vs `Write #FF`.** This module uses `Print` exclusively; modCommon's `WriteText` uses `Write`. The difference matters: `Print` writes raw text with no special formatting; `Write` adds quotes around strings, formats dates with `#...#`, and uses commas as separators. So a file written by modCommon's `WriteText` is parse-friendly (machine-readable, like a CSV), while a file written by modTextFile's `Log2File` is human-friendly (free-form text). Choosing which to use depends on whether the file will be read by the application again or only by humans/external tools.

- **Off-by-one LineNo convention.** `LineNo = 0` returns line 1, `LineNo = 1` returns line 2. This is the kind of API decision that you either embrace and document or fix and break compatibility. The codebase embraces it — every caller knows the convention. But a new developer reading `ReadTextFile "Config", 1, strFile` would reasonably assume that's reading line 1.

- **Three modes via integer parameter overloading is bad API design.** "Negative = last line, zero = first line, positive = nth+1 line" is not discoverable. A modern API would be three functions (`ReadFirstLine`, `ReadLastLine`, `ReadNthLine`) or a clear enum. The current scheme works because it's used in only a few places.

- **`Close` without a file handle in both Read functions.** This closes all open files in the process, not just the one opened above. Defensive only if the function is called with no other files open. In a single-threaded VB6 desktop app, this is safe; in any other environment it's a bug. Worth replacing with `Close #FF`.

- **`Log2File` is uncalled in the codebase as far as the form analyses showed.** It exists, was added in the 17/12/2014 modification, and apparently no one rewired `LogErrorText` callers to use it. So the module has dead-ish code: `Log2File` provides options that no one uses. If the codebase were cleaned up, either the callers should be updated or the function should be removed.

- **Two writers, two readers, all files are line-oriented text.** No XML, no JSON, no INI parsing. The files this app uses are dirt-simple: one or two lines of plain text. `Config.txt` is two lines (path, filename). Error logs are several lines per entry separated by blanks. The simplicity is deliberate — these files are touched rarely and need no parsing complexity.

- **Hardcoded `As #1` in `LogErrorText`** is the only hardcoded handle in the module. All others use `FreeFile`. The inconsistency is small but real. If `LogErrorText` were ever called from within another file-I/O operation (a write that fails, calling `LogErrorText` to log the failure, while file #1 is held by the failing write), it would fail with a "file already open" error. The cascade would be silent — the LogErrorText call would just MsgBox.

- **`MsgBox` on file-I/O errors is questionable.** For `Log2File` and `LogErrorText`, popping a dialog box for a file-write failure means the user gets interrupted to learn that the system couldn't log an error. For a logger, silence-on-failure or fallback-to-Debug.Print would be more appropriate. The MsgBox pattern works because file failures are rare in practice, but the design is fragile.

- **`App.Title` is the MsgBox title.** Compare to other modules' `mstrModule` or `mstrMethod` titles. The use of `App.Title` (the project's product name) is friendlier — the user sees "Hotel Booking System" or whatever's set, not an obscure module name. Inconsistent with the rest of the codebase but better UX.

- **No newline option for writes.** `Print` always emits a newline at the end. There's no way to write a partial line without a trailing CRLF. For Config.txt (where each value is its own line), this is correct. For any future use case needing precise byte control, a different approach would be needed.

- **No file-locking, no concurrency safety.** `Append` mode can race if two processes try to write to the same file simultaneously. In a single-process desktop app this is fine; in any multi-process scenario the logs could interleave or lose entries.

- **Files default to ANSI encoding.** VB6's `Open ... For Output` uses the system's default codepage. So on a Malaysian Windows install, the encoding is whatever the system default is (probably 1252 or 1256). For ASCII-only content (timestamps, English error messages, file paths) this is invisible. For non-ASCII content (a guest name with accented characters), the file would be locale-dependent. Probably never matters for what the system actually logs.

- **`LineNo > 0` loop runs `For i = 0 To LineNo`.** That's `LineNo + 1` iterations. The intent was probably `For i = 1 To LineNo` (read N lines, last one wins), but written as `0 To N` it reads N+1 lines, returning the (N+1)th. This is the source of the off-by-one quirk — looks like a one-off bug that became the convention.

- **No cleanup function.** The module has no "rotate logs," "trim old entries," or "compress" function. Error.txt grows unboundedly over the lifetime of the application. For an installation that runs for years, this could become a multi-megabyte text file. No one's noticed because text files compress well in storage and the app's overall data volume is small.

- **`Modified On : 17/12/2014`** is the same date as `modDatabase`'s last modification. The 17 December 2014 commit added `Log2File` here and updated the database connection helpers there — likely a consolidated "improve error handling and logging" pass.

- **Original `Version : 2.0`** in the second header comment is striking — most modules version-bump in tiny increments (1.2.22 throughout). This module was already at 2.0 by 01/10/2014, suggesting it was the *second* generation of a text-file module after a v1.0 that was rewritten or replaced. The author's note "for handling Text file" suggests v1.0 was inadequate enough to warrant a clean restart.

- **`Option Explicit` is set.** Standard hygiene.

- **The module's job is small but ubiquitous.** Every error handler in the codebase potentially calls `LogErrorText`. The Config.txt round-trip (write at first run, read at every subsequent startup) goes through `WriteTextFile` and `ReadTextFile`. The text-file path is the application's escape hatch for everything that can't go through the database — startup configuration, pre-database errors, anything that needs to survive even when the .mdb is unavailable. That's a critical role for a 100-line module.

- **A modern replacement** would consolidate to one Read function and one Write function (each with full path), use `FileSystemObject` or .NET interop for richer file ops, add UTF-8 encoding support, add basic file-locking, and add log rotation. Total effort: ~50 lines of new code plus migration of all `LogErrorText` callers (which would be straightforward with a search-and-replace). For an internal tool the current module is adequate; for any productization it would be one of the easier modernization targets.
