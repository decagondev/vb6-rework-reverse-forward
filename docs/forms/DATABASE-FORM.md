# frmDatabase — Database File Picker

This is a small first-run / configuration dialog that lets the operator point the application at either a bundled demo Access database or a custom `.mdb` file location. It runs before the splash and login screens.

## Layout

**Form shell**
- `BorderStyle = 3` (Fixed Dialog) — no resize, no system menu icons in taskbar (`ShowInTaskbar = False`).
- 8175 × 3480 twips, dark gray (`#303030`) background, centered on owner.
- Smaller and simpler than the other forms — this is a one-off settings dialog.

**Left side**
- `imgIcon` — a 1095 × 1095 stretched icon image (decorative, top-left).
- `lblBookingID` — leftover from copy-paste, hidden, caption "0", clearly unused on this form.

**Right side — input area**
- Two label/textbox pairs sitting on top of `Shape2` and `Shape3` (the dark `#505050` rectangles that act as fake textbox borders since the textboxes themselves have `BorderStyle = 0`):
  - `lblLabels(0)` "File &Path :" → `txtFilePath`
  - `lblLabels(1)` "File &Name :" → `txtFileName`
- The ampersand makes Alt+P / Alt+N jump focus to the corresponding field.
- Below them: `chkDemo` checkbox + `lblBack` label "Use Demo database". The checkbox starts checked.

**Bottom-right buttons**
- `cmdOK` — "Save (Enter)", `Default = True`, so Enter triggers it.
- `cmdCancel` — "Cancel (Esc)", `Cancel = True`, so Esc triggers it.

## Logic

**`Form_Load`**
Reads the global `gstrDatabasePath`. If it currently points at the bundled `App.Path\Data\DemoData.mdb`, the form opens with the Demo checkbox ticked and the path/name textboxes disabled. Otherwise the checkbox is unticked and the textboxes are enabled for editing.

**`chkDemo_Click`**
Live-toggles the enabled state of `txtFilePath` and `txtFileName` to match the checkbox — so the user can't type a custom path while Demo is selected.

**`cmdCancel_Click`**
Unloads the form and shows `frmUserLogin`. So canceling here doesn't kill the app — it just falls through to login using whatever `gstrDatabasePath` was already set globally.

**`cmdOK_Click`** — the substantive part:

1. **If Demo mode**: checks that `App.Path\Data\DemoData.mdb` actually exists on disk. If missing, shows an error and aborts. Otherwise sets `strPath = App.Path\Data\` and `strFile = "DemoData.mdb"`.

2. **If custom mode**: validates that both `txtFilePath` and `txtFileName` are non-empty (with focus jumping to whichever is missing), then takes their trimmed values.

3. **Normalizes the path**: appends a trailing backslash if the user didn't type one.

4. **Persists the choice**: writes `path + CRLF + filename` into `App.Path\Config.txt` via the helper `WriteTextFile`. This is how the app remembers the database location between launches — a flat two-line text config file alongside the executable.

5. **Sets `gstrDatabasePath`** to the combined path+file.

6. **Branches on existence**:
   - If the file already exists at that path: unload, show `frmSplash` (presumably the splash then loads the existing database).
   - If not: runs three creation helpers in sequence —
     - `CreateData` — likely creates the `\Data\` folder structure if missing.
     - `CreateDB` — creates the empty `.mdb` file with all required tables (Booking, Room, Company, UserData, etc., based on what the other forms reference).
     - `CreateSampleData` — seeds it with starter rows.
   - Each helper returns Boolean and short-circuits with an error message if it fails.
   - On success: unload and show splash.

**Error handling**
Standard `On Error GoTo CheckErr` template — resets cursor, MsgBox, and logs via `LogErrorText` (text-file logger). Note this form uses `LogErrorText` rather than `LogErrorDB` — sensibly so, because at this point in startup the database might not exist yet, so DB-based logging would itself fail.

## Notable points

- **Two-line text config.** The persisted format is just `path\` then a newline then `filename.mdb`. Whatever loads `Config.txt` at startup must read those two lines and join them. Simple but means filenames with embedded newlines (no one does this, but…) would break it.

- **Demo vs custom is a one-way gate.** Once a custom path is saved, the user can come back here, tick Demo, and switch back. So it's reversible — but the previously typed custom path is lost unless they remembered it (the textboxes don't preserve their last value across sessions; only `gstrDatabasePath` does, and that gets overwritten on save).

- **No file picker.** The user has to type the path manually rather than browse — surprising UX for a 2019-added feature. Adding a `CommonDialog` browse button would have been a small lift.

- **No write-permission check.** If the user types a path that exists but isn't writable, `CreateDB` will be the one to fail, with whatever generic ADO error it surfaces — no upfront "this folder isn't writable" guidance.

- **`lblBookingID` is dead.** Hidden, captioned "0", clearly pasted in from `frmBooking` and never removed. Cosmetic only, but worth a cleanup.

- **`Cancel` doesn't actually cancel the database choice.** It just dismisses the dialog and falls through to login. If the user opened this dialog specifically to fix a broken database path, hitting Cancel sends them straight to a login screen that will then fail to connect. There's no "are you sure?" guard.

- **The three creation helpers (`CreateData`, `CreateDB`, `CreateSampleData`) are external** — defined in a standard module elsewhere in the project. Their behavior matters a lot here but isn't visible in this file.

- **First-launch flow probably looks like this**: app starts → reads `Config.txt` → if missing or path doesn't exist, shows `frmDatabase` → user picks Demo or custom → file created if needed → splash → login. On subsequent launches `Config.txt` exists, the path resolves, and this form is skipped entirely.
