# frmModuleAccess — Access Control Matrix

A grid-style permissions editor that maps modules (Booking, Reports, etc.) to user groups (Administrator, Clerk) using checkboxes. This is the F6-from-Dashboard form. Header note dates the auto-logout addition to 27/12/2014.

## Layout (top to bottom)

**Form shell**
- Same dark theme but `BackColor = #505050` (one shade lighter than the other forms — possibly an oversight, all the other main forms use `#303030`).
- Fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template as Dashboard/Booking/FindCustomer: `imgLogo` left, `lblBusinessName` (24pt) right-aligned, `lblCopyright` underneath, `shpCopyright` rectangle behind.

**Toolbar `tbrMenu`** — exceptionally minimal, just one button:
- CLOSE (Esc)
- `fraLogin` on the right with the live clock and `lblUserID` (same idle-watchdog pattern).

**Main body — `Picture1` containing the access matrix**

The grid is laid out manually with labels and checkbox controls:

*Header row (top, dark `#404040` background, white bold text):*
- `lblModule(0)` "Module" — wide column running across the form
- `lblGroup1` "Administrator"
- `lblGroup4` "Clerk"

*Two vertical separator lines* (`Line1` at X=3360, `Line2` at X=5400) divide module column from Administrator and Clerk columns.

*18 rows of modules* — each row is one `lblModule(i)` (full-width module name) plus one `chkGroup1(i)` checkbox (Admin column) plus one `chkGroup4(i)` checkbox (Clerk column). The modules listed in the form designer are:
- Rows 1, 2: "Module" (renamed at runtime from DB)
- Row 3: "Disabled Module" — checkbox is permanently `Enabled = False` for both groups
- Rows 4–10: "Module"
- Row 11: "Module" (after a small gap)
- Rows 12–17: "Report" (a dedicated reports section)
- Row 18: "Reserved" — placeholder for a future module

*Halfway down, a banner row:*
- `lblReport` (full-width, "Report")
- `Label1` "Administrator"
- `Label2` "Clerk"
This visually splits the matrix into a top "Module" half and a bottom "Report" half — exactly as referenced by `frmDashboard.UserAccessModule` calls (`MOD_REPORT_LIST` etc.).

**Administrator column quirk** — every Admin checkbox has `Enabled = False` and `Value = vbChecked`. So the Admin role is hardcoded as full access; nothing in this UI lets you remove permissions from Admins. Only the Clerk column is interactive. Row 3's Admin checkbox is the one exception — also disabled but *not* checked by default, because the row corresponds to a "Disabled Module" placeholder.

The Clerk column (`chkGroup4`) row 3 is also `Enabled = False`, matching the "Disabled Module" semantics.

## Module-level state

```vb
intTick   ' Idle-logout counter — same pattern as everywhere else
```

## Logic

**`Form_Load`**
1. Sets the title (`COMPANY_PRODUCT_NAME`), live clock, user ID label.
2. Calls `LoadModuleAccess`.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog → `frmDialog`).

**`Form_Unload`** — re-shows `frmDashboard`.

**`Form_KeyDown`** — Esc closes. That's the only shortcut.

**`tbrMenu_ButtonClick`** — only handles CLOSE.

**`LoadModuleAccess`** — reads the `ModuleAccess` table (one row per module) and populates the grid:

```sql
SELECT * FROM ModuleAccess WHERE Active = TRUE ORDER BY ModuleID
```

For each row where `ModuleID < 19`:
- Sets `lblModule(ModuleID).Caption = " " + ModuleDesc1` (note the leading space for visual padding — same as the form designer captions).
- Sets `chkGroup1(ModuleID).Value` from the `Group1` boolean column.
- Sets `chkGroup4(ModuleID).Value` from the `Group4` boolean column.

The `Group2` and `Group3` blocks are commented out — clearly the schema supports four groups but only Administrator (Group1) and Clerk (Group4) are wired up in the UI. Either the middle two groups are reserved for future use, or this app only ever needed two roles and the schema was over-designed early on.

The 19-row cap is hardcoded and matches the 18 designed rows (indices 1–18). Module ID 0 is implicitly the header label, not a real permission.

**`UpdateModuleAccess(moduleID, groupName, value)`** — generic single-cell writer:

```sql
UPDATE ModuleAccess SET <groupName> = <value> WHERE ModuleID = <moduleID>
```

`groupName` is passed in as a string (`"Group1"` or `"Group4"`) so a single helper updates the right column. No transaction batching — every checkbox click is one UPDATE statement.

**`chkGroup1_Click(Index)` / `chkGroup4_Click(Index)`** — the click handlers for the control arrays. Each one reads the new checkbox value and immediately calls `UpdateModuleAccess` to persist it.

The commented-out lines inside both handlers show an earlier UI iteration where the checkbox would also display "Yes"/"No" text and turn green/red — abandoned for the cleaner check-only style.

The `chkGroup2_Click` and `chkGroup3_Click` subs are entirely commented out, consistent with `LoadModuleAccess` not loading those columns.

## Notable points and quirks

- **Live persistence with no Save/Cancel.** Every checkbox click writes immediately to the database. There's no batching, no undo, no confirmation. For an admin tool this is fine — the operator sees the change instantly across the system — but it does mean an accidental click is permanent. A "Save Changes / Cancel" workflow would be safer.

- **Administrator permissions are immutable in the UI.** All Admin checkboxes are disabled, hardcoded checked. This is deliberate: the Admin role always has full access, and the form prevents anyone from accidentally locking Admins out of any module. Sensible policy, but it means this UI is really only for editing the Clerk role's permissions — calling it "Module Access" oversells what it actually does.

- **The "Disabled Module" row (index 3) is a placeholder.** Its checkboxes are disabled, its label is empty in the database (or it would be overwritten), and the form designer caption stays. So if you look at the running form you'd see a permanently grayed-out row. Probably reserved for a feature that was planned but never shipped.

- **Module 18 "Reserved" is similarly a placeholder.** Both its `chkGroup1(18)` and `chkGroup4(18)` are interactive, so the permission infrastructure is there — just no module bound to it yet.

- **Two-group design (Group1 + Group4).** The schema clearly supports four groups but the UI uses only the first and last. The skip from Group1 to Group4 (rather than Group1 to Group2) suggests Groups 2 and 3 were originally intended for "Manager" and "Supervisor" tiers between Admin and Clerk, then dropped. Worth flagging if the system ever needs an intermediate role — the database column already exists.

- **No save indicator.** Unlike `frmBooking` which fires "Booking is updated!" MsgBoxes, the commented-out `MsgBox "Module Access updated!"` line in `UpdateModuleAccess` was removed. This is the right call for a click-to-toggle UI (you'd be drowning in popups), but it also means there's no feedback that the write succeeded short of relying on errors being raised.

- **Permissions take effect immediately for the editing user too.** If an Admin opens this form and toggles a Clerk-row checkbox, the next time `frmDashboard` calls `UserAccessModule` the new value is read. So changes propagate live across all open forms. There's no per-session caching mentioned anywhere.

- **`SQL_SELECT_ALL` instead of explicit columns.** This is the only form so far that uses a `SELECT *` style helper. Most other code paths use explicit columns. Fine for this small admin table, but means a schema change adding a column won't break this code while it might break the explicit ones — different failure modes worth noting.

- **The `i = rst!ModuleID` (rather than auto-increment) lookup.** This means the database `ModuleID` values don't have to be contiguous — gaps are fine, and ordering by ModuleID just controls visual order. The design is robust to inserting new modules without renumbering.

- **The `ModuleAccess` table is presumably joined back to `MOD_*` constants** referenced in `frmDashboard` (`MOD_REPORT_LIST`, `MOD_FIND_CUSTOMER`, `MOD_MAINTAIN_ROOM`, `MOD_MAINTAIN_USER`, `MOD_ACCESS_CONTROL`, `MOD_BOOKING`, `MOD_BOOKING_VOID`). Each of those constants is presumably an integer matching a `ModuleID` in this table. The `UserAccessModule` global function is then doing the lookup `WHERE ModuleID = ? AND <user's group column> = TRUE`.

- **Self-lockout potential.** Nothing prevents an Admin from unchecking the Clerk's "Access Control" permission for themselves — but since they're logged in as Admin (Group1) and Group1 always has access, this isn't actually exploitable. The hardcoded Admin row is what makes this safe. If the system ever introduces another role with edit access here, that role *could* lock itself out.

- **`COMPANY_PRODUCT_NAME` is set in `Form_Load`** but not in some other forms (Booking sets it via designer caption). Mild inconsistency across the project.

- **The grid's middle "Report" banner row** (`lblReport` + Label1 + Label2 at Y=4440) is purely visual — no associated checkboxes. It just splits the table into a Module section above and a Report section below for readability. The actual report-permission rows are 12–17.

- **Row 11 sits visually above the "Report" banner** but is captioned "Module" — it bridges the gap between the module section (1–10) and the report section (12–17). Why the layout has this odd intermediate row isn't clear from the code; possibly it's the row for a "Reports List" toggle that turns the entire reports section on/off.
