# frmRoomTypeMaintain — Room Type CRUD Editor

A small companion form to `frmRoomMaintain`. It manages the lookup table that drives the "Room Type" dropdown — categories like "DORM", "DELUXE", "STANDARD" etc. that get assigned to individual rooms. Reachable only from `frmRoomMaintain` via the Type (Ctrl+T) toolbar button.

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template: `imgLogo` left, `lblBusinessName` (24pt, right-aligned), `lblCopyright` underneath, `shpCopyright` rectangle.

**Toolbar `tbrMenu`** — 5 buttons matching the same pattern as `frmRoomMaintain`:
- CLOSE (Esc)
- CLEAR (Ctrl+C) — blank fields
- RESET (Ctrl+R) — re-read selected row
- SAVE (Ctrl+S)
- DELETE (Ctrl+D) — *defined but `Enabled = False, Visible = False`*; the delete operation is intentionally disabled

`fraLogin` on the right with the live clock and `lblUserID`.

**Main body — two-panel master/detail**

*Left — `Picture1`:*
- `lvRoomTypes` ListView: hidden ID, "Short Description", "Long Description"
- Color-coded: gray for active, pink for inactive

*Right — `Picture2`:*
- `lblRoomTypeID` — read-only display of the type's ID (label styled as a textbox, "0" default)
- `txtTypeShortName *` — short code (MaxLength 30, what `frmRoomMaintain.cboRoomType` actually displays)
- `txtTypeLongName` — full description (MaxLength 255)
- `chkActive` — Yes/No, label colored pink (matching the soft-delete pattern)

The Active label being pink (`#8080FF`) is consistent with the rest of the project's "Active false = pink" convention.

## Module-level state

```vb
intTick   ' Idle-logout counter
COL_GRAY = #E0E0E0
COL_PINK = #8080FF
```

## Logic

**`Form_Load`** — sets the live clock, user ID label, and calls `ListRoomType`. The commented-out `lvRoomTypes_Click` at the end is interesting: `ListRoomType` already calls it internally, so the duplication was removed.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog → `frmDialog`).

**`Form_Unload`** — calls `frmRoomMaintain.Form_Load` (note: explicitly invoking another form's `Form_Load` event handler — VB6 lets you, but it's unusual) then re-shows `frmRoomMaintain`. The reason: `frmRoomMaintain` populates `cboRoomType` only at load time, so to make any newly-added or renamed types visible, the parent has to re-run its initialization. Calling `Form_Load` directly is a quick way to do that without exposing a public `RefreshRoomTypes` method.

**`Form_KeyDown`** — Esc closes, Ctrl+C clears, Ctrl+R resets, Ctrl+S saves. Ctrl+D for DELETE is commented out (since the button itself is disabled).

**`tbrMenu_ButtonClick`** — same four actions wired through the toolbar.

**`lvRoomTypes_Click`** / **`lvRoomTypes_KeyUp`** — when the user picks a type, calls `PopulateValues`. The keyboard wrapper delegates to the click handler.

**`ListRoomType`** — populates the master list:
```sql
SELECT ID, TypeShortName, TypeLongName, Active FROM RoomType
```
For each row: adds to ListView, color-codes by Active. Calls `lvRoomTypes_Click` at the end so the first row's data auto-loads into the editor.

Note: no `WHERE Active = TRUE` filter — *both* active and inactive types are shown here. This is the editor; you need to see inactive ones to reactivate them. (Compare with `frmRoomMaintain.PopulateRoomType` which filters to active-only because that's a *consumer* of the list.)

**`PopulateValues(plngRoomTypeID)`** — reads the selected row, calls `ResetFields` first to clear, then populates each control. If no row is found (which shouldn't happen since the ID came from the ListView), shows an info MsgBox. Straightforward.

**`SelectRoomType(plngRoomTypeID)`** — public entry point used by `frmRoomMaintain` when launching this form. Same pattern as `SelectReport` and `SelectRoom` in the other Maintain forms: walks the ListView for a match, selects/scrolls/clicks if found, otherwise stamps the ID into the editor for a "create new" workflow. The Ctrl+T launch from `frmRoomMaintain` always passes 1 — meaning it lands on the first row regardless of which room was being edited.

**`SaveRoomType`** — the most interesting method here:

1. **Required-field validation**: short name is mandatory.
2. **Confirmation dialog**.
3. **Re-fetch the row** via `SELECT ALL ... WHERE ID = ?`.
4. **For UPDATE path**:
   - **Critical safety check**: if the user is trying to mark this type as Inactive, calls `IsRoomTypeUsed(mstrShortName)` first. If any room currently has this RoomType assigned, refuses the update with "Room Type is currently used!" — preventing a foot-gun where deactivating a type would orphan the rooms referencing it. (Comment notes this was added 06 May 2015, so it's a later patch.)
   - Builds the UPDATE with TypeShortName, TypeLongName, Active.
5. **For INSERT path** (new type — when ID lookup didn't find a row):
   - Builds the INSERT with `CheckInput`-sanitized values (vs the UPDATE path which uses Trim only — slight inconsistency; UPDATE technically doesn't escape user input, though the typed SQL helpers presumably do their own escaping inside `SQL_SET_Text`).
   - The new ID is auto-generated by the database (no explicit ID column in the INSERT).
6. After save, calls `ResetFields` and `ListRoomType` to refresh.

A trailing comment "Recommend to log out and log in is required" is mysterious — there's nothing in the code that requires a re-login, and changes here don't affect cached data structures. Possibly a leftover note from when the developer thought they'd need to refresh a global cache.

**`ResetFields`** — blanks all four editor controls. Uses `chkActive.Value = 0` (numeric) where elsewhere the codebase uses `vbUnchecked` (the named constant). They're equivalent (0 == vbUnchecked) but it's another small inconsistency.

**`IsRoomTypeUsed(strRoomType)`** — counts rooms with this RoomType:
```sql
SELECT COUNT(ID) AS RoomCount FROM Room WHERE RoomType = ?
```
Returns True if any rooms reference it. The check uses the type's *short name* as the join key, not the ID — because the Room table denormalizes RoomType as a text column (consistent with what we saw in `frmRoomMaintain` and `frmBooking` where room/booking rows store the type name, not a foreign-key ID). This works but means renaming a type would orphan all the rooms referencing the old name.

The function has a subtle resource issue: the `Exit Function` after success doesn't `CloseRS rst` or `CloseDB`. Only the error path closes the recordset. So in the happy path, the recordset stays open until garbage collection — minor leak, probably swallowed by VB6's reference counting since `rst` goes out of scope at function return. The error path does close it. Cleanup opportunity.

**`DeleteRoomType`** is commented out entirely with a header comment "Should not use this." The body is also clearly copy-pasted from a User-management form — it references `cboUserGroup`, `txtUserID`, `UserData`, `ListUsers`, none of which exist on this form. So even if uncommented, it wouldn't compile. The intent is documented: hard delete is unsafe, use the soft-delete (Active = False) path instead.

## Notable points and quirks

- **Hard delete is disabled by design.** The DELETE button is hidden, the Ctrl+D shortcut is commented out, and the `DeleteRoomType` function body is copy-paste-corrupt-and-commented. The system supports only soft delete via `chkActive`. This is the right choice given the denormalization — Room and Booking rows store the type name as text, so a hard delete would create dangling references.

- **Active toggle is gated by usage check.** This is the standout safety feature: you can't disable a type while any room currently has it assigned. Combined with the soft-delete-only design, this means once a type is in use, your only options are (a) reassign all rooms to a different type, then deactivate; or (b) leave it active. Sensible.

- **Inconsistent input sanitization.** UPDATE path passes raw `mstrShortName` and `mstrLongName` directly to `SQL_SET_Text`; INSERT path wraps both in `CheckInput`. Either both should sanitize or neither should — the typed SQL helpers presumably handle quoting consistently, but the asymmetry is visible. Probably fine in practice.

- **Re-login recommendation comment** ("Recommend to log out and log in is required") doesn't seem to correspond to any actual requirement. Either an outdated note from when there was caching, or anticipating something that never got implemented.

- **`Form_Unload` calls `frmRoomMaintain.Form_Load` directly.** This is technically valid in VB6 — form event handlers are just regular subs that you can invoke explicitly. But it's not idiomatic. A cleaner pattern would be `frmRoomMaintain.PopulateRoomType` (which already exists and is what `Form_Load` calls anyway). The direct `Form_Load` call also re-runs the clock setup, button enable/disable logic, ListRooms, and lvRooms_Click — all to refresh one combobox. Functional but heavy-handed.

- **Hidden ID column + visible ID label.** The ListView has the ID column at width 0 (so it can be referenced via `SelectedItem.Text` without showing a redundant column), while the editor surfaces the ID in `lblRoomTypeID` as a read-only field. Same pattern as the other Maintain forms.

- **Soft delete UX:** the gray/pink coloring in the ListView lets the user see at a glance which types are active vs disabled. Combined with the IsRoomTypeUsed check, the user has clear feedback about why an Inactive flag might fail to save.

- **Public `SelectRoomType` always called with `1`.** The `frmRoomMaintain.tbrMenu_ButtonClick` and `Form_KeyDown` both invoke `frmRoomTypeMaintain.SelectRoomType 1` regardless of what room the user was editing. The intent presumably was to land on a known-good first row rather than guessing. A more contextual approach would be to pass the *current* room's RoomType ID, so editing room 12 (a "DORM") would auto-select the DORM type for editing. Minor UX miss.

- **`CheckInput` vs `CheckString`.** Different validation/sanitization helpers used here vs `frmReportMaintain.SaveReport`. Probably the same family of functions with subtly different escaping rules. Worth verifying they handle apostrophes, semicolons, etc., consistently across the codebase.

- **No audit trail.** Unlike `frmRoomMaintain` (which maintains a `LogRoom` history table), this form just overwrites the row. Renaming a RoomType from "DELUXE" to "PREMIUM" leaves no history of what it used to be called — and since rooms reference types by name, those rooms won't auto-update. A user could rename "DELUXE" → "PREMIUM" here, then go back to find that all their existing DELUXE rooms still say "DELUXE" because the Room.RoomType field stores the name verbatim. The form doesn't warn about this.

- **Renaming a used type is allowed but breaks references.** Following from the above: there's no protection against renaming a type that's currently in use. The `IsRoomTypeUsed` check only fires when toggling Active off, not when changing the name. So a user could rename "DORM" to "DORMITORY", and the next time `frmRoomMaintain` loads the room type combo, the rooms that store "DORM" would have their `cboRoomType.ListIndex` fall back to -1 (unset) because the name no longer exists in the list. Real bug, though probably unlikely to be hit in practice since type renames are rare.

- **Tab order on the editor is logical**: ID label (TabIndex 0, but it's a read-only label so unfocusable), ShortName (1), LongName (2), Active (3), then the ListView (4) and toolbar buttons. So tabbing through the editor goes top-to-bottom through the three editable fields, then jumps to the ListView.

- **The MaxLength tooltips on both textboxes** ("Max length = 30" and "Max length = 255") are a small UX touch — telling the user the limit before they hit it. Same treatment isn't given to the textboxes in `frmRoomMaintain`.

- **`Picture2` editor has TabStop = 0** but its child controls have proper tab indices, which is the standard VB6 pattern for using a picture box as a layout container without making it a tab stop itself.
