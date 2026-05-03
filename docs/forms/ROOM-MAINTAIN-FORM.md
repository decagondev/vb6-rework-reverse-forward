# frmRoomMaintain — Room Configuration & Status Editor

This is the F4-from-Dashboard form. It's a master/detail editor for individual rooms — the table that defines what rooms exist, their location/type/price, and current status. Also reachable from `frmDashboard.cmdUnit_Click` when a user clicks an unconfigured room button (which jumps here to set it up first).

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template: `imgLogo` left, `lblBusinessName` (24pt, right-aligned), `lblCopyright` underneath, `shpCopyright` rectangle.

**Toolbar `tbrMenu`** — 5 buttons:
- CLOSE (Esc)
- CLEAR (Ctrl+C) — blanks all fields without deleting
- RESET (Ctrl+R) — re-reads the selected room from DB, discarding edits
- SAVE (Ctrl+S)
- EDIT / "Type" (Ctrl+T) — jumps to `frmRoomTypeMaintain` for editing the available Room Types

`fraLogin` on the right with the live clock and `lblUserID`.

**Main body — three-panel layout**

*Left column — `Picture1` with `fraRoom` "Rooms" header:*
- `lvRooms` ListView with three columns: hidden ID, "Room No", "Location"
- Color-coded by status (see `ListRooms` below)

*Top-right — `Picture2` with `fraDetails` "Room Details" header:*
- `lblRoomID` (hidden) — internal ID
- `lblBookingID` (hidden) — current booking, if any
- `txtRoomShortName *` — e.g. "001", "012" (MaxLength 12)
- `cboRoomType *` — populated from RoomType table
- `cboLocation *` — Level 1/2/3/4 (hardcoded list from `.frx` resource)
- `txtRoomLongName` — full description
- `txtRoomPrice` (MYR)
- `txtBreakfastPrice` (MYR) + `chkBreakfast` (Yes/No, default Yes)
- `chkMaintenance` (Under Maintenance — orange label)
- `chkHousekeeping` (Under Housekeeping — purple label)
- `chkActive` (pink label)

*Bottom-right — `Picture3` with `fraRecord` "Record Details" header (audit info, all read-only labels):*
- Created Date / By
- Last Modified Date / By
- Last Occupied Date — present in the schema but `Visible = False` on the form

The three asterisks mark required fields. Field colors are deliberate: Active is pink (warns when off), Maintenance is orange, Housekeeping is purple — matching the Dashboard's status color scheme.

## Module-level state

```vb
intTick           ' Idle-logout counter
COL_YELLOW/GREEN/BLUE/RED/PURPLE/PINK/BLACK  ' Status colors
```

Earlier color values are commented at the top — same as the Dashboard's "we iterated on the palette" history.

## Logic

**`Form_Load`**
1. Sets the live clock and user ID label.
2. Initializes `lblRoomID` to "0" (no room selected yet).
3. Disables SAVE until a room is selected.
4. Calls `PopulateRoomType` to fill the Room Type dropdown from the `RoomType` table.
5. Calls `ListRooms` to populate the master list.
6. Calls `lvRooms_Click` so the first room is auto-loaded.

The header comment notes the design rule: *"If Status = 'Booked' or 'Occupied' Then cannot change Status = 'Maintenance' or 'Free'"* — this is the core invariant the form enforces.

**`Form_Unload`** — restores the Dashboard's button colors and refreshes all five summary counts (Open/Booked/Occupied/Housekeeping/Maintenance) before re-showing it. Same propagation pattern as `frmBooking`.

**`Form_KeyDown`** — Esc closes, Ctrl+C clears, Ctrl+R resets, Ctrl+S saves, Ctrl+T launches Room Type editor. All gated on the corresponding toolbar button being enabled.

**`tbrMenu_ButtonClick`** — same five actions wired through the toolbar, including the Room Type editor jump (which calls `frmRoomTypeMaintain.SelectRoomType 1` then shows it and hides this form).

**`PopulateRoomType`** — straightforward `SELECT TypeShortName FROM RoomType WHERE Active = TRUE` populating `cboRoomType`. Inactive room types are filtered out — so renaming a room type as inactive removes it from new-room choices but doesn't break old rooms that still reference the old name.

**`lvRooms_Click`** / **`lvRooms_KeyUp`** — when the user picks a room, calls `PopulateValues`. The keyboard wrapper delegates to the click handler, so arrow-key navigation populates the editor too.

**`PopulateValues(intRoomID)`** — reads the room and pushes data into controls. The interesting parts:

1. **For `cboRoomType` and `cboLocation`**, doesn't just `cboRoomType.Text = value` (which would fail silently if the value isn't in the list). Instead it loops the list looking for an exact match and sets `ListIndex`. If not found, ListIndex stays at -1 — so a room storing a no-longer-active room type displays as blank, signaling there's a mismatch to resolve.

2. **The status invariant is enforced here.** If `RoomStatus = "Booked"` or `"Occupied"`, *every* editable field is disabled, plus SAVE. The room is locked from edits while a guest is in or scheduled for it. This protects against changing the price or type mid-booking.

3. **Inactive rooms** also disable Maintenance and Housekeeping checkboxes — you can deactivate a room but you can't put a deactivated room into Maintenance or Housekeeping. Subtle but consistent.

4. **The status is reflected back** as checkbox state: `chkMaintenance` = checked iff status is "Maintenance"; `chkHousekeeping` = checked iff status is "Housekeeping". This means the two checkboxes plus the implicit "neither = Open" provide three of the five statuses (Maintenance, Housekeeping, Open). The other two (Booked, Occupied) can't be reached from this form — they're set by the booking flow.

5. **For new/missing rooms** (no row found), it falls through to a "blank slate" state with `<New>` placeholders in the audit fields. SAVE is enabled if `intRoomID > 0`. So clicking a button on the Dashboard for room 32 (which doesn't exist yet in the Room table) opens this form ready to create that specific room with ID 32.

**`ListRooms`** — populates the left ListView and color-codes each row:

- **Inactive** → pink, bold
- **Maintenance** → blue (bold)
- **Housekeeping** → purple (bold)
- **Open** → green (bold)
- **Booked** → yellow (not bold)
- **Occupied** → red (not bold)
- Else → black

The bold/non-bold split is interesting: actionable states (you can edit them) are bold; in-use states (Booked/Occupied) are non-bold. Visual cue for "this room is busy, don't fiddle."

**`SelectRoom(pintRoomID)`** — public entry point used by `frmDashboard.cmdUnit_Click`:
1. Walks `lvRooms.ListItems` looking for the matching room ID.
2. If found, selects, scrolls into view, calls `lvRooms_Click`.
3. If not found, sets `lblRoomID.Caption` to the new ID and calls `ResetFields` — the "create-new" path.

**`SaveRecord`** — the most code-dense method. It does an audit-log + upsert dance:

1. **ID guard.** If `lblRoomID < 1`, exits silently.
2. **Confirmation dialog** ("Do you want to Save this Room?").
3. **Required-field validation:** Room No, Room Type, Room Location.
4. **Re-fetches the existing row** to determine UPDATE vs INSERT.
5. **For UPDATE path:**
   - **First**, copies the existing row to `LogRoom` via `INSERT INTO LogRoom (...) SELECT ... FROM Room WHERE ID = ?`. So every change creates a historical snapshot in `LogRoom` before being overwritten. This is a proper audit trail — anyone can ask "what did this room look like before this edit?"
   - **Then**, builds the `UPDATE Room SET ...` with all the editor fields plus `LastModifiedDate` and `LastModifiedBy`.
   - **Status logic:** if Maintenance is checked → status = "Maintenance"; else if Housekeeping checked → status = "Housekeeping"; else → "Open". And — interesting subtlety — when Housekeeping is *un*checked the BookingID is also cleared to 0, breaking the link to whatever booking just checked out. (`elseif chkHousekeeping.Value = vbUnchecked` actually handles that path.)
6. **For INSERT path** (new room):
   - Builds an `INSERT INTO Room (...) VALUES (...)` with all the same fields, plus the `intRoomID` as the primary key (Room IDs aren't auto-generated — they're assigned by which Dashboard button was clicked, mapping 1:1 to room number 001–055).
   - Status: "Maintenance" if checked, "Housekeeping" if checked, else "Open".
   - No LogRoom snapshot for new rooms (nothing to snapshot).
7. **Persists** with `QuerySQL`, then refreshes both the ListView (via `ListRooms`) and the editor (via `SelectRoom` + `lvRooms_Click`) so the saved row is re-selected and the audit timestamps update.

There's a commented-out block showing an earlier "free room (clear BookingID)" workflow — it's been folded into the main UPDATE path now.

**`ResetFields`** — clears the editor for a new entry, with one clever detail:

```vb
Select Case intRoomID
    Case 0 To 22:  cboLocation.Text = "Level 1"
    Case 23 To 39: cboLocation.Text = "Level 2"
    Case 40 To 50: cboLocation.Text = "Level 3"
    Case 51 To 61: cboLocation.Text = "Level 4"
End Select
```

The location is auto-suggested based on which Room ID is being created — matching the Dashboard's physical layout (Level 1 has rooms 1–11, Level 2 has 12–33, Level 3 has 34–44, Level 4 has 45–55). Note the ranges here (0–22, 23–39, etc.) don't quite match the Dashboard's panel ranges (1–11, 12–33, 34–44, 45–55). The ranges in this form are wider and overlap differently — possibly leftover from an earlier numbering scheme, or accommodating future expansion. Either way, this is one place where the data and UI bind tightly, and a Dashboard layout change would need to be mirrored here.

**`chkMaintenance_Click` / `chkHousekeeping_Click`** — mutual-exclusion enforcement: checking one auto-unchecks the other. So a room can be in Maintenance OR Housekeeping but not both, and the UI prevents the impossible state.

## Notable points and quirks

- **Audit trail via `LogRoom`.** Every edit to an existing room snapshots the previous state into a separate table before applying the update. This is a real history table, not just a `LastModifiedBy` column — you can reconstruct any room's full edit history. Significantly more thorough than the audit story in `frmBooking` or `frmReportMaintain`. Worth knowing this exists if you ever need to investigate a "who changed this and what was it before?" question.

- **Status invariants enforced in two places.** The DB-level rule "Booked/Occupied rooms can't be edited" is enforced visually by disabling controls in `PopulateValues`, plus implicitly by the `Maintenance/Housekeeping` checkboxes only ever moving the room from Open ↔ Maintenance/Housekeeping. You can't get from Maintenance to Booked here, which is correct — that requires a booking workflow.

- **Room IDs are not auto-generated.** They're assigned by external context (Dashboard button index → room number). This is unusual for a database design but logical for this app: the Dashboard layout is fixed, every button maps to a fixed position, and the Room.ID is just a permanent slot identifier. Means deleting a room and re-inserting it with the same ID is fine.

- **Two combos with hardcoded vs DB-sourced lists.** `cboRoomType` is loaded from the database (configurable via `frmRoomTypeMaintain`); `cboLocation` is hardcoded in the `.frx` resource (Level 1/2/3/4). If a hotel ever added a Level 5, this would need a code change *and* the Dashboard layout would need new buttons. Not flexible but matches the building's physical structure.

- **`Last Occupied Date` field is present but invisible.** Schema supports it, layout reserves space for it, but `Visible = False`. Probably staged for a future "show me rooms that haven't been used recently" report and never enabled.

- **Required-field markers (`*`) on labels but no live validation.** The form only checks at save time, not on focus-leave. Missing values give a MsgBox and refocus the offending field — straightforward but leaves the user free to navigate around invalid state until they hit Save.

- **`Cls` and re-population pattern after save.** Every save → ListRooms → SelectRoom → lvRooms_Click chain. This wipes and rebuilds the entire ListView for one row's change. For 55 rooms this is fine (probably <50ms). For a much larger dataset you'd want to update only the changed row.

- **`chkBreakfast` defaults to Yes** but `Breakfast` and `BreakfastPrice` aren't shown anywhere on the booking flow that I've seen — `frmBooking` has hidden labels for these. So this form *can* set breakfast pricing per room, but the data isn't surfaced to the operator at booking time. Either a feature in flight or one that was deprecated visually but kept structurally.

- **`Active = False` rooms still appear in the ListView** but pink + bold. This is a "soft delete" pattern — inactive rooms aren't removed, just flagged. The Dashboard's `SetButtonProperties` hides them (`cmdUnit(intID).Visible = rst("Active")`), so they vanish from the booking grid but stay editable here. Sensible.

- **The `cboLocation` auto-fill in `ResetFields` uses Room ID ranges** rather than reading the Dashboard's actual physical layout. If someone reorganizes the building (renumbers rooms, moves a level around), this form's auto-suggestion will silently get it wrong without erroring. Worth knowing.

- **The mutual-exclusion checkbox approach is fragile.** Programmatically setting `chkHousekeeping.Value = vbUnchecked` from inside `chkMaintenance_Click` will fire `chkHousekeeping_Click` recursively, which sees Maintenance is now checked and would *re-uncheck* itself. In practice it works because the recursive call's check-state is already vbUnchecked (because we just set it), but it's a subtle re-entrancy that could break if the conditions change. Cleaner would be a single `cboStatus` combo.

- **`SQLText` calls in the LogRoom INSERT-SELECT pattern** — the form uses the typed SQL helper layer to build a `INSERT INTO X (cols) SELECT cols FROM Y WHERE id = ?` statement. Looking carefully, the column list for INSERT and the SELECT list have to match exactly; they're listed twice in code. Easy to break if a column is added to one but not the other. Comments confirm `RoomPreviousPrice` and `Maintenance` were once columns and got commented out — this is the form most affected by schema-creep cleanup.

- **`tbrMenu.Buttons("EDIT")` shortcut is Ctrl+T but the toolbar caption is "Type"** — small label/key mismatch (Ctrl+T = "Type editor", but the button face says "Edit (Ctrl+T)" earlier in this codebase pattern). Looking again, the caption is "Type (Ctrl+T)" so it's actually consistent. Move on.

- **Unlike `frmBooking.SaveBooking`, this form's SAVE doesn't update CreatedDate/CreatedBy on the UPDATE path** — only on INSERT. Correct: those should be immutable after creation. The audit story is via LogRoom snapshots and LastModified*.

- **Color-coded ListView with bold/non-bold to indicate "actionable vs in-progress" status** is a thoughtful UI touch. A user can scan the list and instantly see which rooms they can edit (bold green/blue/purple/pink) vs which are tied up (yellow/red, non-bold).

- **`cboRoomType.Clear` then re-populate** every time `Form_Load` runs but there's no `cboLocation.Clear` — because cboLocation is fixed-list from the form designer. Slight asymmetry but fine.

- **The `chkActive.Value = vbChecked` default** combined with `chkMaintenance.Enabled = False` when Active is false means the form has a minor edge case: if a user unchecks Active, the Maintenance/Housekeeping checkboxes lose their enabled state for the *current view session*, but the form doesn't re-evaluate this on click — only on `PopulateValues`. So unchecking Active and then trying to also set Maintenance in one editing session works (the disable only kicks in next time the room is loaded). Probably fine, but worth knowing.
