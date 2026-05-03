# frmBooking — Room Booking / Check-In-Out Form

This is the form launched when a user clicks a room button on the Dashboard. It handles the full lifecycle of a single booking: create → save → check-in → check-out → print receipt.

## Layout (top to bottom)

**Form shell**
- Same dark theme as the Dashboard (`#303030` background), fixed single border, no min/max, 15375 × 9645 twips, `KeyPreview = True`.

**Header band** — identical pattern to the Dashboard: `imgLogo` left, `lblBusinessName` (24pt) right-aligned, `lblCopyright` underneath, sitting on a `shpCopyright` rectangle.

**Toolbar `tbrMenu`** with 8 buttons (icons via `ImageList1`, 8 images):
- CLOSE (Esc), RESET (Ctrl+R), SAVE (Ctrl+S), VOID (Ctrl+D — *button is hidden/disabled by design*, code commented out), Check-IN (Ctrl+I), Check-OUT (Ctrl+O), TEMPORARY receipt (Ctrl+F11), OFFICIAL receipt (Ctrl+F12).
- `fraLogin` on the right of the toolbar carries the live clock and `lblUserID`.

**Status strip `fraStatus`** — full-width colored bar that doubles as a status indicator:
- `lblBookingNo` "Booking No :" + `lblBookingID` showing either "New" or a 6-digit padded ID (`Format(id, "100000")`).
- `lblStatus` right-aligned (Open / Booked / Occupied / Housekeeping / Maintenance / Void).
- The frame's `BackColor` changes with status — green/yellow/red/purple/blue/gray.

**Main body — two columns**

*Left column (rooms ~240 px from edge):*
- Section header `fraBooking` ("Booking Details") on top of `Picture1`, which holds:
  - `dtpBookingDate` + `dtpBookingTime`
  - `cboTotalGuest` (1–6, populated in `Form_Load`)
  - `cboStayDuration` (1–10 nights, populated in `Form_Load`)
  - `dtpCheckInDate` + `dtpCheckInTime`
  - `dtpCheckOutDate` + `dtpCheckOutTime`
- Section header `fraGuest` ("Guest Details") on top of `Picture2`:
  - `txtGuestName *`, `txtGuestPassport *`, `txtGuestOrigin`, `txtContactNo`
- Section header `fraEmergency` ("Emergency Contact") on top of `Picture3`:
  - `txtGuestEmergencyContactName`, `txtGuestEmergencyContactNo`

*Right column (~9480 px from edge):*
- Section header `fraRoom` ("Room Details") on top of `Picture4`, which holds:
  - `lblRoomID` (hidden) — internal ID
  - Read-only labels: `lblRoomNo`, `lblRoomType`, `lblLocation`, `lblRoomPrice`
  - `lblSubTotal` — calculated from days × rate
  - `txtDeposit` (defaults to "20.00")
  - `lblTotalDue` — sub-total + deposit
  - `txtPayment`, `txtRefund`
- Section header `fraOther` ("Remarks") on top of `Picture5`:
  - `txtRemarks` (multiline, 255 max)
  - Hidden `lblBreakfast` and `lblBreakfastPrice` — feature wired but invisible on the form (probably staged for later).

The two `*` labels (Name, Passport/IC No) are mandatory; the form refuses to save without them.

## Module-level state

```vb
lngBookingID    ' Current booking, 0 = new/unsaved
intRoomID       ' Which room this form was opened for
strStatus       ' Open / Booked / Occupied / Housekeeping / Maintenance / Void
blnActive       ' Declared but unused
intTick         ' Idle-logout counter (same pattern as Dashboard)
datCheckIn / datCheckOut / intDay  ' Date math helpers
```

The same `COL_*` color constants from the Dashboard are duplicated at the top.

## Logic

**Lifecycle**

`Dashboard.cmdUnit_Click` → `frmBooking.SelectRoom(roomId)` → `frmBooking.Show`. That path is the only documented entry.

**`Form_Load`** — sets clock, user ID, populates the two combos (1–6 guests, 1–10 nights), then calls `ResetFields`.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog firing `frmDialog` after `gintUserIdle` seconds, identical to Dashboard).

**`Form_Unload`** — refreshes the Dashboard's button colors and summary counts, then re-shows it. Closing this form effectively returns control to Dashboard with fresh data.

**`SelectRoom(pintRoomID)`** — public entry point called by Dashboard:
1. Stores the room ID and reads room details via `GetRoomDetails`.
2. If the room has no booking yet (`lngBookingID = 0`), calls `CreateTempBookingID` to insert a placeholder row in `Booking` with `Temp = TRUE`, then displays it as "Booking No: 100xxx". The temp flag is flipped to `False` on the first real save.
3. Otherwise calls `PopulateValues` to load the existing booking.

**`CreateTempBookingID` / `GetTempBookingID`** — clever pattern: rather than auto-numbering at save time, every room click creates a draft Booking row immediately, retrieves its identity, and uses that as the visible booking number. The row stays flagged `Temp = TRUE` until SaveBooking clears it.

**`PopulateValues(plngBookingID)`** — single big SELECT joining `Booking B, Room R WHERE B.ID = ? AND B.ID = R.BookingID`. Pulls every guest field, dates, room snapshot (room details are denormalized into Booking — RoomType, RoomLocation, RoomPrice, Breakfast etc. are all stored on the booking too, so historical receipts stay accurate even if the room's current price changes). Recalculates `lblTotalDue` and switches Void↔Unvoid button caption based on `Active`. If no booking found, falls back to `GetRoomDetails` for a fresh form.

**`ShowStatus(pstrStatus)`** — drives the whole UI based on status:
| Status | Bar color | Check-IN | Check-OUT | Refund | Notes |
|---|---|---|---|---|---|
| Open | green | off | off | off | |
| Booked | yellow | **on** | off | off | |
| Occupied | red | off | **on** | **on** | BookingID label turns white |
| Housekeeping | purple | off | off | off | SAVE disabled, fields disabled via `DisableControls` |
| Maintenance | blue | off | off | off | hides BookingNo+ID |
| Void | gray | off | off | off | |

**Date arithmetic on the combos and pickers**

- `cboStayDuration_Click` — when nights changes, computes `CheckOutDate = CheckInDate + N` if check-in is at/after noon, else N-1 (same-day morning check-ins still count as a "night"). Forces check-out time to 12:00 PM.
- `dtpCheckInDate_Change` — if check-in date is not today, snaps check-in time to noon; recomputes check-out date/time, then `SumTotal`.
- `dtpCheckOutDate_Change` — works backwards: pulls check-in date back by `intDay`.
- `dtpCheckOutTime_Change` — the day-offset logic gets re-applied with two thresholds (12:00 PM and 2:00 PM) — the 2:00 PM branch is unreachable as written because the `>= 12:00 PM` test catches it first. Looks like a bug or leftover from a previous policy.

**`SumTotal`** — `SubTotal = nights × roomPrice`, `TotalDue = SubTotal + Deposit`. Re-runs on most field changes.

**`SaveBooking`**
1. Validates Name + Passport (both mandatory) and that both combos have selections.
2. Confirmation dialog.
3. UPDATEs the existing Booking row (the temp record) with all guest + room snapshot + financial fields.
4. If status was "Open" it promotes to "Booked" and stamps `CreatedDate / CreatedBy` (since the Temp row's "create" wasn't really by the user).
5. Sets `Temp = FALSE` to commit.
6. Calls `UpdateRoomStatus` to push the new status onto the Room row (and, for "Booked", sets `Room.BookingID`).

**`Check_IN`**
- Refuses if booking unsaved or `IsPaid = False` (payment must equal SubTotal + Deposit).
- Asks confirmation showing the Check-IN date/time.
- UPDATEs Booking with the actual check-in datetime, sets Room status to "Occupied".

**`Check_OUT`** — the most policy-heavy method:
- Refuses if unsaved or unpaid.
- **Late checkout rule:** if check-out time ≥ 2:00 PM, refund is forced to 0.00 and the field is disabled with a "DEPOSIT NO REFUND" message.
- Otherwise asks "fully refund the deposit?" — Yes copies deposit into refund, No keeps current refund value.
- Final confirmation showing the actual refund amount.
- UPDATEs Booking with check-out datetime + refund, sets Room status to **"Housekeeping"** (not Open) — so the room must be cleaned before it's bookable again. This is consistent with the Nov 2014 modification note at the top.

**`IsPaid(plngBookingID)`** — returns `Payment = SubTotal + Deposit`. Used as a gate before Check-IN and Check-OUT.

**`UpdateRoomStatus(status, roomID, [bookingID])`** — central writer for `Room.RoomStatus`. Only sets `Room.BookingID` when transitioning to "Booked". Maintenance/Occupied transitions don't touch BookingID.

**`VoidBooking`** — toggles `Booking.Active`. The toolbar VOID button is currently disabled/hidden but the function and Void↔Unvoid caption swap in `PopulateValues` are wired up. Note: `WHERE ... AND Temp = FALSE` so temp bookings can't be voided.

**`PrintReceipt(strType)`**
- Builds a SQL string by string concatenation (not via the SQL helper layer used everywhere else — unusual) joining `Company` and `Booking`, then opens a new `frmPrint` form which in turn presumably calls a Crystal Report (`gstrReportFileName = "Official Receipt.rpt"` or `"Temporary Receipt.rpt"`).
- Temporary receipt totals are `Payment - Deposit` for sub-total, then deposit shown separately.
- Official receipt totals are `Payment - Refund`.
- If the join returns no rows, falls back to a Company-only query with placeholder zeros so the report still renders.
- F12 (Official receipt) does an extra "this room is not Checked Out — continue?" prompt unless the room is already in Housekeeping (i.e., already checked out).

**`ResetFields`** — clears all guest text boxes; resets booking date/time and check-in date/time to "now"; deposit defaults back to "20.00" (clearly hard-coded house policy); selects first item in both combos; runs SumTotal.

**`DisableControls`** — locks every editable field. Called by `ShowStatus` when status is Housekeeping.

**Currency formatting on lost-focus** — `txtDeposit_LostFocus`, `txtPayment_LostFocus`, `txtRefund_LostFocus` all run `ConvDbl` then `FormatCurrency` so the user can type "50" and have it become "50.00". The deposit one also re-runs the TotalDue calc.

## Notable points and quirks

- **Denormalized booking snapshot.** Booking stores its own copy of RoomType, RoomLocation, RoomPrice, Breakfast, BreakfastPrice. Sensible for receipts that need to reflect prices at booking time, but it means `PopulateValues` reads room data from `Booking`, while `GetRoomDetails` reads from `Room`.
- **Two parallel `intDay` variables** — one module-level and one local in `SumTotal`. The local shadows the module one inside that sub, which means the value computed via `cboStayDuration_Click` is what date-shift logic ends up using; works but easy to misread.
- **The 2:00 PM branch in `dtpCheckOutTime_Change` is dead code** — the prior `>= 12:00 PM` condition always matches first.
- **Inconsistent SQL building.** Most of the form uses the typed helper layer (`SQL_SET_Text`, `SQL_WHERE_Long`, etc.) but `PrintReceipt` builds raw SQL by string concatenation. Not parameterized, so any apostrophe in `gstrUserID` would break it (probably fine in practice since user IDs are controlled).
- **Cancellable "ConvText" pattern.** `ConvText`, `ConvDbl`, `ConvInt`, `ConvLng` are clearly app helpers that null-coalesce database values. They're called consistently before any nullable column is read.
- **Deposit hard-coded to MYR 20.00** in `ResetFields` — currency is Malaysian Ringgit, confirmed by the `MYR` labels.
- **The Void button is intentionally not implemented** — the comment says "VOID is not implemented since not a request". Both keyboard handler and toolbar handler have the entire Void block commented out, but `VoidBooking` itself is fully written and `PopulateValues` swaps its caption — so it's one boolean and a code uncomment away from being live.
- **Temp booking pattern is fragile.** If a user opens a room, gets a TempBookingID, then closes the form without saving, the temp Booking row stays in the database. Next time `CreateTempBookingID` runs it picks up the orphan via `GetTempBookingID` (TOP 1 ORDER BY ID DESC by the look of it) and reuses it — which works for single-user, but in a multi-user environment two staff opening different rooms simultaneously could race onto the same temp row. Worth flagging.
- **`Form_Unload` chain back to Dashboard.** The form acts as a modal-ish child even though it's not actually shown modally — closing it always re-shows the dashboard with refreshed data. If the dashboard isn't loaded for any reason this would error, but in practice it always is.
- The status colors duplicate the constants from the Dashboard rather than referencing a shared module, so a theme change would need to touch both files.
