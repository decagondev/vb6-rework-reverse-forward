# frmDashboard — Room Booking System Dashboard

This is the main dashboard form for a hotel/hostel room booking application. Here's the full breakdown:

## Layout (top to bottom)

**Form shell**
- Fixed single border, no min/max buttons, 15375 × 9645 twips, dark gray background (`&H505050`), centered on owner. `KeyPreview = True` so the form catches keystrokes before its child controls.

**Header band (the `shpCopyright` rectangle behind everything)**
- `imgLogo` on the left
- `lblBusinessName` right-aligned, large (24pt bold) — set at runtime to `COMPANY_PRODUCT_NAME`
- `lblCopyright` underneath: "Computerise System Solutions 2014-2021"

**Toolbar `tbrMenu`** (with an `ImageList1` of 9 icons)
Eight buttons, each with an F-key shortcut:
- CLOSE (Esc), REPORT (F2), CUSTOMER (F3), ROOM (F4), USER (F5), ACCESS (F6), BLINK (F7), SECURITY (F8)

Inside the toolbar sits a `fraLogin` frame on the right holding:
- `lblSystemDateTime` — live clock
- `lblUserID` — "User ID : <name>"
- `tmrClock` (1 sec interval, idle-logout watchdog)
- `tmrBlink` (500 ms, disabled by default)

**Summary strip `Frame1`**
A row of colored squares + labels showing live counts:
- 🟢 Open (green `#00E676`)
- 🟡 Booked (yellow `#FFEA00`)
- 🔴 Occupied (red `#FF1744`)
- 🟣 Housekeeping (purple `#D500F9`)
- 🟠 Maintenance (orange `#2979FF` — actually labeled as blue in code, `COL_BLUE`)

**Four level panels (PictureBoxes)** stacked vertically:
- `Picture4` — Level 4: rooms 045–055 (11 buttons)
- `Picture3` — Level 3: rooms 034–044 (11 buttons)
- `Picture2` — Level 2: rooms 012–033 (22 buttons in two rows)
- `Picture1` — Level 1: rooms 001–011 (11 buttons)

Each panel has a "Level N" label in its top-left corner. All buttons are members of the `cmdUnit` control array (Index 0 is a hidden "R000" button — likely a template/dummy). Total addressable indices: 0–55, with the code looping up to 61 in places (over-allocated for safety).

**Hidden popup menu `mnuPop`**
Surfaced via right-click on a room button:
- Booking
- Edit Room
- Change Status → Free / Occupied / Housekeeping / Maintenance

## Logic

**Startup (`Form_Load` → `Form_Activate`)**
1. Sets the title, clock, and user ID label.
2. `ReCaptionButton` — clears every room button's caption and paints it gray.
3. `SetButtonProperties` — the core renderer (see below).
4. `LoadBlinkSetting` — reads `UserData.DashboardBlink` for the current user; if true and any rooms need attention, starts the blink timer.
5. Calls `ShowSummary1`–`ShowSummary5` to populate the count labels.
6. Enables/disables each toolbar button against `UserAccessModule(...)` — module-level permission checks for REPORT, CUSTOMER, ROOM, USER, ACCESS.

**`SetButtonProperties` — the heart of the form**
Runs a SQL join `Room LEFT JOIN Booking ON Room.BookingID = Booking.ID`, then for each row:
- Writes the caption as `RoomShortName` + room type on a new line.
- Colors the button by `RoomStatus`:
  - Maintenance → blue, no blink
  - Housekeeping → purple, no blink
  - Booked → yellow, blinks if `AlertBooking` returns true (i.e., `Now > GuestCheckIN`, meaning the guest is late)
  - Occupied → red, blinks if `Now > GuestCheckOUT` (overdue checkout)
  - Anything else → green
- Hides the button if the room is inactive.

**`tmrBlink_Timer`** alternates `blnDim` and swaps the backcolor between the status color and gray for any room where `blnBlink(i) = True`. Maintenance rooms are explicitly excluded from blinking (per the Nov 2014 modification note).

**Room button left-click `cmdUnit_Click`**
1. Permission check (`MOD_BOOKING`).
2. Refuses if the room is under Maintenance.
3. If the room hasn't been set up in the `Room` table, offers to open `frmRoomMaintain` (only if user has that permission).
4. Otherwise opens `frmBooking` for that room and hides the dashboard.

**Room button right-click `cmdUnit_MouseUp`**
Shows the popup with a different enable-set depending on current status:
- Open → Booking + go to Housekeeping/Maintenance
- Booked → Booking + go to Occupied (i.e., check in)
- Occupied → Booking only (status change disabled)
- Housekeeping → go to Free or Maintenance
- Maintenance → go to Free only

**Status-change menu handlers** (`mnuFree`, `mnuOccupied`, `mnuHousekeeping`, `mnuMaintenance`)
Each one builds an `UPDATE Room SET RoomStatus = '...', LastModifiedDate, LastModifiedBy WHERE ID = <button index>` via the `SQL_*` helper layer, runs it, refreshes the relevant summary labels, and re-renders the buttons. `mnuOccupied` additionally updates the linked Booking row's `GuestCheckIN` to `Now` and asks for confirmation first.

**Keyboard shortcuts (`Form_KeyDown`)**
Mirrors the toolbar: Esc closes, F2–F8 fire the same actions as the toolbar buttons, gated by the same `.Enabled` flags.

**Idle auto-logout (`tmrClock_Timer`)**
Increments `intTick` every second; when it exceeds the global `gintUserIdle`, disables itself and shows `frmDialog` modally (a re-auth/lockout dialog). Updates the date/time label once per minute (`If Second(Now) = 0`).

**Form unload** falls back to `frmUserLogin.Show` — closing the dashboard returns to the login screen.

**Blink toggle (`BlinkButton`)**
Flips `tmrBlink.Enabled`, persists the preference per user via `UpdateBlinkSetting` writing to `UserData.DashboardBlink`, swaps the toolbar button's caption/icon between "Blink (F7)" and "Unblink (F7)", and re-renders.

## Notable structural points

- Heavy use of a custom SQL builder (`SQL_SELECT`, `SQL_FROM`, `SQL_WHERE_Text`, `SQL_SET_*`, etc.) that writes to a global `gstrSQL` string — classic VB6 pattern but means every query mutates shared state.
- Error handling everywhere follows the same template: `On Error GoTo CheckErr`, `CloseRS` + `CloseDB`, MsgBox, then `LogErrorDB`.
- Color constants are defined in BGR-hex form at the top of the module with a comment explaining the byte-reversal versus standard RGB.
- Two parallel arrays `blnBlink(61)` and `strBackColor(61)` track which buttons should blink and what color to flash — sized 61 for headroom even though only ~55 rooms exist.
- `cmdUnit(0)` is invisible and likely just keeps the control array valid from index 0.
- The version header notes three modifications in late 2014: idle auto-logout, skipping blink for Maintenance, and the `NeedBlink` helper.

One small bug worth flagging: `Picture4` (Level 4) is positioned at `Top = 2640` but `Picture3` (Level 3) is at `Top = 4080`, `Picture2` at `5520`, and `Picture1` at `8040` — so Level 4 sits *above* Level 3 visually, which matches building floors but means the Level 2 panel (height 2520) overlaps slightly with the summary strip area depending on screen scaling. Worth verifying at runtime.
