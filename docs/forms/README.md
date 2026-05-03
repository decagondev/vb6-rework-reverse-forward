# Room Booking System — Forms Index

A Visual Basic 6 application for hotel/hostel room management, originally branded for "StarHotel" and dated 2014–2021. The system handles room booking, check-in/out, payment tracking, customer history, reporting, and access control for a building with up to 55 rooms across four floors.

## System overview

### What the application does

The system is a single-machine, single-user-at-a-time room management tool backed by a password-protected Microsoft Access database (`.mdb`). Operators (clerks and admins) use it to:

- See the current state of every room at a glance via a color-coded grid (the Dashboard)
- Make new bookings, check guests in, check guests out, and process payments
- Track guest history across visits via passport/IC number lookup
- Run scheduled reports (daily, weekly, monthly, yearly bookings)
- Print receipts (temporary during stay, official on checkout)
- Manage rooms, room types, users, permissions, and report definitions

The currency is Malaysian Ringgit (MYR), the default deposit is hardcoded at MYR 20.00, and the building layout (rooms numbered 001–055 across Levels 1–4) is baked into the Dashboard's button grid.

### Architectural shape

The system is a **strict hub-and-spoke**:

- `frmDashboard` is the central hub
- Every other operational form launches from the Dashboard and returns to it on close
- There is no cross-traffic between leaves — to go from booking to customer search, you return to the Dashboard first
- The Dashboard's color states (Open, Booked, Occupied, Housekeeping, Maintenance) drive the entire visual language of the system

There are two embedded sub-hubs:

- `frmReport` → `frmReportMaintain` (Ctrl+E expert path) — editing report metadata is gated by a developer password
- `frmRoomMaintain` → `frmRoomTypeMaintain` (Ctrl+T) — room type CRUD is nested inside room CRUD

A universal print form (`frmPrint`) wraps Crystal Reports and is fed by globals from any caller that needs to print.

### Key design patterns

- **Metadata-driven reports.** Reports are rows in a `Report` table, not code. Adding a new report means inserting a row and dropping a `.rpt` file in the Reports folder.
- **Soft delete everywhere.** Every editable entity has an `Active` flag. Hard delete buttons exist in the form designer but are universally disabled.
- **Color-coded ListViews.** Active = gray, inactive = pink, plus status-specific colors (green/yellow/red/purple/blue) consistent across forms.
- **Audit trail via `LogRoom` and `LogBooking`** for room and booking edits — full historical snapshots, not just last-modified columns. Other tables (Users, Reports, Room Types) just overwrite without history.
- **Denormalized booking snapshots.** Booking rows store room details (type, location, price) at the time of booking, so old receipts stay accurate when current rooms get repriced.
- **Idle-logout watchdog** wired into every form via `tmrClock`, but currently inert because of a debug override (`gintUserIdle = 0`) in `frmUserLogin`.
- **Forced password change** flow gates first-login and admin-reset users via the global `gblnUserChangePassword`.

### Security model

The system uses a hand-rolled `GoldFishEncode` hash with a 4-character salt, capped at 8–10 character passwords. Other notable security characteristics:

- Two user groups in practice: Group 1 (Admin, full access) and Group 4 (Clerk, configurable). Schema supports four groups; the middle two are unused.
- Permissions are cell-by-cell in the `ModuleAccess` table, edited via `frmModuleAccess`.
- Default credentials are `admin` / `admin`, advertised on the login screen.
- Clicking the copyright text on the login screen auto-fills the admin credentials.
- A hardcoded developer password (`expert`) gates report editing.
- The Void operation has full database support but is intentionally disabled at the UI layer in two places.

The security posture is appropriate for an internal hotel POS but would not be suitable for an internet-exposed system without significant hardening.

## Form catalog

The forms below are listed in the order a user encounters them in a normal session.

### Bootstrap & authentication

#### `frmSplash`
The application's first screen. Despite looking like a branding splash, it does almost all the heavy lifting: reads `Config.txt`, locates the database file, validates its schema version, runs migrations (1.1 → 1.2 → 1.3) using a backup-then-copy strategy, instantiates the Crystal Reports application, then hands control to `frmUserLogin`. Contains a commented-out beta kill-switch that would expire the system 99 or 364 days after Jan 1, 2021.

#### `frmDatabase`
First-run database picker. Lets the operator either use the bundled demo database (`DemoData.mdb`) or specify a custom path and filename. On confirm, writes the choice to `Config.txt` (a two-line text file: path then filename). Will create a new database file by calling `CreateData`/`CreateDB`/`CreateSampleData` if the chosen file doesn't exist.

#### `frmUserLogin`
The login screen. Validates User ID and password against the `UserData` table using `GoldFishEncode + Salt`. Locks accounts after 3 wrong attempts (for non-admins). On success, either jumps to `frmUserChangePassword` if the user has `ChangePassword = True` set, or proceeds to `frmDashboard` if they have `MOD_DASHBOARD` permission. Cancel = `End` (terminate app).

#### `frmUserChangePassword`
Maximized full-screen password change form. Used in two flows:

- **Forced** (set by login when `ChangePassword = True`): Cancel terminates the app; the user must change their password to proceed
- **Voluntary** (F8 from Dashboard): Cancel returns to the Dashboard

The `gblnUserChangePassword` global flag distinguishes the two. Validates that old password matches, new password is at least 4 chars, and new + confirm match. Generates a fresh 4-char salt on every change.

### Hub

#### `frmDashboard`
The system's central screen and the form most users live in. Displays a 55-button grid arranged across four floor panels (Level 1: rooms 001–011, Level 2: 012–033, Level 3: 034–044, Level 4: 045–055). Each button is colored by its room's current status; bookings approaching their check-in/out time blink (if blink is enabled per-user). Provides toolbar shortcuts:

- F2 → Reports (`frmReport`)
- F3 → Find Customer (`frmFindCustomer`)
- F4 → Maintain Rooms (`frmRoomMaintain`)
- F5 → Maintain Users (`frmUserMaintain`)
- F6 → Access Control (`frmModuleAccess`)
- F7 → Toggle button blinking
- F8 → Change Password (`frmUserChangePassword`)
- Esc → Logout (returns to login)

Right-clicking a room button shows a context menu for status changes (Free, Occupied, Housekeeping, Maintenance) and shortcuts to edit the room or open the booking form. Left-clicking a room button opens `frmBooking` for that room — or, if the room isn't yet configured, jumps to `frmRoomMaintain` to set it up.

### Transactional forms

#### `frmBooking`
The booking lifecycle form, opened by clicking a room on the Dashboard. Implements the full state machine: New (pre-save) → Booked → Occupied (Check-IN) → Housekeeping (Check-OUT). Manages guest details, dates, payments, deposits, and refunds. Enforces several business rules:

- Refund is forced to 0 if checkout is at or after 2:00 PM (late checkout = no deposit refund)
- Booking cannot be saved without a guest name and passport
- Check-IN and Check-OUT both require full payment
- Rooms transition to "Housekeeping" on checkout, not "Open" — they must be cleaned before becoming bookable again

Creates a temporary draft Booking row immediately when opened, marked `Temp = TRUE`; the flag is cleared on first save. The Void button is wired but intentionally disabled.

#### `frmFindCustomer`
The customer search form (F3 from Dashboard). Three-panel master/detail layout:

1. Search criteria on the left (name, passport, country, contact, plus AND/OR connectors and a date range)
2. Matching distinct customers in the top-right
3. Selected customer's bookings in the bottom-wide panel

Click Find to search, click a customer to load their bookings, click Print to generate a Customer Transaction History report.

#### `frmReport`
The report catalog (F2 from Dashboard). Lists all reports defined in the `Report` table, color-coded by active state, gated per-user by `MOD_REPORT_*` permissions. The user picks a report and a "report as at" date, then clicks Preview (or double-clicks). The form assembles the SQL query by combining the report's `ReportQuery` with date filters (Range / Weekly / Monthly / Yearly / Since Start / Single), variable substitution (`$UserID$`, `$BookingID$`), and a SubQuery tail, then hands off to `frmPrint`. The Edit button is gated by both `MOD_REPORT_EDIT` permission *and* a hardcoded developer password (`expert`).

#### `frmAdmin`
A small modal step-up authentication dialog. Designed to authorize Void operations from `frmBooking`, but currently a no-op stub — the block that actually performs the Void is commented out at both ends. Authenticates and then exits without taking action. Kept in the codebase ready to be enabled by uncommenting four blocks across two files.

### Print

#### `frmPrint`
Universal Crystal Reports preview/print form. Called by every form that prints — receipts (`frmBooking`), customer history (`frmFindCustomer`), and catalog reports (`frmReport`). Loads a `.rpt` template from `App.Path\Report\` and pushes a recordset into it as the data source. Toolbar buttons: Close (F4), Refresh (F5), Setup (F6), Export (F7), Print (F8). Uses a 100ms timer polling `GetKeyState` for keyboard shortcuts because the Crystal viewer ActiveX control swallows `KeyDown` events when focused.

### Admin tools

#### `frmRoomMaintain`
Room CRUD editor (F4 from Dashboard). Master/detail with the room list on the left and editor on the right. Edits Room No, Type, Location, Price, Breakfast/Price, plus active/maintenance/housekeeping flags. Enforces the invariant that Booked or Occupied rooms cannot be edited (all fields disabled). Maintains a full audit trail by inserting a snapshot into `LogRoom` before every UPDATE. Provides a Ctrl+T shortcut to `frmRoomTypeMaintain`.

#### `frmRoomTypeMaintain`
Room Type CRUD editor, reachable only from `frmRoomMaintain` via Ctrl+T. Edits the lookup table that drives the `cboRoomType` dropdown. Won't let you mark a type Inactive while any room currently uses it (`IsRoomTypeUsed` check, added 06 May 2015 as a later patch). On unload, calls `frmRoomMaintain.Form_Load` to refresh the parent's dropdown with any new types.

#### `frmUserMaintain`
User account CRUD editor (F5 from Dashboard). Edits user group, name, idle timeout, active flag, change-password flag, and (optionally) password. The password-update path is opt-in via a checkbox to prevent accidental resets. Includes a "Reset login attempts (N)" checkbox showing the current count, useful for unfreezing locked accounts. No audit trail (unlike room edits).

#### `frmModuleAccess`
The access-control matrix (F6 from Dashboard). A grid of modules (rows) by user groups (columns: Administrator and Clerk). Each cell is a checkbox; checking or unchecking immediately writes to the database — there is no Save/Cancel. Admin permissions are hardcoded to "checked and disabled" so admins can never be locked out. Only the Clerk column is interactive. Two of the four groups in the schema (Group 2 and Group 3) are unused.

#### `frmReportMaintain`
Report metadata CRUD editor, reachable from `frmReport` via Ctrl+E (after the developer password challenge). Lets you edit a report's name, title, "as on" prefix, active flag, plus the Expert pane with the actual SQL query, sub-query tail, null-result fallback query, date field/type configuration, and `.rpt` filename. Two-tier permission: `MOD_REPORT_EDIT` to enter the form; `MOD_REPORT_EDIT_EXPERT` to access the SQL pane (hidden by default until toggled). Save only updates existing rows — there's no Insert path here.

## Common shared elements

Every primary form shares the same chrome:

- Dark gray (`#303030`) background, fixed-single border, 15375×9645 twips
- A header band with logo, business name (24pt), and copyright line
- A toolbar with at minimum a Close button, often Save/Reset/Clear too, plus form-specific actions
- A login frame on the right of the toolbar showing the live clock and current user ID
- An idle-watchdog timer (`tmrClock`) that would fire `frmDialog` after `gintUserIdle` seconds — currently inert due to the login bug
- Standard error handling pattern: `On Error GoTo CheckErr`, MsgBox, `LogErrorDB` to log into the database

The dialog forms (`frmAdmin`, `frmDatabase`, `frmDialog`) use a smaller fixed-dialog layout with no taskbar entry. The bootstrap forms (`frmSplash`, `frmUserChangePassword`) use unusual layouts (frameless splash, maximized full-screen) to enforce specific UX states.

## File reference

| Form | Type | Trigger | Returns to |
|---|---|---|---|
| `frmSplash` | Bootstrap | App start | `frmUserLogin` or `frmDatabase` |
| `frmDatabase` | Dialog | Splash, or login Option button | Splash or login |
| `frmUserLogin` | Auth | After splash, or logout | Dashboard or change-password |
| `frmUserChangePassword` | Auth | F8 from Dashboard, or forced | Dashboard, or `End` |
| `frmDashboard` | Hub | After login | login (on close) |
| `frmBooking` | Transactional | Click room on Dashboard | Dashboard |
| `frmFindCustomer` | Transactional | F3 from Dashboard | Dashboard |
| `frmReport` | Transactional | F2 from Dashboard | Dashboard |
| `frmAdmin` | Modal | Voiding from Booking (disabled) | Booking |
| `frmPrint` | Modal | Print from Booking, FindCustomer, Report | Caller |
| `frmRoomMaintain` | Admin CRUD | F4 from Dashboard | Dashboard |
| `frmRoomTypeMaintain` | Admin CRUD | Ctrl+T from RoomMaintain | RoomMaintain |
| `frmUserMaintain` | Admin CRUD | F5 from Dashboard | Dashboard |
| `frmModuleAccess` | Admin CRUD | F6 from Dashboard | Dashboard |
| `frmReportMaintain` | Admin CRUD | Ctrl+E from Report (with developer password) | Report |
