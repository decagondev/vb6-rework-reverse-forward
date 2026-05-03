# frmSplash — Application Bootstrap Screen

This is the application's first screen, shown immediately at startup. Despite looking like a simple branding splash, it does almost all the heavy lifting of the system bootstrap: loading config, locating the database, validating it exists, running schema migrations between versions, and finally launching the login screen.

## Layout

**Form shell**
- `BorderStyle = 0` (None) — no chrome, frameless window
- `ControlBox = False`, no min/max, no taskbar entry
- `StartUpPosition = 2` (Center Screen) — centers itself
- `KeyPreview = True` (though no keys are actually handled — leftover)
- `AutoRedraw = True` for smoother repaints during status updates
- 7800 × 4245 twips, dark gray (`#505050`)

**Single inner `Picture1` panel** (slightly darker `#303030`) acts as a framed box. Inside it:

- `lblCompanyProduct` (top-left, 18pt bold) — set at runtime from `COMPANY_PRODUCT_NAME`
- `lblLicenseTo` ("Unlicensed", but `Visible = False`) — top, full-width, hidden by default
- `lblProductName` (24pt bold, peach color `#FF8080`) — main product banner
- `lblPlatform` (15.75pt, "Windows XP or above") — right-side platform note
- `lblVersion` (12pt bold, peach) — set at runtime from `App.Major.Minor.Revision`
- `lblStatus` — the live "what's happening" line; this is the one that changes during boot
- `lblCopyright`, `lblCompany` (8.25pt, bottom-right) — set at runtime from `App.LegalCopyright` and `App.CompanyName`
- `lblWarning` (bottom) — "Note: This software is full version. Please check license for more information."
- `Timer1` (1000ms interval) — the bootstrap engine (see below)

## Module-level state

None. Everything happens in the timer tick.

## Logic

**`Form_Load`** — populates the static labels from `App.*` properties (Major/Minor/Revision, ProductName, LegalCopyright, CompanyName) and the global `COMPANY_PRODUCT_NAME` constant. That's it for load — the actual bootstrap runs on the first timer tick.

**`Timer1_Timer`** — the bootstrap state machine. It runs *once* (because the last line of the success path disables the timer), but the design uses a 1-second timer rather than `Form_Activate` so the user gets a full second of seeing the splash before any work begins. It updates `lblStatus` between phases and calls `Me.Refresh` to force the redraw, giving the operator real-time progress feedback. The phases are:

1. **Beta-version expiry check.** Computes `DateDiff("d", "01 Jan 2021", Now)`. There's a commented-out enforcement block that would compare to 99 or 364 days and `End` the application with "This system is expired!" / "Please contact us at sales@computerise.my". Both checks are disabled — the app is currently shipping with no expiry. Worth noting this kill-switch infrastructure exists.

2. **Database password generation.** `gstrPassword = GenWord` — a runtime-generated password used to open the Jet-encrypted Access database. This means the password isn't hardcoded as a string constant; `GenWord` (defined elsewhere) presumably derives it deterministically from something the app knows (executable signature, hardcoded seed, etc.) so that every install produces the same password without it being obvious in a hex dump. Reasonable obfuscation for the period.

3. **Default config.** If `gstrDatabasePath` is empty, sets `gstrCompanyID = "StarHotel"` (the default company), `gstrDatabaseExt = ".mdb"`, and constructs `App.Path\Data\StarHotel.mdb` as the default. So the system has a "company brand" name baked in — this is a system originally branded for StarHotel, made generic but with the brand still visible as the default config string.

4. **Config file handling.**
   - If `App.Path\Config.txt` doesn't exist: dismisses the timer, unloads the splash, and shows `frmDatabase` modally with the default path/file pre-filled. This is the first-launch flow — same path that `frmDatabase` writes to on its OK click.
   - If it does exist: reads two lines via `ReadTextFile "Config", 0, strPath` and `ReadTextFile "Config", 1, strFile`. Reconstructs `gstrDatabasePath` from path + filename, normalizing the trailing backslash if missing.

5. **Database file existence.** If the file doesn't exist on disk, opens `frmDatabase` modally again to let the user pick or create. Same pattern as step 4's first branch.

6. **Connect.** Calls `ConnectDB`. If it fails, sets the status to "Error during loading application." and unloads. No retry, no detailed message — silently dies. The internal MsgBox/log calls are commented out, so this is genuinely silent.

7. **Database version check.** Calls `DB_Version` (presumably `SELECT DatabaseVersion FROM Company`) and branches on the result:
   - **`= 1.2`**: backs up the current database to `<file>.bak` (deleting any existing backup first), then runs `Update_Database 1.3` to apply 1.2→1.3 migration.
   - **`= 1.1`**: runs `Migrate_Database` (the heavy migration to 1.2 — see below), backs up the old database, copies the migrated `.tmp` file over the old one, then runs `Update_Database 1.2`.
   - **`< 1.1`**: runs `Update_Database 1.1` to apply 1.06→1.0.10/1.1 migration. (The version label here is a bit confused — see notes.)
   - **Otherwise** (already at 1.3 or higher): "Database is already updated."

8. **Crystal Reports application.** `Set CrApplication = New CRAXDRT.Application`. This is the global Crystal Runtime instance referenced by `frmPrint` — it's created once at startup and reused. Explains why `frmPrint` could reference `CrApplication` without declaring it.

9. **Login.** Hides the status, disables the timer, shows `frmUserLogin`, unloads itself.

**Error path** sets the status to "Error loading application", logs to a text file (since the database may not be accessible), unloads.

## `Migrate_Database` (Version 1.1 → 1.2 migration)

This is the substantial migration routine. The strategy is interesting:

- A separate temporary database (`<companyID>.tmp.mdb`) has already been pre-created with the 1.2 schema (presumably `CreateData`/`CreateDB` from `frmDatabase` is what produces it).
- This sub opens *both* databases simultaneously — `OpenDB` for the source (1.1), `OpenData` for the destination (1.2).
- It then walks each source table and writes equivalent rows to the destination via `QueryDataSQL` (a separate helper that targets the data-side connection, not the main `gstrSQL`/`QuerySQL` pair).
- Logs each step to `MigrateDB.log` via `Log2File`.

The migration order, table by table:

- **UserData** — UPDATEs existing rows in the new DB (which presumably has User rows pre-seeded), copying over UserGroup, UserName, password hash, salt, login attempts, etc. Note: UserID isn't copied because it's the WHERE key.

- **RoomType** — UPDATEs by ID, copying short/long name and Active.

- **Room** — the trickiest of the bulk migrations. The 1.1 schema apparently had a `Maintenance` boolean and an `xRoomStatus` column (probably the original `Status`). The 1.2 schema consolidates these into a single `RoomStatus` column with the values "Open", "Booked", "Occupied", "Housekeeping", "Maintenance". The migration logic translates:
  - `Maintenance = True` → `RoomStatus = "Maintenance"` and clears BookingID
  - `Maintenance = False, BookingID = 0` → `RoomStatus = "Open"`
  - `Maintenance = False, BookingID > 0`: maps from `xRoomStatus` ("Booked"/"Occupied") or defaults to "Open" and clears BookingID
  - Plus careful null-handling on every nullable column (RoomLongName, BreakfastPrice, audit fields), using `SQLText "X = NULL"` to write actual NULLs rather than empty strings. This is the migration paying attention to data-quality details.

- **ModuleAccess** — completely renumbers and renames all 17 modules. The 1.1 IDs are remapped to new 1.2 IDs (e.g., 1=Booking → 2, 2=Dashboard → 1, 3=MaintainRoom → 9, etc.). Also stamps in proper `ModuleDesc1` strings ("Booking", "Dashboard", "Maintain Room", "Access Control", "Edit Report (Expert)", and so on). This explains the `+11` offset used in `frmReport.LoadReports` (`UserAccessModule(rst!ReportID + 11)`) — after this migration, report IDs 1–7 are mapped to ModuleAccess IDs 12–18. The renumbering was also where the slot numbering got established. The Case 16 block is commented out (the "Disabled Module" / Void Booking row, kept inactive).

- **LogRoom** — INSERTs all rows from the source LogRoom table into the destination, again with careful null handling. Note the field name `xRoomStatus` in the source maps to `RoomStatus` in the destination — the rename was done at migration time.

- **LogBooking** — same INSERT-all pattern. Notice the field renames: source's `GuestEmergencyContact` becomes destination's `GuestEmergencyContactNo`; source's `TotalDue` becomes destination's `SubTotal`; source's `PaidDeposit`/`PaidBalance` become `Deposit`/`Payment`. So this isn't just a schema bump — fields were renamed for clarity and a new `Refund` column (defaulted to 0) was added.

- **Booking** — same pattern as LogBooking. `Temp = False` is hardcoded (no temp bookings carry over).

- Finally adds an `Idle` column to UserData via `ALTER TABLE` — this is the user-configurable idle-logout interval (`gintUserIdle`).

After the migration completes, the caller in `Timer1_Timer` deletes the original `.mdb`, copies the `.tmp` over it (renaming `.tmp` to `.mdb`), then runs `Update_Database 1.2` to add anything from version 1.2 itself that wasn't part of the data copy.

## `Update_Database(dblVersion)` (in-place schema upgrades)

For smaller version bumps that don't need a full data copy, this sub applies in-place ALTERs and INSERTs. Three branches:

**`dblVersion = 1.1`** (1.06 → 1.0.10/1.1):
- Adds `DashboardBlink Bit` column to UserData.
- Renames `TotalDue` to `SubTotal` in both Booking and LogBooking.
- Updates `Company.DatabaseVersion` to 1.1.

**`dblVersion = 1.2`** (1.1 → 1.2):
- Inserts Report ID 5 ("Shift Report for User") if not exists, with a query that filters by `CreatedBy = '$UserID$'`.
- Inserts Report ID 6 ("Shift Report (All Users)") if not exists.
- Updates DatabaseVersion to 1.2.

**`dblVersion = 1.3`** (1.2 → 1.3):
- Inserts Report ID 7 ("Official Receipt (Reprint)") if not exists. Notably this is the report that uses `$BookingID$` substitution — the prompt-for-booking-ID flow you saw in `frmReport.PrintReport`.
- Activates ModuleAccess ID 18 with name "Official Receipt (Reprint)", granting access to Group1 (Admin) only — clerks can't reprint official receipts.
- Updates DatabaseVersion to 1.3.

The "if not exists" pattern (`If rst.EOF Then INSERT`) makes the migration idempotent — running it twice is safe.

## Notable points and quirks

- **The whole bootstrap runs in `Timer1_Timer`, not `Form_Load`.** This is a deliberate choice for UX: `Form_Load` finishes synchronously before the form paints, so heavy work there would mean a frozen splash. By kicking work off in a 1-second timer, the splash gets a chance to draw fully (with all the version/copyright info) before the actual loading begins. The user perceives the app as starting cleanly.

- **The timer is one-shot.** It disables itself after running once, but uses Timer rather than DoEvents+code in Form_Load. Slight waste of mechanism but works.

- **Beta kill-switch is disabled but ready.** Two `If intDay < 0 Or intDay > N Then End` blocks are commented out — one with a 99-day limit, one with 364. The infrastructure is sitting there fully written; uncomment one and the app becomes time-limited from `01 Jan 2021`. As of today (2026-05-03), `intDay` would be roughly 1949 — well past either limit, so enabling the kill-switch right now would brick all installations. Worth noting if anyone's tempted to flip it back on.

- **Three database lifecycles handled:**
  1. **First launch** (no Config.txt) → `frmDatabase` to set up.
  2. **Pre-1.1 install** → in-place `Update_Database 1.1` (and the prior `1.06 → 1.0.10` rename trick is captured in the same branch).
  3. **1.1 install** → full data migration via `Migrate_Database`, with the new schema pre-built in a `.tmp` file.
  4. **1.2 install** → in-place `Update_Database 1.3` (skipping 1.2 → directly add the new reports/modules to bring it current).
  5. **1.3+** → no-op, "already updated".

  This is a thoughtful upgrade strategy: small changes get ALTERs, big changes get table-by-table copy with logging. The backup-before-migrate pattern (`.mdb.bak` or `.bak` file) is proper data-safety practice.

- **The `.tmp` database is pre-created where?** The migration sub references `App.Path\Data\<companyID>.tmp` and writes to it via `OpenData` — but nothing in `Migrate_Database` itself creates this file. Either it's shipped alongside the executable (a known-good 1.2 schema template), or `CreateData`/`CreateDB` from `frmDatabase` (referenced in the database-not-found branch) handles it. The latter would only fire on first install, so for an upgrade the `.tmp` file must be installed by whatever ships the new version. Worth confirming with the deployment story.

- **The `dblUserDBVersion < 1.1` branch is misleading.** It says "Updating Database..." and runs `Update_Database 1.1`, but `Update_Database 1.1` is actually documented as "Version 1.06 to 1.0.10" in its body. So the version label inside `Update_Database` doesn't quite match the version it brings the database up to. Possibly the result of multiple iterations — the function got renamed but the comments didn't catch up.

- **The chained-update pattern.** A 1.0 install would: `Update_Database 1.1` brings it to 1.1 → restart → `Migrate_Database` brings it to 1.2 → `Update_Database 1.2` finalizes 1.2 → restart → `Update_Database 1.3` brings it to 1.3. Or rather, in a single splash run, only one branch fires per restart. So a deeply old database needs three sequential application launches to fully upgrade. Each one shows progress, succeeds, and the user must close-and-reopen. This is the stated pattern from the commented-out "Please restart program" message that would have shown after each upgrade.

- **No "are you sure?" before migration.** First boot after upgrade: the migration just runs. The user sees status messages but isn't asked to confirm. The backup-before is the safety net. For an internal hotel app this is fine; for distributed software this would be risky.

- **`Log2File` is used during migration but `LogErrorDB` is used in the catch.** Once the migration writes its first row to the new database, future `LogErrorDB` calls would land in the new database — appropriate. But the migration *itself* logs to a text file, which is the right call given the database is in flux.

- **Default brand "StarHotel".** This is a baked-in default that appears in the file naming convention. The system was clearly originally built for one customer, then made generic via the `gstrCompanyID` indirection. The constant lives on as the default value if no config is set.

- **`COMPANY_PRODUCT_NAME` is a global constant** displayed as the company-product label. Since it's referenced across most of the forms (Dashboard, Booking, Module Access), the rebrand is a single recompile to push to all forms.

- **Two parallel SQL helper sets.** `gstrSQL` + `QuerySQL` for the main DB; `QueryDataSQL` for the migration's dual-database scenario. The `OpenData` / `QueryDataSQL` family lets the migration sub talk to the destination database without disturbing the main one — same pattern that `frmPrint` uses for its private connection.

- **Salt + GoldFishEncode is the password story** — visible in the migration code as `Salt` and `UserPassword` fields being copied across. As discussed in `frmAdmin`, this is hand-rolled and weak by modern standards. The migration preserves whatever was there; it doesn't upgrade the hash to a stronger algorithm. If this system were modernized, that's where stronger crypto would land.

- **`Me.Refresh` after every `lblStatus.Caption` update** forces a synchronous redraw before the next phase. Without this the status would only repaint when Windows decided to, which during heavy DB work could be never. Standard splash-screen technique.

- **The flow is unidirectional.** Once the splash starts loading, there's no "Cancel" button, no way to stop. If you realize you've selected the wrong database, you have to wait for failure (or kill the process). Acceptable for an internal tool but worth noting.

- **`frmDatabase` shows modally** when the splash needs config, then control returns to splash — but actually no: looking again, when the database setup dialog is needed, `Timer1` is disabled, the splash unloads, and `frmDatabase.Show vbModal`. After that modal returns, control goes back into `frmDatabase`'s OK handler which calls `frmSplash.Show` again. So the flow is splash → splash needs DB → unload splash, show frmDatabase modal → frmDatabase OK → re-show splash → splash succeeds → login. A round trip rather than a continuation. Clean enough.

- **`lblWarning`'s "full version" text is misleading** if the kill-switch were ever re-enabled. The label says "This software is full version. Please check license for more information." — ironic given the commented-out time-bomb above it.
