# Bugs, Security Flaws & Mitigations

A consolidated catalog of bugs, half-implemented features, security flaws, and architectural smells in the original `StarHotel` VB6 codebase under `ORIGINAL-CODE/star-hotel-vb6-main/`. Each finding includes its location, severity, what's wrong, and a concrete mitigation. The catalog is sourced from the long-form analyses in `docs/forms/`, `docs/modules/`, and `docs/project/`, with the highest-impact claims verified against the original source.

This document is intended as a working list for the eventual rewrite — what must be fixed, what can be deleted, what to leave alone, and what to design out of the next architecture.

## Severity guide

| Severity | Meaning |
|---|---|
| **Critical** | Direct compromise of authentication or data integrity, or features so broken they fail their stated purpose entirely. Fix before any production use. |
| **High** | Significant security weakening, broken-but-recoverable feature, or design choice that will block modernization. Fix in any rewrite. |
| **Medium** | Incorrect behavior, latent bugs, defense-in-depth gaps, or smell that is workable today but should not survive a rewrite. |
| **Low** | UX rough edges, dead code, naming inconsistencies, cosmetic issues, or minor code smells. Optional. |

## Summary by category

| Category | Findings |
|---|---|
| 1. Security: Authentication & Credentials | 16 |
| 2. Security: SQL Injection & Input Sanitization | 9 |
| 3. Security: Authorization & Permissions | 12 |
| 4. Disabled or Half-Implemented Features | 16 |
| 5. Data Integrity & Concurrency | 13 |
| 6. Error Handling & Logging | 11 |
| 7. UI/UX Bugs | 22 |
| 8. Architectural Smells | 17 |
| 9. Build / Deployment Risks | 10 |
| 10. Dead Code & Inconsistencies | 24 |
| **Total** | **150** |

---

## 1. Security: Authentication & Credentials

### Hand-rolled, reversible "GoldFishEncode" used as password hash — Critical
- **Location:** `Module/modEncryption.bas:GoldFishEncode`
- **Description:** The function feedback-XORs each input byte against a rolling state seeded with the password length, producing a hex-encoded output. Given output + password length, the input is fully recoverable — it is an obfuscation, not a one-way hash. There is no key stretching, so brute-forcing a leaked database is trivial on commodity hardware.
- **Mitigation:** Replace with a vetted KDF (bcrypt, scrypt, or argon2id via Win32 BCrypt API or .NET interop). Plan a migration that re-hashes on next login since the existing hashes cannot be transparently upgraded.

### Default `admin/admin` credentials seeded and advertised on login screen — Critical
- **Location:** `Module/modDatabase.bas:CreateSampleData`, `frmUserLogin.lblDemo` (line 401)
- **Description:** A fresh database seeds `admin/admin` and `clerk/clerk` with `ChangePassword = False`, so they work immediately. The login screen displays "User ID: admin / Password: admin" via `lblDemo`.
- **Mitigation:** Force `ChangePassword = True` on seeded accounts; remove `lblDemo` for production builds; consider gating the demo hint behind a build flag.

### Login copyright label auto-fills admin credentials — Critical
- **Location:** `frmUserLogin.lblCopyright_Click` (lines 595–598, verified)
- **Description:** Clicking the copyright text on the login form sets `txtUserID = "ADMIN"`, `txtPassword = "admin"`. Combined with seeded defaults this is a one-click admin login if the password was never changed.
- **Mitigation:** Delete `lblCopyright_Click`. Even removing the seeded defaults doesn't fix this — the auto-fill is a vulnerability of its own.

### Salt length silently capped at 6 and halved — High
- **Location:** `Module/modEncryption.bas:GenSalt` (lines 27–28, verified)
- **Description:** `If StringLen > 6 Then StringLen = 6` then `StringLen = StringLen \ 2`, then iterates that many times, each iteration emitting 1–2 hex chars. Callers requesting 4, 6, 10, or 16 silently get the same 3-iteration result — actual salts are 2–6 chars wide. The schema reserves `TEXT(50)` so the cap is purely in code.
- **Mitigation:** Honor the requested length (or rename to make the cap explicit), and increase to ≥16 bytes generated via `CryptGenRandom`/`BCryptGenRandom` rather than `Rnd()`.

### No iteration / key stretching on password hash — High
- **Location:** `Module/modEncryption.bas:GoldFishEncode`
- **Description:** Hash evaluation is O(n) in password length with trivial constant cost. A modern GPU evaluates billions per second, so a leaked `UserData` row falls to brute force in seconds — combined with the 8–10 character password cap.
- **Mitigation:** Adopt bcrypt/argon2 with cost factor tuned to ~100ms per evaluation.

### `MaxLength = 10` (or 8) on password input fields — High
- **Location:** `frmUserLogin.txtPassword`, `frmUserChangePassword.txtPassword*`, `frmAdmin.txtPassword`, `frmUserMaintain.txtPassword*`
- **Description:** All password textboxes cap input at 10 characters (user-facing change/login forms) or 8 characters (admin-creates-user). A 10-character ceiling is weak by 2026 standards and dramatically shrinks the keyspace `GoldFishEncode` operates on.
- **Mitigation:** Raise `MaxLength` to ≥64. Reconcile the 8 vs 10 mismatch between admin-create and user-change paths.

### Inconsistent salt length across callers (4 vs 6) — Medium
- **Location:** `frmUserChangePassword`, `frmUserMaintain.SaveUser`, `Module/modDatabase.bas:CreateSampleData`
- **Description:** `GenSalt(4)` is used by user-driven password changes/creation while `GenSalt(6)` is used by seed data. Combined with the cap-and-halve bug, real-world salts vary 2–4 vs 3–6 chars. No security difference at this scale, but it shows there is no defined salt-length policy.
- **Mitigation:** Standardize all callers on a single helper producing a sufficient-length cryptographic salt.

### Plaintext password hash and salt held in module globals for the entire session — Medium
- **Location:** `Module/modGlobal.bas:gstrUserPassword`, `gstrUserSalt` set in `frmUserLogin.cmdOK_Click`
- **Description:** After login, the encoded password hash and salt remain in `gstrUserPassword`/`gstrUserSalt` for the whole session — even after explicit logout (the `frmDialog` cleanup never clears them). A memory dump or process inspection leaks them. They are read in only one place (`frmUserChangePassword`) and could plausibly be removed entirely.
- **Mitigation:** Clear all session globals on logout. Consider not stashing them at all.

### Hardcoded Jet OLE DB password derived from company name — Medium
- **Location:** `Module/modCommon.bas:GenWord`, set in `frmSplash.Timer1_Timer`
- **Description:** `GenWord` rebuilds the literal "Computerise" from ASCII codes to keep it out of `strings(1)` output. The decoded string is then used as the Jet password for every install. Anyone who decompiles the binary or reads the function recovers the password trivially.
- **Mitigation:** Acceptable for low-threat opportunistic snooping but inadequate for any data-protection requirement. Move to per-install random keys stored via DPAPI, or use a stronger storage mechanism.

### `Randomize` re-seeded inside salt-generation loop — Low
- **Location:** `Module/modEncryption.bas:GenSalt` (line 30, verified)
- **Description:** `Randomize` is called every iteration inside the `For` loop. VB6's `Randomize` re-seeds from the system clock; consecutive calls within the same millisecond produce the same seed. Combined with `Rnd()` (a non-CSPRNG) for security material, salts have low entropy.
- **Mitigation:** Replace `Rnd()` with `CryptGenRandom`. If staying with `Rnd()`, call `Randomize` once outside the loop.

### `Dec2Bin(0)` returns empty string — Low
- **Location:** `Module/modEncryption.bas:Dec2Bin` (verified)
- **Description:** The `Do Until dec = 0` loop never runs, then `Len("") Mod 8 = 0` is true so padding is skipped. Only triggerable if a password contains `Chr(0)`, which input fields don't allow — but a latent correctness bug.
- **Mitigation:** Add an explicit `If dec = 0 Then Dec2Bin = "00000000": Exit Function` guard.

### High-bit drop on non-ASCII password characters — Low
- **Location:** `Module/modEncryption.bas:GoldFishEncode` / `Dec2Bin`
- **Description:** For `Asc(char) > 255`, `Dec2Bin` pads to 16/24 bits but the inner XOR loop only iterates 8 bits, silently dropping high-order bits. Accented or Unicode passwords would have collisions.
- **Mitigation:** Either validate-and-reject non-ASCII input or process the full bit-width.

### No constant-time hash comparison — Low
- **Location:** `frmUserLogin.cmdOK_Click` (`If strPassword = gstrUserPassword Then`)
- **Description:** VB6 string compare short-circuits on first mismatch. Irrelevant for a local single-user app but a timing-attack vector in any networked scenario.
- **Mitigation:** If exposed to network traffic, use a fixed-length XOR-and-OR comparison.

### `gstrPassword` (Jet password) held as plaintext global for whole session — Low
- **Location:** `Module/modDatabase.bas:gstrPassword`, set in `frmSplash`
- **Description:** The decoded Jet password lives in memory for the lifetime of the process. Acceptable for a local kiosk but worth noting for any process-dump exposure.
- **Mitigation:** Use SecureString-style storage (zero-on-free) if the threat model expands.

### MsgBox-on-error inside cryptographic primitives — Low
- **Location:** `Module/modEncryption.bas:Encrypt`, `GenSalt`, `GoldFishEncode`, `Dec2Bin` (verified)
- **Description:** Each function has `On Error GoTo CheckErr` → `MsgBox Err.Number & ...`. A failure inside hashing pops a dialog with internal details mid-login.
- **Mitigation:** Re-raise to the caller or return a sentinel. Never MsgBox from a security primitive.

### Commented-out `imgIcon_DblClick` in `frmUserLogin` would have launched change-password without auth — Low
- **Location:** `frmUserLogin.frm` (commented block)
- **Description:** The dead code, if re-enabled, lets anyone double-click an icon on the login screen and reach the password-change form without authenticating. Removal was correct; flagging that the dead code is still in the file.
- **Mitigation:** Delete the commented block.

---

## 2. Security: SQL Injection & Input Sanitization

### No parameterized queries anywhere in the codebase — Critical
- **Location:** `Module/modCommon.bas` (entire `SQL_*`/`SQLData_*`/`SQL_SET_*` family) and every form that calls them
- **Description:** Every value is concatenated into the SQL string. SQL safety relies entirely on `CheckString` (escape apostrophes, strip brackets/asterisks/ampersands) and `CheckInput`. Any path that bypasses these (and several do) is injectable.
- **Mitigation:** Migrate to `ADODB.Command` with `Parameters` collection. This is a large refactor since every form goes through `gstrSQL`.

### `CheckInput` reserved-word stripping is naive and corrupts data — High
- **Location:** `Module/modCommon.bas:CheckInput`
- **Description:** Strips substrings "password", "salt", "author", "code", "username", "select", "from" case-insensitively. A user named "Pselectter" becomes "Pter"; a city "Frome" becomes "e". The stripping does nothing for the documented test attack `1=1'or'1=1` (no banned words), so the actual SQL safety reduces to the apostrophe-doubling at the end.
- **Mitigation:** Remove the substring stripping entirely. Move to ADO `Command` parameterized queries; if that's too disruptive, at least keep only the apostrophe-doubling.

### Asymmetry: `SQLData_Text` sanitizes via `CheckString`, `SQL_SET_Text` does not — High
- **Location:** `Module/modCommon.bas` (the SQL builder family)
- **Description:** INSERT paths automatically run text through `CheckString`; UPDATE paths emit values raw. So updating a guest record with an apostrophe in the name fails or injects, while inserting succeeds. Several form callers (e.g., `frmRoomTypeMaintain.SaveRoomType`) compensate inconsistently — INSERT calls `CheckInput`, UPDATE doesn't.
- **Mitigation:** Either add `CheckString` to `SQL_SET_Text` (auditing callers that already escape, to avoid double-escaping) or move to parameterized queries.

### `PrintReceipt` and `PrintHistory` build SQL by raw concatenation — High
- **Location:** `frmBooking.PrintReceipt`, `frmFindCustomer.PrintHistory`
- **Description:** Both bypass the typed builder layer and concatenate strings directly. `gstrUserID` and customer-typed search fields are inserted unescaped. Apostrophes in user IDs or guest names break or inject.
- **Mitigation:** Route through `SQLText` after `CheckString`, or migrate to parameters. At minimum, escape inputs.

### Operator-precedence trap with user-chosen AND/OR connectors — High
- **Location:** `frmFindCustomer.ListCustomers`
- **Description:** Dynamic clauses like `Name LIKE 'a' OR Passport LIKE 'b' AND BookingDate BETWEEN ...` parse as `Name LIKE 'a' OR (Passport LIKE 'b' AND BookingDate ...)`. Any name match returns regardless of date or active flag.
- **Mitigation:** Wrap user-chosen criteria in an outer parenthesis: `(criteria1 AND/OR criteria2 ...) AND BookingDate ... AND Active = TRUE AND Temp = FALSE`.

### `PrintHistory` contact-filter branch corrupts `gstrSQL` — Medium
- **Location:** `frmFindCustomer.PrintHistory`
- **Description:** The function builds `gstrSQL` via local concatenation, then inside the contact filter calls `SQL_AND_LIKE_Text "GuestContact", strGuestContact` which rewrites `gstrSQL`. The resulting SQL is malformed.
- **Mitigation:** Either build the entire query via the helper layer or via concatenation — don't mix. Verify whether this branch ever executes.

### `$UserID$` and `$BookingID$` are plain string substitution in report SQL — Medium
- **Location:** `frmReport.PrintReport`
- **Description:** `Replace(gstrSQL, "$UserID$", gstrUserID)` is unescaped. A user ID with an apostrophe (or a malicious admin-set value) breaks reports. `$BookingID$` is sourced from an `InputBox` with no validation.
- **Mitigation:** Validate `$BookingID$` is numeric; escape `$UserID$` via `CheckString` before substitution.

### Direct string concatenation of `strUserID` in `modFunction` helpers — Medium
- **Location:** `Module/modFunction.bas:UserGroup`, `AdminUser`, `NeedChangePassword`, `UserAccessModule`
- **Description:** Each builds `WHERE UserID = '" & strUserID & "'"` with no `CheckInput`/`CheckString`. Safe in practice because callers pass `gstrUserID` (sanitized at login) or hardcoded values, but the function signature doesn't enforce that contract.
- **Mitigation:** Either sanitize at the function boundary or move to parameters.

### `CheckString` strips ampersands and asterisks indiscriminately — Low
- **Location:** `Module/modCommon.bas:CheckString`
- **Description:** "Lee & Associates" stored as "Lee  Associates"; "Note: see room *007*" loses asterisks. Legitimate punctuation lost in stored data.
- **Mitigation:** Move to parameters so wildcards/ampersands don't need stripping.

---

## 3. Security: Authorization & Permissions

### Hardcoded `expert` developer password gates report editing — High
- **Location:** `frmReport.frm:Form_Load` (`strDeveloperPassword = "expert"`), checked at the Edit and Save paths
- **Description:** The string is plaintext in compiled code, the prompt uses `InputBox` without masking, and combined with `MOD_REPORT_EDIT` it is the second of two gates protecting the SQL editor. Anyone with the source or who can guess the obvious password gets full edit access.
- **Mitigation:** Remove the developer-password gate entirely; rely on `MOD_REPORT_EDIT_EXPERT` permission. If retained, prompt with a password textbox (masked) and at least store a hash.

### Group 1 (Admin) has unlimited login attempts — Medium
- **Location:** `frmUserLogin.cmdOK_Click`, `frmAdmin.cmdOK_Click`
- **Description:** The wrong-password branch increments `LoginAttempts` only `If gintUserGroup > 1`. Group 1 admins have no rate-limit at the application or database level — an attacker brute-forcing the admin can keep trying indefinitely.
- **Mitigation:** Apply the same attempt counter to Group 1 (with a higher threshold if desired). Consider a global/IP-based rate limit too.

### `UserAccessModule` falls through unrecognized groups to Clerk permissions — Medium
- **Location:** `Module/modFunction.bas:UserAccessModule`
- **Description:** The if/elseif chain handles UserGroup 1/2/3 explicitly; everything else (including 5, 6, 99) hits the `Else` branch and inherits Group4 (Clerk) flags. A future migration adding a "Read-only Auditor" group at value 5 would silently grant Clerk-level access.
- **Mitigation:** Default-deny on unrecognized groups: explicitly check `UserGroup = 4` instead of relying on the catch-all `Else`.

### No "at least one admin must remain active" guard in user editor — Medium
- **Location:** `frmUserMaintain.SaveUser`
- **Description:** An admin can deactivate their own account, demote their own group, or freeze themselves. Combined with deactivating other admins, it's possible to reach a "no active admins" state from which no further admin actions can be taken without direct DB access.
- **Mitigation:** Block save if it would leave zero active admins, and block self-deactivation.

### Permission checks happen at form-load only, not per operation — Medium
- **Location:** Cross-cutting; documented in `docs/modules/DATA-FLOW-DIAGRAM.md`
- **Description:** If an admin disables a user's "Edit Report" permission while the user is inside that form, the user can still complete actions until they navigate away. There's no mid-session permission re-check.
- **Mitigation:** Re-check permissions before destructive operations (Save, Delete) in addition to form load.

### `gintUserIdle = mintIdle` in user save mutates the global for whichever user is being edited — Medium
- **Location:** `frmUserMaintain.SaveUser`
- **Description:** When admin edits Clerk's record, the admin's session global `gintUserIdle` is overwritten with the Clerk's value. Currently no observable impact because the watchdog is inert, but if the watchdog is fixed this becomes an active bug.
- **Mitigation:** Only update the global when `mstrUserID = gstrUserID`.

### No audit trail for `UserData` or `ModuleAccess` changes — Medium
- **Location:** `frmUserMaintain.SaveUser`, `frmModuleAccess.UpdateModuleAccess`
- **Description:** Unlike `LogRoom`/`LogBooking`, these tables have no history table. Who promoted a user to admin, who unlocked an account, who disabled a permission — none of it is recoverable.
- **Mitigation:** Add `LogUser` and `LogModuleAccess` audit tables and snapshot before update.

### Module ID renumbering when `MOD_BOOKING_VOID` was disabled is unmigrated — Medium
- **Location:** `Module/modGlobal.bas:MOD_*` constants; commented-out `MOD_BOOKING_VOID = 3`
- **Description:** The constant was deleted and `MOD_REPORT_LIST` reused slot 3, shifting subsequent IDs. No migration in `Update_Database` exists for customers whose `ModuleAccess` rows still reflect the old numbering — their permissions silently point at wrong modules after upgrade.
- **Mitigation:** Add an idempotent `Update_Database` step that detects and remaps old IDs.

### Two of four user groups (Group2, Group3) defined but unwired in UI — Low
- **Location:** `frmModuleAccess.frm`, `Module/modDatabase.bas:CreateSampleData` (4 UserGroups, 2 inactive)
- **Description:** Schema/checkbox columns Group2/Group3 exist but have no UI representation. The seed marks Manager and Supervisor as `Active = False`. If anyone activates them via direct DB edit, no UI surfaces those permissions.
- **Mitigation:** Either remove the columns and seeded inactive groups, or wire them into `frmModuleAccess`.

### `frmModuleAccess` writes immediately with no Save/Cancel/confirmation — Low
- **Location:** `frmModuleAccess.chkGroup1_Click`, `chkGroup4_Click`
- **Description:** Every checkbox click is a live UPDATE. An accidental click is permanent and unlogged.
- **Mitigation:** Add a Save/Cancel dirty-check workflow, or at minimum a confirmation MsgBox.

### `Report.ReportID + 11` magic offset for permission lookups — Low
- **Location:** `frmReport.LoadReports` (`UserAccessModule(rst!ReportID + 11)`)
- **Description:** Reports are numbered 1–7; permissions live at ModuleAccess IDs 12–17. If anyone inserts a new module-level permission and shifts numbering, all report permissions break silently.
- **Mitigation:** Replace the offset with a `Report.PermissionModuleID` foreign-key column.

### Module ID 18 ("Official Receipt") has no `MOD_*` or `REP_*` constant — Low
- **Location:** `Module/modGlobal.bas`
- **Description:** `frmReport` permission gating consumes ReportID + 11. The constants list ends at `REP_SHIFT_ALL_USER (17)`. The 18th ModuleAccess row is referenced only via raw integer if at all.
- **Mitigation:** Add a `REP_OFFICIAL_RECEIPT = 18` constant and use it where appropriate.

---

## 4. Disabled or Half-Implemented Features

### Idle-logout watchdog disabled by debug override — High
- **Location:** `frmUserLogin.cmdOK_Click` line 474, verified: `gintUserIdle = 0 '10 ' Test`
- **Description:** Right after the database loads `Idle` into `gintUserIdle`, this line zeroes it. Every form's `tmrClock_Timer` gates with `If gintUserIdle > 0 Then` so the body never runs. The entire idle-logout feature, the `frmDialog` warning, and per-user idle settings are inert system-wide. A subsequent sanity-check `If gintUserIdle > 3600 Or gintUserIdle < 0 Then gintUserIdle = 0` runs after the override, so it's also moot.
- **Mitigation:** Delete line 474. Test the full idle-logout chain end to end.

### Beta-version kill-switch commented out but would brick all installs if re-enabled — High
- **Location:** `frmSplash.Timer1_Timer`
- **Description:** Two `If intDay < 0 Or intDay > N Then End` blocks (N=99 and 364) anchored at "01 Jan 2021" are commented out. As of 2026-05-03, `intDay` is ~1949 — re-enabling either branch instantly disables every install. The commented-out `lblDemo` text in `frmUserLogin` ("This free demo is valid until ...") references the same anchor.
- **Mitigation:** Delete the dead kill-switch code rather than leaving it commented. If a time-bomb is needed for trial builds, use `#If` conditional compilation.

### `frmUserMaintain` Confirm field is decorative (password not validated against it) — High
- **Location:** `frmUserMaintain.SaveUser`
- **Description:** Three password fields (Old/New/Confirm) but only `txtPasswordNew` is consumed by the UPDATE/INSERT. Admin can type "abc" in New and "xyz" in Confirm — the user's password becomes "abc" with no warning. Compare to `frmUserChangePassword`, which correctly checks New == Confirm.
- **Mitigation:** Add `If txtPasswordNew.Text <> txtPasswordConfirm.Text Then ... Exit Sub`.

### `frmUserMaintain` INSERT path has no empty-password guard — High
- **Location:** `frmUserMaintain.SaveUser` (INSERT branch)
- **Description:** New-user creation calls `Encrypt(txtPasswordNew.Text, mstrSalt)` unconditionally. If the admin doesn't tick `chkUpdatePassword`, the disabled field's text (likely empty) becomes the password. Anyone who knows the salt can compute the hash and discover the password is empty. Login does block empty input, but the hash sits on disk vulnerable to brute-force.
- **Mitigation:** Block INSERT when `chkUpdatePassword` is unchecked or `txtPasswordNew.Text` is empty.

### `frmAdmin` Void step-up authentication is a no-op stub — Medium
- **Location:** `frmAdmin.cmdOK_Click`, `frmBooking` Void handlers
- **Description:** `frmAdmin` authenticates the user but the block that performs the Void (`frmBooking.VoidBooking lblBookingID`) is fully commented out. A successful login closes nothing, performs no action, and provides no return signal to the caller. The matching `frmAdmin.lblBookingID.Caption = ...; frmAdmin.Show vbModal` calls are commented out in `frmBooking`. The whole Void feature is inert in two places.
- **Mitigation:** Either uncomment all four blocks across the two files (and fix the freeze-account `Active = TRUE` bug — see UI/UX section) or remove `frmAdmin` entirely.

### Save-doesn't-insert in `frmReportMaintain` — Medium
- **Location:** `frmReportMaintain.SaveReport`
- **Description:** The UPDATE is built inside `If Not rst.EOF Then`, so if the row doesn't exist (e.g., a "create new" path) the SAVE silently does nothing — no INSERT, no error message.
- **Mitigation:** Add an INSERT branch, or display "row not found, please refresh" when EOF.

### No password-strength enforcement at user creation — Medium
- **Location:** `frmUserMaintain.SaveUser`
- **Description:** Length check (`Len < 4`) lives only in `frmUserChangePassword`. Admin can create a 1-character password.
- **Mitigation:** Mirror the change-password validation in the create-user path; ideally enforce a uniform policy.

### `CompactDB` fully commented out — no Access database compaction routine — Medium
- **Location:** `Module/modDatabase.bas` (bottom)
- **Description:** The `JRO.JetEngine.CompactDatabase` call is commented out, apparently due to a Crystal Reports `P2smon.dll` interference. Customers running this for years see steady `.mdb` growth with no maintenance routine.
- **Mitigation:** Re-implement as an offline maintenance sub that fully closes Crystal first, or schedule compaction as an external process.

### Void button on `frmBooking` is wired but disabled — Low
- **Location:** `frmBooking.tbrMenu` VOID button, `frmBooking.VoidBooking`
- **Description:** `VoidBooking` and the caption-swap in `PopulateValues` are fully written; the toolbar/keyboard handlers are commented out. The feature is one boolean flip away from being live.
- **Mitigation:** Decision required — enable (with the rest of the chain) or remove.

### Hidden DELETE buttons on Maintain forms — Low
- **Location:** `frmRoomTypeMaintain.tbrMenu("DELETE")`, `frmUserMaintain.tbrMenu("DELETE")`, `frmReportMaintain.tbrMenu("PRINT")`, others
- **Description:** DELETE buttons (and `frmReportMaintain`'s PRINT button) are defined with `Enabled = False, Visible = False`. The corresponding handlers (`DeleteRoomType`, `DeleteUser`) are commented out, and at least one body has been corrupted by copy-paste from another form (references `cboUserGroup` etc. that don't exist on this form).
- **Mitigation:** Remove the dead controls and code, or reimplement properly with safety checks.

### `CloseAllForms` defined in `frmDialog` but commented out — Low
- **Location:** `frmDialog.cmdOK_Click`, `frmDialog.tmrClock_Timer`
- **Description:** The hardcoded list of `Unload frm*` calls is duplicated in two places; the more general `CloseAllForms` helper is defined but the call is commented out in both.
- **Mitigation:** Wire up `CloseAllForms` so future-added forms don't need cleanup-list updates.

### `Last Occupied Date` column displayed nowhere — Low
- **Location:** `frmRoomMaintain` (`Visible = False` field)
- **Description:** Schema reserves space, layout positions a label, but `Visible = False`. Probably staged for a "rooms not used recently" report that never shipped.
- **Mitigation:** Either expose or remove.

### Bilingual `*1`/`*2` schema columns never wired — Low
- **Location:** `Report.ReportName1/Title1/AsOn1/DateField1/DateType1`, `ModuleAccess.ModuleDesc1`, `Company.F01`, `gstrCompanyName2`
- **Description:** Schema has `_1`/`_2` pairs suggesting English/Malay support; only `_1` is read. `Company.F01` is unused; `gstrCompanyName2` is declared but never assigned from the schema (no `CompanyName2` column exists).
- **Mitigation:** Either implement the second language or remove the dead schema/global.

### Theme-related code commented out in `frmUserMaintain` — Low
- **Location:** `frmUserMaintain` (`gintUserThemeID`, `cboTheme`)
- **Description:** Per-user themes were planned, never shipped. Comments allude to a `UserThemeID` lookup that doesn't exist.
- **Mitigation:** Remove the dead code.

### `PRINT` button on `frmReportMaintain` defined invisible — Low
- **Location:** `frmReportMaintain.tbrMenu("PRINT")`
- **Description:** Wired into `ImageList(4)` with `Enabled = False, Visible = False`. Intent was preview-from-editor; never shipped.
- **Mitigation:** Remove.

---

## 5. Data Integrity & Concurrency

### Multi-step writes are not wrapped in transactions — High
- **Location:** `frmBooking.SaveBooking`, `frmBooking.Check_OUT`, `frmRoomMaintain.SaveRecord`, others
- **Description:** `QuerySQL` uses `BeginTrans`/`CommitTrans` per single statement only. A multi-statement save (UPDATE Booking, INSERT LogBooking, UPDATE Room) is not atomic. A crash mid-sequence leaves partial state. No form ever calls `ACN.BeginTrans`/`CommitTrans` itself.
- **Mitigation:** Add explicit transaction wrapping for multi-statement business operations.

### Temp booking row pattern is racy in multi-user scenarios — Medium
- **Location:** `frmBooking.CreateTempBookingID`, `GetTempBookingID`
- **Description:** Opening a room button creates a `Temp = TRUE` Booking row immediately. Closing without saving leaves orphans. `GetTempBookingID` (`TOP 1 ORDER BY ID DESC`) reuses orphans on next open. Two staff opening different rooms simultaneously could race onto the same temp row.
- **Mitigation:** Filter by current user's `CreatedBy`; mark orphans aged > N hours as cleanup candidates; or never create a row until first save.

### `Room.ID` is non-autoincrement, caller-supplied — Medium
- **Location:** `Module/modDatabase.bas:CreateDB` (`SQL_COLUMN_ID "ID", True, False, True` for Room)
- **Description:** Room IDs come from the Dashboard button index. Inserting a new room at an existing ID (after soft-delete) would conflict. The bind to physical button layout is fragile.
- **Mitigation:** If physical layout matters, store the layout slot in a separate column and let `Room.ID` autoincrement. Otherwise, document the contract clearly.

### No foreign-key constraints anywhere in schema — Medium
- **Location:** `Module/modDatabase.bas:CreateDB`
- **Description:** `Booking.RoomID` references `Room.ID` only by convention. Orphans are technically possible. The application enforces integrity via soft delete and pre-checks, but DB-level enforcement is absent.
- **Mitigation:** Add explicit `FOREIGN KEY` constraints in `CreateDB`. Consider cascade soft-delete logic.

### Renaming `RoomType.TypeShortName` orphans existing rooms — Medium
- **Location:** `frmRoomTypeMaintain.SaveRoomType`, joined to `Room.RoomType` by name
- **Description:** Room and Booking store RoomType as text (denormalized). Renaming a type doesn't propagate; rooms now reference a non-existent type and `cboRoomType.ListIndex` falls back to -1 (blank) on next load.
- **Mitigation:** Block renames via `IsRoomTypeUsed`, or cascade the rename to all referencing rows in a transaction.

### `UpdateWeekDayTable` mutates a shared table as a side effect of running a report — Medium
- **Location:** `frmReport.UpdateWeekDayTable`
- **Description:** `WeeklyBooking` has 7 rows that get UPDATEd to the report-date's week. Two users running the report simultaneously race; the loser sees inconsistent buckets.
- **Mitigation:** Generate the date series in-query (use a derived table approach in Jet, or compute in VB), avoid mutating shared state per report run.

### Audit-trail gaps: no `LogUser`, no `LogModuleAccess`, no `LogReport` — Medium
- **Location:** Cross-cutting
- **Description:** Only `LogRoom` and `LogBooking` exist. Changes to user accounts, permissions, or report definitions leave only `LastModifiedBy` (no historical "who changed what to what when").
- **Mitigation:** Add audit tables for these entities.

### Denormalized booking snapshot creates double-source-of-truth — Low
- **Location:** `frmBooking.PopulateValues` reads room data from `Booking`; `GetRoomDetails` reads from `Room`
- **Description:** Booking stores its own RoomType, RoomLocation, RoomPrice, Breakfast snapshots — appropriate for receipts but means edits to Room don't propagate to existing bookings. A user reading "what was this booking's room like at booking time" gets correct data; reading "what is the current room called" requires a different query.
- **Mitigation:** Document the contract; if any UI conflates the two, make it consistent.

### `MonthDay30` misnamed but math is correct — Low
- **Location:** `Module/modCommon.bas:MonthDay30`
- **Description:** Despite the name, returns the actual last day of the month (28/29/30/31). However, `frmFindCustomer` defaults the search-to date to `MonthDay30(Now)`, which would exclude future bookings beyond month-end (hotels often book months ahead).
- **Mitigation:** Rename to `MonthLastDay`. Reconsider `frmFindCustomer.ResetFields` default range.

### Dead 2:00 PM branch in `frmBooking.dtpCheckOutTime_Change` — Low
- **Location:** `frmBooking.dtpCheckOutTime_Change`
- **Description:** A `>= 12:00 PM` branch always matches before the 2:00 PM branch can fire — leftover from previous policy logic. The "no refund if checkout >= 14:00" rule is never enforced.
- **Mitigation:** Restore intent or remove dead code.

### Soft-delete inactive room types still listed in cboRoomType after restart — Low
- **Location:** `frmRoomMaintain.PopulateRoomType`
- **Description:** Filter is `WHERE Active = TRUE`. Rooms whose stored type was deactivated end up with `cboRoomType.ListIndex = -1` (blank), silently disconnecting them from a type. No warning surfaced.
- **Mitigation:** Either keep deactivated types selectable for existing rooms, or warn when loading a room whose type is no longer active.

### No write-permission check at chosen DB path — Low
- **Location:** `frmDatabase.cmdOK_Click`
- **Description:** A user-typed path that exists but isn't writable surfaces a generic ADO error from `CreateDB` rather than upfront guidance.
- **Mitigation:** Probe writability before commit.

### `Config.txt` two-line format has no escaping — Low
- **Location:** `frmDatabase.cmdOK_Click`, `Module/modTextFile.bas:WriteTextFile`
- **Description:** `path + CRLF + filename` only works if neither contains an embedded newline. Edge case but no defense.
- **Mitigation:** Use a structured format (INI/JSON) or quote values.

---

## 6. Error Handling & Logging

### Error handlers `MsgBox` first, log second — including in low-level helpers — Medium
- **Location:** Almost every error handler across the codebase
- **Description:** Pattern is `MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod` then `LogErrorDB` or `LogErrorText`. A user with no context gets a raw ADO/COM error code dialog. From the encryption module, this surfaces internal failures during login.
- **Mitigation:** Swallow expected error categories; show user-friendly messages for the rest; always log first.

### Logger-shows-MsgBox cascade on file-I/O failure — Medium
- **Location:** `Module/modTextFile.bas:Log2File`, `LogErrorText`
- **Description:** A failed write inside the logger pops a MsgBox. If invoked from another error handler, the user sees two MsgBoxes from one root cause.
- **Mitigation:** Logger should fail silently or fall back to `Debug.Print`.

### `SelectField`/`UpdateField` don't manage their own connections — Medium
- **Location:** `Module/modFunction.bas:SelectField`, `UpdateField`
- **Description:** Unlike sibling helpers (`UserGroup`, `AdminUser`), these expect the caller to have already called `OpenDB`. Inconsistent — a developer calling either without an outer `OpenDB` crashes. No naming convention indicates this.
- **Mitigation:** Either make them self-managing or rename to communicate the contract.

### `LogErrorText` uses hardcoded file handle `#1` — Low
- **Location:** `Module/modTextFile.bas:LogErrorText`
- **Description:** Every other writer uses `FreeFile`. Hardcoded `#1` collides if any caller has another file open on that handle.
- **Mitigation:** Use `FreeFile`.

### `Close` (without handle) in `ReadTextFile`/`ReadFile` closes all open files — Low
- **Location:** `Module/modTextFile.bas`
- **Description:** Bare `Close` shuts every open handle in the process, not just the one this function opened. Safe in practice because callers don't nest, but a hazard.
- **Mitigation:** Use `Close #FF`.

### Error-log file grows unboundedly (no rotation) — Low
- **Location:** `Module/modTextFile.bas`
- **Description:** No log rotation, no size cap. Long-running installs accumulate `Error.txt` indefinitely.
- **Mitigation:** Add rotation/truncation policy.

### `LogErrorDB`'s catch block fails to `CloseDB` after a failed `OpenDB` — Low
- **Location:** `Module/modFunction.bas:LogErrorDB`
- **Description:** If the inner `OpenDB` fails, the function swallows the error and falls back to text logging — but does not call `CloseDB` to clean up partial state.
- **Mitigation:** Add `CloseDB` to catch.

### Several `_DB` log calls are commented out, leaving only text fallback — Low
- **Location:** `Module/modDatabase.bas` (in `OpenDB`, `ConnectDB`, `CloseDB`, `CreateData`, `CreateDB`, `CreateSampleData`, `OpenSQL`, `OpenQuery`, `OpenRS`, `OpenTable`, `CloseRS`); `modEncryption.bas` (all functions)
- **Description:** Intentional in some cases (the DB might not be open), but inconsistently muffles logging from layers where the DB *is* available.
- **Mitigation:** Audit each commented `LogErrorDB` and decide: if the function is post-connection, enable it.

### `LogError` table missing per-call user/SQL context for some paths — Low
- **Location:** `Module/modFunction.bas:LogErrorDB`
- **Description:** Only `gstrUserID` is stamped automatically; the failed SQL is included only if the caller passes it via `Err.Description`. Inconsistent depending on call site.
- **Mitigation:** Add an explicit `pstrSQL` parameter passed by callers from `QuerySQL`/`OpenRS`.

### `UpdateField` returns `ADODB.Recordset` but never sets a return value — Low
- **Location:** `Module/modFunction.bas:UpdateField`
- **Description:** Function signature claims a return type that's never assigned. Callers would always receive Nothing.
- **Mitigation:** Change to `As Boolean` or just `Sub`.

### `UpdateField` uses magic-string sentinel `"INCREMENT_ONE"` — Low
- **Location:** `Module/modFunction.bas:UpdateField`, called from `frmUserLogin.cmdOK_Click`
- **Description:** A literal string `"INCREMENT_ONE"` is intercepted to mean `column = column + 1`. Adding more sentinels would be brittle; the API hides intent.
- **Mitigation:** Add a separate `IncrementField` function.

---

## 7. UI/UX Bugs

### `frmUserLogin` "freeze account" sets `Active = TRUE` instead of `FALSE` — High
- **Location:** `frmUserLogin.cmdOK_Click` line 494 (verified): `UpdateField "UserData", "Active", "BOOLEAN", "TRUE", " UserID = '" & CheckInput(txtUserID.Text) & "'"`; same bug in `frmAdmin.cmdOK_Click`
- **Description:** After 3+ DB-persisted attempts, the freeze branch sets `Active = TRUE`, leaving the account unfrozen. The displayed message claims the account is frozen, but it isn't. Combined with the local `intAttempt` guard that `End`s the app, the user can simply relaunch.
- **Mitigation:** Change to `"FALSE"`.

### `End` used as the response to many failure modes — Medium
- **Location:** `frmUserLogin` (frozen, too many attempts, cancel), `frmUserChangePassword.cmdCancel_Click` (forced path)
- **Description:** `End` immediately terminates without orderly cleanup of open recordsets, connections, or COM references. A typo on the login screen → restart the app.
- **Mitigation:** Use `Unload`/`Show frmUserLogin` patterns; reserve `End` for true forced-quit cases.

### `frmDialog` default button is the destructive action — Medium
- **Location:** `frmDialog.cmdOK` (`Default = True`)
- **Description:** The auto-logout warning has OK on the left as default. A user with their finger on Enter (perhaps from typing in the form they were on) triggers immediate logout.
- **Mitigation:** Make Cancel the default.

### `frmDialog` hardcoded form-list misses `frmFindCustomer` — Medium
- **Location:** `frmDialog.cmdOK_Click`, `frmDialog.tmrClock_Timer`
- **Description:** Eight `Unload` calls are hardcoded but `frmFindCustomer` is reachable from Dashboard via F3 and is not in the list. Auto-logout while on FindCustomer leaves it open after Dashboard unloads, with the login screen appearing underneath.
- **Mitigation:** Add `Unload frmFindCustomer`. Better, wire up `CloseAllForms` and stop hardcoding.

### Session globals not cleared on logout — Medium
- **Location:** `frmDialog.cmdOK_Click` and unload chain
- **Description:** After logout returns to login, `gstrUserID`, `gintUserGroup`, `gstrUserPassword`, `gstrUserSalt`, etc. retain previous user's values until next login overwrites them.
- **Mitigation:** Clear all session globals on logout.

### Off-screen layout: `Form_Resize` math in `frmUserChangePassword` puts panel above center — Low
- **Location:** `frmUserChangePassword.Form_Resize`
- **Description:** `(Me.Height - fraLogin.Height) \ 2 - fraLogin.Height` results in the panel near the top of the screen, not centered. Probably deliberate for kiosk displays but unusual.
- **Mitigation:** Verify intent; fix arithmetic if center is desired.

### `frmDialog` countdown grammar awkward, no plural handling — Low
- **Location:** `frmDialog.tmrClock_Timer`
- **Description:** "System will be log out automatically after 1 second(s)" — grammar wrong ("logged out") and parenthetical "(s)" handles plurals lazily.
- **Mitigation:** Fix copy and add `If intSecondsLeft = 1 Then "second" Else "seconds"`.

### `frmUserLogin` reveals which credential was wrong; `frmAdmin` doesn't — Low
- **Location:** `frmUserLogin.cmdOK_Click` ("User ID not found!"), `frmAdmin.cmdOK_Click` (silent fail on unknown user)
- **Description:** Inconsistent UX and security stance. Login leaks user existence; admin step-up doesn't.
- **Mitigation:** Pick one. For internal app, "Invalid User ID or Password" without distinguishing is the conventional choice.

### `frmReport.PrintReport` no date-range validation — Low
- **Location:** `frmReport.PrintReport`
- **Description:** A future date for a Yearly report yields data for that future year. No upper bound check.
- **Mitigation:** Cap at `Now()` or warn.

### `frmFindCustomer.PrintHistory` uses typed field for contact filter but selected ListView item for passport — Low
- **Location:** `frmFindCustomer.PrintHistory`
- **Description:** Print uses `txtGuestContact.Text` (typed), not the selected customer's actual stored contact.
- **Mitigation:** Read from the selected ListView row consistently.

### `SelectRoomType` always called with 1 from `frmRoomMaintain` — Low
- **Location:** `frmRoomMaintain.tbrMenu_ButtonClick` Ctrl+T branch
- **Description:** Lands on the first row regardless of which room was being edited. Should auto-select that room's current type.
- **Mitigation:** Pass the current room's RoomType ID.

### `ListBooking` accepts but ignores `name` and `origin` parameters — Low
- **Location:** `frmFindCustomer.ListBooking`
- **Description:** Dead parameters left from a refactor.
- **Mitigation:** Remove.

### `frmRoomMaintain.cboLocation` auto-fill ranges don't match Dashboard ranges — Low
- **Location:** `frmRoomMaintain.ResetFields`
- **Description:** Form's case ranges (0–22, 23–39, 40–50, 51–61) don't match Dashboard's panel ranges (1–11, 12–33, 34–44, 45–55). New rooms get suggested with the wrong floor.
- **Mitigation:** Synchronize the ranges.

### `frmRoomMaintain` checkbox mutual-exclusion is recursive-unsafe — Low
- **Location:** `frmRoomMaintain.chkMaintenance_Click`, `chkHousekeeping_Click`
- **Description:** Setting one programmatically fires the other's click handler. Works in practice but fragile.
- **Mitigation:** Replace with a single `cboStatus` combo.

### `frmReportMaintain` save jumps selection to first row — Low
- **Location:** `frmReportMaintain.SaveReport`
- **Description:** After save, `LoadReports` resets selection to row 1, disorienting an editor working through several reports.
- **Mitigation:** Remember and reselect the saved ReportID.

### `frmFindCustomer`/`frmReport` no empty-state message — Low
- **Location:** `frmFindCustomer.ListCustomers`, `frmReport.LoadReports`
- **Description:** Zero-result searches/loads show empty ListViews with no "no rows" indicator.
- **Mitigation:** Add a "No results" hint or row.

### `frmPrint` keyboard polling can fire shortcuts twice per keypress — Low
- **Location:** `frmPrint.Timer1_Timer`
- **Description:** 100ms `GetKeyState` polling without debounce. A held key triggers multiple Print/Export actions.
- **Mitigation:** Track last-state and only fire on transition.

### `frmPrint` close shortcut is F4 not Esc — Low
- **Location:** `frmPrint.tbrButton`
- **Description:** Inconsistent with every other form (Esc). Possibly Esc was eaten by Crystal viewer.
- **Mitigation:** Pick a consistent close key; document the reason if Esc must remain unmapped.

### `frmPrint` taskbar offset hardcoded as `-680` — Low
- **Location:** `frmPrint.Form_Load`
- **Description:** Centering math assumes a bottom taskbar of fixed thickness; mispositions on borderless or top/side taskbars.
- **Mitigation:** Use `Screen.WorkingArea`; otherwise leave to default `StartUpPosition = 1`.

### `Picture4` (Level 4) overlaps summary strip area at certain DPIs — Low
- **Location:** `frmDashboard` form designer
- **Description:** Top coordinates overlap slightly with the summary strip depending on screen scaling.
- **Mitigation:** Verify and fix at common scaling factors.

### `lblBookingID` orphan on `frmDatabase` and no browse button — Low
- **Location:** `frmDatabase.frm`
- **Description:** Hidden `lblBookingID = "0"` left from copy-paste from `frmBooking`. Cosmetic. Also: user must type a path manually for the database file (no `CommonDialog` browse).
- **Mitigation:** Delete the orphan label; add a browse button.

### `frmRoomTypeMaintain.Form_Unload` calls `frmRoomMaintain.Form_Load` directly — Low
- **Location:** `frmRoomTypeMaintain.Form_Unload`
- **Description:** Invoking another form's event handler is unidiomatic and re-runs unrelated init (clock, button enable logic, ListRooms).
- **Mitigation:** Expose `frmRoomMaintain.PopulateRoomType` and call it directly.

---

## 8. Architectural Smells

### Global `gstrSQL` mutated by every database call — High
- **Location:** `Module/modGlobal.bas:gstrSQL`, mutated by every `SQL_*`/`SQLData_*`/`SQL_SET_*`/`SQLText` helper in `modCommon`
- **Description:** The defining architectural choice. Cannot build two queries in parallel from one thread; any function that builds a query and then calls another function that *also* builds one corrupts both. Migration required a parallel `OpenData`/`QueryDataSQL` to avoid the conflict.
- **Mitigation:** Move to per-call query strings (parameters or local builders). System-wide refactor.

### Single-threaded synchronous design freezes UI on long ops — Medium
- **Location:** `.vbp:MaxNumberOfThreads=1`; cross-cutting
- **Description:** A complex report query takes seconds — UI is unresponsive. No progress, no cancel.
- **Mitigation:** Move long ops to a background timer or a separate process; add progress UI.

### Hardcoded form lists in `frmDialog` cleanup duplicate logic — Medium
- **Location:** `frmDialog.cmdOK_Click`, `frmDialog.tmrClock_Timer`
- **Description:** Same eight `Unload` calls duplicated. New form additions require updating both blocks.
- **Mitigation:** Extract a `PerformLogout` sub.

### Comma management is the form code's responsibility (`blnEndComma` flag) — Medium
- **Location:** All `SQL_SET_*`/`SQLText` calls across the codebase
- **Description:** Caller passes `False` on the last call to suppress trailing comma. Forgetting → SQL syntax error. Conditional save blocks where the "last" column varies are particularly fragile.
- **Mitigation:** A modern pattern would build a list and emit commas automatically.

### Two ADO connections (`ACN`, `DAT`) maintained as globals — Low
- **Location:** `Module/modDatabase.bas`
- **Description:** Necessary architecturally for the migration's source/destination simultaneity, but means the codebase has two parallel connection lifecycles to manage.
- **Mitigation:** Acceptable if migration path remains; otherwise consolidate.

### `OpenSQL` and `OpenQuery` are functional duplicates — Low
- **Location:** `Module/modDatabase.bas:OpenSQL`, `OpenQuery`
- **Description:** Same forwards-only/optimistic/text-cmd open. Form callers pick arbitrarily.
- **Mitigation:** Remove one and update callers.

### `ExecuteSelectSQL` likely dead — Low
- **Location:** `Module/modDatabase.bas:ExecuteSelectSQL`
- **Description:** Caller-configurable cursor type; the form analyses don't show it being called.
- **Mitigation:** Audit callers; remove if dead.

### `UnlockDB`'s `Con` parameter is unused — Low
- **Location:** `Module/modDatabase.bas:UnlockDB`
- **Description:** Function takes `Con As ADODB.Connection` but uses `ACN` from `OpenDB` internally. The parameter is misleading.
- **Mitigation:** Remove parameter or use it.

### `UserAccessModule` does two queries when one would do — Low
- **Location:** `Module/modFunction.bas:UserAccessModule`
- **Description:** Reads ModuleAccess row (one query), then UserData row (second query) — could JOIN. And `gintUserGroup` is already in memory but ignored.
- **Mitigation:** Single-JOIN query, or use `gintUserGroup` directly when checking the current user.

### Three product names — Low
- **Location:** `Module/modGlobal.bas:COMPANY_PRODUCT_NAME`, every form's `lblBusinessName`, `Module/modDatabase.bas:CreateSampleData`, `.vbp` `Title`
- **Description:** "Hotel Booking System", "Room Booking System", "STAR HOTEL", and `App.ProductName` are different names for the same product, surfaced in different places.
- **Mitigation:** Single source of truth, ideally in the Company table.

### No connection pooling, no retry logic — Low
- **Location:** `Module/modDatabase.bas:OpenDB`
- **Description:** Every call creates a fresh `ADODB.Connection`. Harmless for single-user but no defensive layer for transient failures.
- **Mitigation:** Add retry on transient ADO errors.

### `pstrType` parameter in `UpdateField` only meaningfully checks "STRING" — Low
- **Location:** `Module/modFunction.bas:UpdateField`
- **Description:** "NUMBER", "BOOLEAN", "DATE" all go through the same non-quoted path. Documentation-as-code, not validation.
- **Mitigation:** Drop the parameter or add per-type validation.

### Module name vs filename mismatch (`modGlobalVariable` in `modGlobal.bas`) — Low
- **Location:** `Module/modGlobal.bas`, `.vbp` Module entry
- **Description:** In-source `Attribute VB_Name = "modGlobalVariable"` doesn't match filesystem name. Confuses search/navigation.
- **Mitigation:** Rename file or `Attribute`.

### Every form's idle-watchdog code is duplicated rather than centralized — Low
- **Location:** Every form's `tmrClock_Timer`
- **Description:** Same `intTick > gintUserIdle` logic copy-pasted across ~15 forms.
- **Mitigation:** Move into a shared sub, or use a single timer running on a hidden form.

### Color constants duplicated across forms — Low
- **Location:** Multiple forms (Dashboard and Booking each define `COL_*`)
- **Description:** Theme changes require touching multiple files.
- **Mitigation:** Centralize in a shared module.

### Print form opens its own ADO connection — Low
- **Location:** `frmPrint.OpenRDB`
- **Description:** Crystal Reports has its own connection — appropriate, but `gstrPassword` is a plaintext global available everywhere.
- **Mitigation:** Pass password ephemerally where possible.

### `SetDataSource rs, 3, 1` — unexplained magic numbers in Crystal call — Low
- **Location:** `frmPrint.Form_Load`
- **Description:** Numeric constants are CR-specific and undocumented in code.
- **Mitigation:** Add comments or named constants.

---

## 9. Build / Deployment Risks

### Crystal Reports 8.5 is end-of-life since ~2005 — High
- **Location:** `.vbp` references; `frmPrint`
- **Description:** CR 8.5 was released 2001, EOL ~2005. New installs need to source merge modules from archived installers; redistribution rights changed hands (Seagate → Crystal Decisions → BO → SAP) and are murky. Customers on new Windows hardware must hunt for the 20-year-old runtime.
- **Mitigation:** Plan migration to Crystal XI, BusinessObjects, or replace entirely (e.g., inline VB drawing, FastReports, or interop with a modern library).

### VB6 IDE itself is end-of-life — High
- **Location:** Entire codebase
- **Description:** VB6 IDE last updated 2008. Compiles only under VB6 IDE; new developers need an obsolete tool. The runtime ships with Windows, but builds require the IDE.
- **Mitigation:** Migrate to VB.NET / C#, accepting the substantial port cost.

### `Config.txt`/`Error.txt` write to `App.Path` would fail under Program Files install — Medium
- **Location:** `frmDatabase`, `Module/modTextFile.bas`
- **Description:** `asInvoker` privilege level is correct — but combined with `App.Path`-relative writes, an install in `C:\Program Files\` will silently fail to write `Config.txt` or `Error.txt`.
- **Mitigation:** Move user-writable files to `%APPDATA%\StarHotel\` or similar.

### Crystal viewer 8.0 + runtime 8.5 version mismatch — Low
- **Location:** `.vbp` Object/Reference declarations
- **Description:** Viewer at 8.0 and runtime at 8.5 — generally compatible but can produce subtle rendering bugs on certain installs.
- **Mitigation:** Match versions.

### ADO 2.8 is frozen forever — Low
- **Location:** `.vbp` references
- **Description:** Last released 2002. Works on every Windows since XP but has no further updates. Microsoft has hinted at deprecating Jet.
- **Mitigation:** Plan a future move to ADO.NET / SQLite / SQL Server LocalDB if the OS support window is critical.

### All compiler safety checks disabled in release config — Low
- **Location:** `.vbp` (`BoundsCheck=0, OverflowCheck=0, FlPointCheck=0, FDIVCheck=0, UnroundedFP=0`)
- **Description:** Standard release config. Latent bugs (integer overflow in `ConvInt`, array OOB) manifest as silent corruption rather than runtime errors.
- **Mitigation:** Keep checks on in debug builds.

### No DPI awareness in manifest — Low
- **Location:** `StarHotel.res` manifest
- **Description:** No `<dpiAware>True/PM</dpiAware>`. On 4K monitors the app is bitmap-upscaled and blurry.
- **Mitigation:** Add DPI awareness manifest entry; expect to re-test layout at non-100% scaling.

### No long-path-aware manifest entry — Low
- **Location:** `StarHotel.res` manifest
- **Description:** Bound by 260-character path limit. Matters if user picks a deeply-nested DB path.
- **Mitigation:** Add `<longPathAware>true</longPathAware>` and Windows registry opt-in.

### `AutoIncrementVer = 0` means version stays at 22 across rebuilds — Low
- **Location:** `.vbp`
- **Description:** Two binaries can both report 1.2.22.
- **Mitigation:** Enable auto-increment for build distinguishability.

### Manifest XML and `modMain.bas` "Markdown link" rendering is a doc artifact, not a code bug — Low
- **Location:** `docs/modules/MAIN-MODULE.md`, `docs/project/MANIFEST-FILE.md`
- **Description:** The actual on-disk `modMain.bas` and `.res`-derived manifest are clean. The corruption is in the docs only (a paste pipeline that wraps `Microsoft.Windows`/`frmSplash.Show` in Markdown links). The docs claim "this likely doesn't compile," which is false: the source file is clean.
- **Mitigation:** Fix the doc; do not touch the source.

---

## 10. Dead Code & Inconsistencies

### `DeleteRoomType`, `DeleteUser` bodies are corrupted copy-paste — Low
- **Location:** `frmRoomTypeMaintain.DeleteRoomType`, `frmUserMaintain.DeleteUser`
- **Description:** Reference controls/tables from other forms (`cboUserGroup`, `txtUserID`, `UserData`, `ListUsers`). Even if uncommented they wouldn't compile. The comment "Should not use this" implies intent.
- **Mitigation:** Remove entirely.

### Duplicated Read/Write helpers across `modCommon` and `modTextFile` — Low
- **Location:** `Module/modCommon.bas:WriteText/ReadText` vs `Module/modTextFile.bas:WriteFile/WriteTextFile/ReadFile/ReadTextFile`
- **Description:** Two near-identical implementations. `Print` vs `Write` differs slightly in formatting. `WriteFile` and `WriteTextFile` are functionally identical.
- **Mitigation:** Consolidate to one writer and one reader, full-path-only.

### Off-by-one `LineNo` in `ReadTextFile`/`ReadText`/`ReadFile` — Low
- **Location:** `Module/modTextFile.bas`, `modCommon.bas`
- **Description:** `For i = 0 To LineNo` runs `LineNo + 1` times; `LineNo = 1` returns line 2. Confusing API.
- **Mitigation:** Document or refactor to one-based indexing.

### `Log2File` defined but never called — Low
- **Location:** `Module/modTextFile.bas:Log2File`
- **Description:** Added 17/12/2014; the form analyses show it's unused.
- **Mitigation:** Adopt as the standard logger, or remove.

### Commented-out `DB_Version_Update` in `modFunction` — Low
- **Location:** `Module/modFunction.bas`
- **Description:** Migration code inlines the update; helper was abandoned.
- **Mitigation:** Remove.

### Commented-out `SetString`, `SetNumeric` in `modCommon` — Low
- **Location:** `Module/modCommon.bas`
- **Description:** Earlier-generation helpers superseded by `SetRecord`/`SetCheck`.
- **Mitigation:** Remove.

### Three module-level variables in `frmUserChangePassword` are unused — Low
- **Location:** `frmUserChangePassword.frm` (`strPassword`, `intAttempt`, `blnChange`)
- **Description:** Declared, never read or written. Dead state.
- **Mitigation:** Remove.

### `gblnReportSelect` and `gblnMaintainReport` flags commented in multiple places — Low
- **Location:** `frmReport`, `frmReportMaintain`
- **Description:** Cross-form coordination flags from a previous design.
- **Mitigation:** Remove.

### `PopulateGroupName` commented-out at top of `frmFindCustomer` — Low
- **Location:** `frmFindCustomer.frm`
- **Description:** Copy-paste leftover referencing `cboUserGroup` (doesn't exist on this form).
- **Mitigation:** Delete.

### `frmReport.PrintReport` has redundant second `OpenRS(gstrSQL)` — Low
- **Location:** `frmReport.PrintReport`
- **Description:** Recordset is unused; `frmPrint` opens its own connection and recordset.
- **Mitigation:** Remove.

### `frmDatabase` Cancel doesn't actually cancel the DB choice — Low
- **Location:** `frmDatabase.cmdCancel_Click`
- **Description:** Just dismisses the dialog and shows login, leading to login failure if path was broken. No "are you sure?" guard.
- **Mitigation:** Either `End` on cancel or warn.

### Inconsistent `OpenSQL` vs `OpenRS` choices across forms — Low
- **Location:** `frmAdmin`, `frmUserLogin`, `frmUserChangePassword` use `OpenSQL`; rest use `OpenRS`
- **Description:** Functional duplicates with different cursor types. The choice doesn't follow a clear pattern.
- **Mitigation:** Pick one default; use the other only when explicitly justified.

### Inconsistent `vbUnchecked` vs `0` literal for checkbox values — Low
- **Location:** `frmRoomTypeMaintain.ResetFields` uses `0`; rest use `vbUnchecked`
- **Description:** Equivalent but inconsistent.
- **Mitigation:** Standardize on `vbUnchecked`.

### `SQLText "0"` for LoginAttempts INSERT bypasses typed helpers — Low
- **Location:** `frmUserMaintain.SaveUser`
- **Description:** No `SQLData_Long` overload for raw integer literals; resorts to raw `SQLText`.
- **Mitigation:** Add an integer-literal helper.

### Form-designer dummy controls — Low
- **Location:** `frmDashboard` (`cmdUnit(0)` hidden template; `blnBlink(61)`/`strBackColor(61)` arrays sized 61 for ~55 rooms)
- **Description:** Cosmetic.
- **Mitigation:** Leave or trim.

### `IconForm = "frmBooking"` is a non-obvious choice — Low
- **Location:** `.vbp`
- **Description:** The `.exe` icon comes from `frmBooking` — typically you'd pick splash or dashboard.
- **Mitigation:** Change to `frmSplash` or `frmDashboard`.

### `Option Explicit` missing in `modMain` — Low
- **Location:** `Module/modMain.bas` (verified)
- **Description:** Every other module sets it. No undeclared variables exist, so harmless, but inconsistent.
- **Mitigation:** Add `Option Explicit`.

### `Dim dblVersion As String` in `DB_Version` then return as Double — Low
- **Location:** `Module/modFunction.bas:DB_Version`
- **Description:** Local typed as String, function declared `As Double`. Implicit conversions through `Val()` work but obscure.
- **Mitigation:** Type as Double directly.

### `ResFile32` only contains a manifest, no icon/version/string resources — Low
- **Location:** `StarHotel.res`
- **Description:** Unusual for VB6 — typical `.res` holds icons and version. The `.vbp`'s separate `VersionProductName` etc. and `IconForm` cover those, so no real impact.
- **Mitigation:** Cosmetic; consolidate if convenient.

### No comments in manifest XML for OS GUIDs — Low
- **Location:** Manifest XML
- **Description:** The `<supportedOS>` GUIDs have no inline comments naming the OS. Maintainers need an external lookup.
- **Mitigation:** Add `<!-- Vista -->` etc. comments.

### `PopulateGroupName`/`IsRoomTypeUsed` cleanup gaps — Low
- **Location:** `frmUserMaintain.PopulateGroupName` (catch missing `CloseDB`); `frmRoomTypeMaintain.IsRoomTypeUsed` (happy path missing `CloseRS`/`CloseDB`)
- **Description:** Cleanup relies on VB6 ref-counting rather than explicit close.
- **Mitigation:** Always close.

### Earlier resize-code variants commented out across forms — Low
- **Location:** `frmPrint.Form_Resize`, `frmUserChangePassword.Form_Resize`, others
- **Description:** Iterations preserved as comments rather than deleted.
- **Mitigation:** Delete or extract to git history.

### `ICC_*` constants and earlier `Sub Main` variant commented out in `modMain` — Low
- **Location:** `Module/modMain.bas` (top, verified)
- **Description:** Reference material left in source. Actual code uses `ICC_ALL_CLASSES` only. Also: `Modified On` header missing and `mstrModule` const not declared (every other module has both).
- **Mitigation:** Move reference material to documentation; add header date and `mstrModule` for consistency.

### `Form_KeyPress` debug code in `frmPrint` — Low
- **Location:** `frmPrint.frm` (top, commented)
- **Description:** `If KeyAscii = vbKeyP Then MsgBox "P pressed"` debug remnant.
- **Mitigation:** Delete.

---

## How to use this document

- **Triaging the rewrite.** Filter by Critical/High first — those define what *must* change. Anything Low is optional and can be designed-around rather than fixed.
- **Citing in PRs.** Each finding's location field is grep-friendly; reference it when fixing in the original or carrying the fix forward.
- **Updating.** When a finding is fixed (or proven false), strike it through with `~~~` rather than deleting — keeps the audit trail. When new findings surface, append them under the right category and bump the summary count.
- **Source-of-truth boundary.** Findings draw from `docs/` analyses; the highest-impact ones (lines marked "verified") were cross-checked against the source under `ORIGINAL-CODE/`. If a finding contradicts what you read in source, trust the source and update this file.
