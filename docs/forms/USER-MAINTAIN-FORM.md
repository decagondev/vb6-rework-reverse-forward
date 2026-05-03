# frmUserMaintain — User Account CRUD Editor

The F5-from-Dashboard form. This is the master/detail editor for user accounts: adding new users, changing their group/permissions/idle settings, resetting passwords, and unfreezing locked accounts. The last file in the project's CRUD-editor pattern, and the third (with `frmAdmin` and `frmUserChangePassword`) to deal with the password machinery.

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template: `imgLogo` left, `lblBusinessName` (24pt, right-aligned), `lblCopyright` underneath, `shpCopyright` rectangle.

**Toolbar `tbrMenu`** — 5 buttons matching the Maintain-form pattern:
- CLOSE (Esc), CLEAR (Ctrl+C), RESET (Ctrl+R), SAVE (Ctrl+S)
- DELETE (Ctrl+D) — *defined but `Enabled = False, Visible = False`*; intentionally disabled, same as `frmRoomTypeMaintain`

`fraLogin` on the right with the live clock and `lblUserID`.

**Main body — two-panel master/detail**

*Left — `Picture1`:*
- `lvUsers` ListView: hidden ID, "User ID", "User Name"
- Color-coded:
  - Pink (`#8080FF`) — inactive users
  - Green (`#76E600`) — Group 1 (Admin)
  - Gray (`#E0E0E0`) — everyone else

*Right — `Picture2`:* the editor pane, top to bottom:
- `cboUserGroup` (combo, populated from `UserGroup` table)
- `txtUserID *` (MaxLength 10, matches login form's cap)
- `txtName` (MaxLength 50)
- `txtIdle` — auto-logout seconds, MaxLength 4 (so 0–9999, validated to 0–3600). **`Enabled = False` by default** — only enabled when a non-default user group is selected
- `chkActive` — Yes/No (label colored pink, soft-delete convention)
- `chkReset` — "Reset login attempts (N)" — caption shows current count
- `chkChange` — "Change Password when login" — sets the forced-change flag
- `chkUpdatePassword` — gate for the three password fields below
- `txtPasswordOld`, `txtPasswordNew`, `txtPasswordConfirm` — all `Enabled = False` and `MaxLength = 8` (smaller than login form's 10!) initially, enabled by `chkUpdatePassword`. `PasswordChar = "ï¿½"` (the encoded character for password masking — different from the Wingdings trick used in login/admin/change-password forms; here it's a plain `*`-style mask using a literal box character)

Note the password field caps: 8 chars here vs 10 in the login form. So an admin creating a new user can only set an 8-character password, but the user could later change their own up to 10. Probably an oversight — the limits should match.

## Module-level state

```vb
intTick                  ' Idle-logout counter
COL_PINK = #8080FF
COL_GREEN = #76E600
COL_GRAY  = #E0E0E0
COL_DISABLED = #505050    ' Dark gray when password fields are disabled
COL_ENABLED  = #C0FFFF    ' Light cyan when enabled (visual cue)
```

The `COL_ENABLED`/`COL_DISABLED` pair is unique to this form — it visually toggles the password fields' background color when `chkUpdatePassword` flips them between writable and read-only.

## Logic

**`Form_Load`**
1. Sets the live clock and user ID label.
2. Calls `PopulateGroupName` to fill the user-group combo from `UserGroup` table.
3. Calls `ListUsers` to populate the master list.
4. The commented-out `lvUsers.ListItems(1).Selected = True` and `lvUsers_Click` show that auto-selecting the first user was deliberately removed — probably so the form opens with all fields blank, ready for a new entry.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog).

**`Form_Unload`** — re-shows `frmDashboard`.

**`Form_KeyDown`** — Esc closes, Ctrl+C clears, Ctrl+R resets, Ctrl+S saves. Ctrl+D for DELETE is commented out.

**`tbrMenu_ButtonClick`** — same four actions wired through the toolbar.

**`cboUserGroup_Click`** — when a non-zero group is selected, enables `txtIdle`. The intent is presumably "Group 0 = some kind of system role that doesn't need an idle timer," but in practice with two groups (Admin/Clerk), this just controls whether the idle field is editable. The `' Any point?` comment in the code suggests even the developer wasn't sure why this exists.

**`chkUpdatePassword_Click`** — toggles all three password fields and their colors based on the checkbox. Setting focus to `txtPasswordOld` when enabled. So password updates are explicitly opt-in — not every save touches the password.

**`PopulateGroupName`** — uses a `GetList` helper to fetch active groups:
```sql
SELECT GroupID, GroupName FROM UserGroup WHERE Active = TRUE ORDER BY SecurityLevel
```
Populates `cboUserGroup`. Each combo item's text is `GroupName`; its `ItemData` is the numeric `GroupID` — the standard VB6 pattern for "display name, store ID."

**`ListUsers`** — populates the master list with all users:
```sql
SELECT ID, UserID, UserName, UserGroup, Active FROM UserData
```
For each row: adds to ListView, color-codes by Active (pink) → UserGroup=1 (green) → else (gray). Calls `lvUsers_Click` at the end to auto-load the first row's data.

**`lvUsers_Click` / `lvUsers_KeyUp`** — when the user picks a row, calls `PopulateValues`. The keyboard wrapper delegates to the click handler.

**`PopulateValues(lngUserID)`** — reads the selected user via an INNER JOIN with `UserGroup` to get the group name in one query:
```sql
SELECT G.GroupName, D.ID, D.UserID, D.UserName, D.Active, D.Idle, D.LoginAttempts, D.ChangePassword
FROM UserData D INNER JOIN UserGroup G ON D.UserGroup = G.GroupID
WHERE D.ID = ?
```

Pushes each field into its corresponding control. The `chkReset.Caption` is dynamically updated to show the *current* login-attempt count: e.g., "Reset login attempts (2)" — so the admin can see at a glance whether the user is approaching the freeze threshold. Nice touch.

If no row found (which shouldn't happen since the ID came from the ListView) → "User not found" MsgBox.

**`SaveUser`** — substantial method handling both UPDATE and INSERT paths:

1. **Confirmation dialog**.
2. **User ID validation**: not empty.
3. **Idle-time validation**: 0–3600 seconds (1 hour max). Note: typed in seconds, so 3600 = 60 minutes. The cap matches the validation in `frmUserLogin` (`If gintUserIdle > 3600 Or gintUserIdle < 0`).
4. **Generate fresh salt** via `GenSalt(4)` — every save (whether or not the password is being changed!) computes a new salt. But the salt is only persisted when `chkUpdatePassword` is checked.
5. **Re-fetch the row** to determine UPDATE vs INSERT.
6. **For UPDATE path:**
   - Sets all the basic fields (UserGroup, UserID, UserName, Idle, Active).
   - **Password update is conditional** on `chkUpdatePassword` — only writes new UserPassword + Salt when explicitly opted in.
   - **Login attempts reset is conditional** on `chkReset` — clears the counter to 0.
   - **ChangePassword flag** is always written (toggling forced-change on/off).
7. **For INSERT path** (new user):
   - Inserts all the standard fields.
   - **Password is mandatory** for new users: `Encrypt(txtPasswordNew.Text, mstrSalt)` is unconditional — there's no "create a user without a password" path. But there's also **no validation that `txtPasswordNew` is non-empty** for the INSERT case! If the admin creates a new user without first checking `chkUpdatePassword` to type a password, the new account gets an empty password (or whatever's in the disabled field). Combined with `Encrypt("" & salt)`, the new user's hash would be predictable — anyone with the salt can compute it.
   - LoginAttempts is hardcoded to 0 (note the unusual `SQLText "0"` rather than `SQLData_Long 0` — bypassing the typed helper because there's no `SQLData_Integer` or `SQLData_Long` for a literal number? Looks like an inconsistency in the SQL helper API).
8. **`gintUserIdle = mintIdle`** at the end — *if the admin is editing their own record, this immediately updates the global idle setting in memory.* Subtle: if Admin updates Clerk's idle time, the Admin's own `gintUserIdle` gets overwritten with the Clerk's new value. That's a bug — should only apply if `mstrUserID = gstrUserID`.

**`ResetFields`** — blanks all editor fields. Sets `cboUserGroup.ListIndex = cboUserGroup.ListCount - 1` (last item, presumably the lowest-privilege role for a new user — Clerk if there are two groups). Resets all four checkboxes to unchecked. Clears the login-attempts caption text.

**`DeleteUser`** — fully commented out. Like `frmRoomTypeMaintain.DeleteRoomType`, the body is corrupted from copy-paste (it referenced this form's own controls but with stale logic). The message is the same: hard delete is unsafe; use Active = False instead.

## Notable points and quirks

- **Password length cap is 8 here, 10 in login.** A new password created by an admin in this form is capped at 8 characters; the login form accepts up to 10. So if an admin sets a user's password to "12345678" (8 chars max), the login is fine. But if a user later changes their own password (via `frmUserChangePassword` which has the 10-char cap), they could create a longer password than this form would allow. Inconsistency that probably hasn't bitten anyone in practice but should be reconciled.

- **Empty password on user creation is not blocked.** The INSERT path doesn't validate `txtPasswordNew.Text`. If the admin creates a user without checking `chkUpdatePassword`, the password hash is `Encrypt("", newSalt)`. That hash plus the publicly-stored salt is enough for anyone to compute the same hash and discover that the password is empty. The login form's check is `If Trim(txtPassword.Text) = ""` so an empty password from the user side would be rejected at login — but the hash is still computed and stored, and could be brute-forced. Worth fixing.

- **The `gintUserIdle = mintIdle` overwrite affects the wrong user.** When admin edits Clerk's record, the admin's session inherits Clerk's idle setting. Combined with `frmUserLogin`'s `gintUserIdle = 0` debug override (which makes the watchdog inert anyway), this bug currently has no observable effect — but if anyone fixes the login bug, this one will surface.

- **`chkReset` caption updates dynamically.** Showing the current LoginAttempts count next to the "Reset" checkbox is a nice UX touch — the admin can see if a user is being repeatedly locked out, suggesting a forgotten password.

- **`chkChange` is the freeze→change-password→login flow.** Setting this checkbox makes the user must change their password on next login — the same flag that `frmUserLogin` reads via `NeedChangePassword(gstrUserID)` and `frmUserChangePassword` clears via `SQL_SET_Boolean "ChangePassword", False`. Standard "first login forces password change" pattern.

- **Salt is regenerated on every save**, even if the password isn't being changed. The new salt is only persisted to the DB when `chkUpdatePassword` is checked, but `mstrSalt = GenSalt(4)` runs unconditionally. Wasted entropy, no harm done. (Though there's a subtle ordering trap: if you ever moved `SQL_SET_Text "Salt", mstrSalt` outside the `chkUpdatePassword` branch, you'd silently invalidate the password hash — because the existing hash was computed against the old salt.)

- **No audit trail for user changes.** Unlike `frmRoomMaintain` which copies to LogRoom on every UPDATE, `frmUserMaintain` just overwrites. Changes to who has access, who's frozen, who's an admin — none of this is logged. For an admin tool managing security, this is a gap. The comment "Recommend to log out and log in is required" alongside a TODO about `gintUserThemeID = UserThemeID(gstrUserID)` suggests the developer knew there were still rough edges.

- **`PopulateGroupName` doesn't `CloseDB` in its error handler.** Looking at the catch block, only `CloseRS` is called — the comment `'CloseDB` shows the intentional removal. So if the error happens after OpenDB but before CloseRS, the connection stays open. Combined with VB6's reference counting this probably gets cleaned up eventually, but it's inconsistent with the pattern elsewhere.

- **The `cboUserGroup_Click` enable-disable logic for txtIdle** — when index is 0 (the first group, presumably Admin since the SQL orders by SecurityLevel ascending), Idle is disabled. When >0, it's enabled. Translation: Admin users don't need idle timeouts; Clerks do. Reasonable policy. But combined with the `' Any point?` comment, you can sense the developer wasn't sure if this was the right design.

- **`InnerJoin UserGroup` in PopulateValues** is the only place in the codebase using this helper — most queries use `LEFT JOIN`. INNER is correct here because every user *must* have a group (foreign key constraint), but it would crash if the group was deleted. Soft-delete saves us here too.

- **No own-account protection.** An admin editing this form can disable their own Active flag, demote their own UserGroup, or freeze themselves out. The form doesn't refuse to do any of this. So it's possible to create a "no admins exist" scenario by deactivating the only admin account. Worth a guard like "you cannot deactivate your own account" or "at least one admin must remain active."

- **Soft-delete UX:** the gray/green/pink coloring in the master list lets the admin see at a glance: green = admins, gray = clerks, pink = disabled.

- **Tab order through the editor** is logical: UserGroup, UserID, UserName, Idle, Active, Reset, Change, UpdatePassword, then the three password fields. So the password fields are at the bottom because they're typically not touched.

- **The `chkUpdatePassword` opt-in pattern** is the right design — it prevents accidental password resets when the admin meant to update something else. Combined with the visual color cue (light cyan when enabled, dark when disabled), the form clearly communicates which fields are about to change.

- **`SQLText "0"` for LoginAttempts INSERT** is a minor SQL-helper API hole. The typed family has `SQLData_Long`, `SQLData_Integer`, `SQLData_Boolean`, etc., but emitting a literal `0` requires escaping into raw `SQLText`. Inconsistent.

- **No password-strength enforcement for new users.** The login form requires ≥4 chars only on change-password (`If Len(txtPasswordNew.Text) < 4`). This form has no such check. So an admin can create a user with a 1-character password if they remember to type one at all.

- **The `Encrypt(strPassword, salt)` wrapper is used here,** matching `frmUserChangePassword`. Both inline computations resolve the same way as `frmAdmin` and `frmUserLogin`'s direct `GoldFishEncode(password & salt)` calls — so cross-form, the password storage and verification scheme is consistent. As noted in the change-password analysis, if `Encrypt` ever drifts from raw concat-and-encode, login will silently fail — there's an integration contract worth verifying.

- **`Modified On : 27/12/2014`** matches the same date as several other forms — the same commit added the auto-logout timer everywhere. Consistent rollout.

- **`mintIdle = ConvInt(txtIdle.Text)` happens even when the field is disabled.** If the user picked a Group-0 user (admin) and disabled the idle field, then somehow saved, the `mintIdle` value would still come from whatever was in the textbox. In practice the textbox can't be edited while disabled so its value is the previously-loaded value or empty, which becomes 0.

- **Three password fields but only `txtPasswordNew` is actually used in the UPDATE path.** `txtPasswordOld` and `txtPasswordConfirm` exist for the visual flow but aren't validated against each other in the save. The admin can type "abc" in New, "xyz" in Confirm, and the user's password becomes "abc" — Confirm is decorative. (Compared to `frmUserChangePassword` where Confirm is actually checked.) This is a real bug — the admin could typo a password without noticing.

- **`tbrMenu.Buttons("DELETE").Enabled = True`** in the commented-out section of `PopulateValues` shows the original design: select a user → DELETE button enables. Now permanently disabled, the soft-delete via Active flag is the only path.

- **`gintUserThemeID` and `cboTheme`** appear in commented-out code throughout, alluding to a theme system that was planned but never shipped. Same fate as the bilingual report metadata in `frmReportMaintain`. The schema has the support; the UI doesn't surface it.
