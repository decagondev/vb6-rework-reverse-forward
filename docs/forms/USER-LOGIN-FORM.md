# frmUserLogin — Authentication Form

The login screen. Reached after the splash bootstrap completes (or after `frmDatabase` finishes setup), and re-shown whenever a user logs out (closes the Dashboard).

## Layout

**Form shell**
- `BorderStyle = 1` (Fixed Single), 15375 × 9645 twips, no min/max, `KeyPreview = True`. Standard form chrome — unlike the maximized `frmUserChangePassword`, this one is a normal sized window.

**Header band** — same template as the rest: `imgLogo` left, `lblBusinessName` (24pt, right-aligned), and `shpCopyright` rectangle. Note: there's no `lblCopyright` directly under `lblBusinessName` here — the copyright label exists but is positioned at the bottom of the form (per the comments in `Form_Resize`, it was originally meant to follow the centered panel).

**Centered panel `fraLogin`** (a 9375 × 5220 PictureBox, manually centered by `Form_Resize`):
- `lblCompanyProduct` — set at runtime, but `Visible = False` (hidden by default — the company name shows in the header band instead)
- `lblProductName` (32.25pt bold, peach `#FF8080`) — large product banner
- `imgIcon` — small icon to the left of the input fields
- `txtUserID` (18pt, **orange** `#FF7929`) — MaxLength 10, MaxAttribut. UCase forced via `txtUserID_KeyPress`
- `txtPassword` (Wingdings 24pt, **green** `#76E600`, `PasswordChar = "n"`) — MaxLength 10, same masking trick as `frmAdmin` and `frmUserChangePassword`
- `cmdOK` "OK (Enter)" + `cmdCancel` "Cancel (Esc)" + `cmdOption` "Option" (no shortcut, but tooltip says "Change database file")
- Three small labels at the bottom: `lblVersionApp` (set at runtime to `App.Major.Minor.Revision`), `lblVersionDB` (from `DB_Version`), `lblDemo`
- Two `Shape2`/`Shape3` rectangles behind the textboxes for the dark-grey faux-border look

The `lblDemo` says "User ID: admin / Password: admin" — explicit credentials shown in the UI for the demo mode. Anyone with a screenshot of the login screen has the credentials. This is consistent with the "demo database" first-launch story but means the system really is open by default until someone changes the admin password.

## Module-level state

```vb
strPassword     ' Used as a working buffer in cmdOK_Click
intAttempt      ' Local-session attempt counter, never persisted
```

## Logic

**`Form_Load`**
1. Sets the title bar (`lblBusinessName`) and the inner panel headers from `COMPANY_PRODUCT_NAME` and `App.ProductName`.
2. Populates the App version label from `App.Major.Minor.Revision`.
3. Populates the DB version label from `DB_Version` (which reads `Company.DatabaseVersion` from the database).
4. Sets the demo hint text. Note the commented-out original version that included an expiry date — "valid until DateAdd('D', 364, '01 Jan 2021')" — pointing back to the same Jan 1, 2021 anchor used in `frmSplash`'s commented-out kill-switch.

**`Form_Resize`** — re-centers the inner `fraLogin` panel:
```vb
fraLogin.Left = (Me.Width - fraLogin.Width) \ 2
fraLogin.Top = (Me.Height - fraLogin.Height) \ 2
```
True center this time, unlike the offset-up arithmetic in `frmUserChangePassword`. The commented-out lines below show that the copyright label and a separator line used to follow the panel's position dynamically — that's been disabled, leaving the copyright pinned to the form designer's static position.

**`cmdOption_Click`** — opens the database picker:
1. Unloads the login form.
2. Reads the current path/filename from `Config.txt`.
3. Pre-fills `frmDatabase` with them and shows it modally.

This is the user-facing way to change which database the app connects to, after first run. Useful if there are multiple data files (different companies, test/production).

**`cmdCancel_Click`** — `End`. Same heavy-handed termination as the forced-password-change cancel path. There's no graceful logout from the login screen — Cancel = quit the app.

**`cmdOK_Click`** — the heart of the form. The login flow:

1. **Empty-field guards** for User ID and Password. (Note: the commented-out testing block at the top would auto-fill `CLERK`/`1` if both fields were empty — convenient for development, removed before release.)

2. **Lookup query** via the typed SQL helper:
   ```sql
   SELECT UserGroup, UserID, UserName, UserPassword, Salt, Idle, Active, LoginAttempts, ChangePassword
   FROM UserData
   WHERE UserID = <CheckInput(typed user)>
   ```

3. **Unknown user** → "User ID not found!" message and refocus. *Different behavior from `frmAdmin`, which silently failed.* The login form leaks more information about which credential was wrong; the admin form was more cautious. Inconsistency.

4. **Stash session globals** for use throughout the app:
   - `gintUserGroup` — drives `UserAccessModule` checks everywhere
   - `gstrUserID` — shown on every form's user label
   - `gstrUserName` — full name
   - `gstrUserPassword`, `gstrUserSalt` — kept for password-change verification
   - `gintUserIdle` — drives the idle watchdog timer in every form
   - **But then** immediately overwrites: `gintUserIdle = 0` with a comment "Test". This is a bug — the per-user idle setting from the database is read, then immediately discarded. Auto-logout is effectively disabled because of this leftover debug line. Worth flagging prominently.
   - The 0/3600 sanity-check on `gintUserIdle` runs *after* the override, so it's also moot.

5. **Frozen account check.** If `Active = False`, shows "Your User ID has been frozen" and `End`s the application. Note: `End` rather than `Exit Sub`. Login failure → app quits. No way to try a different account in the same session.

6. **Attempt-cap check.** If `LoginAttempts > 2`, the same logic as `frmAdmin`: tries to "freeze" the account by `UpdateField ... Active = TRUE` (still the same bug — sets Active to TRUE, not FALSE), then shows the frozen message and `End`s. So the freeze never actually fires from this branch; what does fire is the wrong-password loop below.

7. **Password verification.** Concatenates `password + salt`, runs through `GoldFishEncode`, compares to stored hash. Same as `frmAdmin` and `frmUserChangePassword`.

8. **On success:**
   - Resets `LoginAttempts` to 0 in the database.
   - Calls `NeedChangePassword(gstrUserID)`. If True (forced password change is pending), unloads self and shows `frmUserChangePassword`. **Important:** sets `gblnUserChangePassword = False` here — meaning when the change-password form runs Cancel, it'll go to the False branch and `End` the app. So forced password change at login is enforced strictly.
   - Otherwise, checks `MOD_DASHBOARD` permission. If granted, unloads self and shows Dashboard. If not granted, shows "Your access has been disabled!" and unloads. Without anything else taking over, the app effectively exits via running out of forms.

9. **On wrong password:**
   - If `gintUserGroup > 1` (non-superuser), increments `LoginAttempts` in the DB via `UpdateField` with the special `"INCREMENT_ONE"` flag (presumably the helper interprets this as `LoginAttempts = LoginAttempts + 1`).
   - Local `intAttempt` increments. After 3 wrong tries in this session, "Too many attempts. Application quit!" and `End`. So wrong passwords don't just lock the account — they also kill the app, forcing the user to relaunch.
   - "Invalid Password, please try again!" message, refocus password, `SendKeys "{Home}+{End}"` to select the existing text for retyping. Same UX touch as `frmAdmin`.

10. **Error handler** has an unusual addition: `If Err.Number = 0 Then Exit Sub` — this catches cases where the error handler was triggered by a `Resume`-style flow but no actual error occurred. Defensive but odd to need it.

**`txtUserID_KeyPress`** — uppercases every keystroke at typing time. So the User ID is always entered in caps, regardless of caps-lock state.

**`txtUserID_Validate`** — uppercases the whole text on focus-leave. Belt and braces.

**`lblCopyright_Click`** — auto-fills "ADMIN" / "admin" into the fields. **This is a developer convenience that shipped to production.** Anyone clicking the copyright text at the bottom of the screen gets the admin credentials filled in automatically. Combined with the on-screen "User ID: admin / Password: admin" hint in `lblDemo`, the system advertises its own default credentials in two places. For a demo install this is intentional; for a production deployment it's a security hole.

**Commented-out `imgIcon_DblClick`** — would have launched `frmUserChangePassword` directly from the login screen. Probably removed because it'd let users change their password without authenticating first — which is exactly the wrong direction.

## Notable points and quirks

- **`gintUserIdle = 0` debug override is the most important bug.** Right after the database loads the per-user idle setting, the line `gintUserIdle = 0 '10 ' Test` blanks it out. The comment `' Test` makes clear this was for a development session, and the previous `'10` value (10-second auto-logout for testing) was the intermediate. The production path was meant to be `gintUserIdle = ConvInt(rst!Idle)` with the override removed. As written, **the entire idle-watchdog feature** that's wired into every Form's `tmrClock_Timer` across the codebase is **inert** — `intTick > gintUserIdle` is always `intTick > 0`, but the gate is `If gintUserIdle > 0 Then` so the body never runs. So the auto-logout never fires anywhere in the system. This matches the pattern across all the forms where the watchdog code is present but appears never to trigger in practice.

- **`lblCopyright_Click` cheat code.** Click the copyright text → admin credentials auto-fill. Combined with the on-screen hint, the system has zero meaningful login security on a fresh install.

- **The frozen-account "freeze" bug repeats.** Same as `frmAdmin`: `UpdateField ... Active = TRUE` is meant to freeze but actually sets Active to True. So the auto-freeze on too many DB-persisted attempts can't actually freeze the account. The local `intAttempt` counter does cap a session at 3 wrong attempts and `End`s the app, but a determined attacker just relaunches the application — `intAttempt` resets to 0 each launch.

- **`End` is used as the response to many failure modes.** Frozen account → `End`. Too many attempts → `End`. Cancel button → `End`. Forced password change cancelled → `End` (via the chain into `frmUserChangePassword.cmdCancel_Click`). So the application's stance is: "if anything goes wrong with auth, quit." This is consistent and uncompromising but means there's no way to recover from a typo without restarting the app.

- **User ID is forced UPPERCASE** at the keyboard level. So `admin` and `ADMIN` are the same login. This affects how `WHERE UserID = ?` matches — the database stores them in uppercase too. Removes a class of "user typed the wrong case" bugs.

- **Wrong-password handling is gated on `UserGroup > 1`.** Group 1 is Admin (per `frmModuleAccess` analysis). So Admin users have unlimited login attempts — the LoginAttempts counter doesn't increment for them, and the local intAttempt also doesn't trigger the "too many attempts" lockout. Practical reasoning: the admin shouldn't be able to lock themselves out. But it also means a brute-force attempt against the admin password has no rate limit at the application level.

- **`OpenSQL` (not `OpenRS`) again.** Third form to use the alternate helper. Probably a synonym, but the inconsistent naming across the codebase remains a smell.

- **No "remember me" or session persistence.** Every restart requires a fresh login. Appropriate for a multi-user POS-style system.

- **The `imgIcon_DblClick` change-password shortcut is commented out.** Smart removal — a "change password without logging in" path would be a serious vulnerability.

- **`MaxLength = 10`** on both User ID and Password — same constrained limits as everywhere else. A 10-character password ceiling combined with the `GoldFishEncode` hash is the system's effective security strength. For an internal hotel app it's adequate; for an internet-facing system it would be insufficient.

- **`gblnUserChangePassword = False` is set just before showing the change-password form** in the forced-change path. This is the flag that tells `frmUserChangePassword.cmdCancel_Click` whether Cancel should go gracefully back to the Dashboard or `End` the app. Setting it to False here means the forced-change Cancel will `End` — strict enforcement.

- **Dashboard is shown without re-fetching the user's permissions.** The `UserAccessModule(MOD_DASHBOARD)` check happens once at login. If an admin disabled this user's Dashboard access while they were on the login screen (unlikely but possible in a multi-user system), the user wouldn't find out until the Dashboard tried to use those permissions for individual buttons. The login allows the Dashboard to load even if the underlying `UserData.Active` flag was just toggled — race window.

- **Login attempts counter persistence across sessions** is the right design — without it, an attacker could just relaunch the app to reset the counter (which they can do anyway thanks to the local `intAttempt` reset and the freeze-account bug).

- **`CheckInput(txtUserID.Text)` is used for SQL safety** in the WHERE clause and in subsequent UPDATE statements. The password is never concatenated into SQL — it's hashed before comparison — so injection through the password field isn't possible. The User ID field is sanitized.

- **The login form is the entry point for the database picker** via the Option button, which is nice — power users can swap databases without going through `frmSplash`. But it's the only labeled affordance for "advanced" actions on the form.

- **Tab order through the panel:** UserID (0), Password (1), OK (2), Cancel (3). Logical. Option is at TabIndex 13 — at the back, which makes sense (you don't normally tab to it).

- **`Modified On : 19/05/2018` and `13/07/2019`** in the header — two iterations after the original 2014 version. Hard to tell from the code what changed in each — probably the maximize/centering layout (2018) and the Option button to choose database (2019, matching the `frmDatabase` "Added On : 13/07/2019" note).

- **The whole codebase converges on one or two security primitives** — `GoldFishEncode + Salt`, `MaxLength = 10`, `UserGroup > 1` for rate-limit gating. They appear in `frmAdmin`, `frmUserLogin`, and `frmUserChangePassword` consistently. So patching the security model is, in principle, a focused operation across three files.
