# frmAdmin — Temporary Admin Login Dialog

A small modal credential prompt designed to step up privileges mid-session — specifically intended to authorize a Void operation when the regular logged-in user doesn't have that permission. The header comment dates it back to 04/01/2014.

## Layout

**Form shell**
- `BorderStyle = 3` (Fixed Dialog), 8175 × 2790 twips, dark gray (`#303030`), centered on owner.
- Hidden from the taskbar (`ShowInTaskbar = False`) — it's a child dialog, not a top-level window.
- No min/max buttons.

**Visual elements**
- `imgIcon` on the left — 1095 × 1095 stretched icon (presumably a padlock or admin symbol).
- Two label/textbox pairs on the right, each sitting on a `Shape2`/`Shape3` `#505050` rectangle that simulates a textbox border (the textboxes themselves have `BorderStyle = 0`):
  - `lblLabels(0)` "&User ID :" → `txtUserID` (large 18pt, orange `#FF7929` foreground, MaxLength 10)
  - `lblLabels(1)` "&Password :" → `txtPassword` (Wingdings 24pt with `PasswordChar = "n"` → renders typed characters as a Wingdings glyph for password masking; foreground is the green `#76E600`, MaxLength 10)
- The accelerator keys `Alt+U` and `Alt+P` jump focus to the corresponding textbox.

**Buttons**
- `cmdOK` "OK (Enter)", `Default = True`
- `cmdCancel` "Cancel (Esc)", `Cancel = True`

**`lblBookingID`** — hidden, caption "0", positioned behind other controls. Unlike the `lblBookingID` on `frmDatabase` (which was dead/copy-pasted), this one is actually load-bearing: `frmDashboard` and `frmBooking` were originally going to set it to the booking that needs to be voided before showing this form modally. The commented-out block in `cmdOK_Click` shows the intent — read the booking ID from this hidden label and call `frmBooking.VoidBooking` on it.

## Logic

**`cmdCancel_Click`** — just `Unload Me`. No fallback navigation, since this dialog is shown modally on top of another form.

**`cmdOK_Click`** — the substantive part. Validates credentials against `UserData`:

1. **Empty-field guards** for both User ID and Password, with focus jumping to the missing one.

2. **Lookup** via the typed SQL helper layer:
   ```
   SELECT UserID, UserGroup, UserPassword, Salt, Active, LoginAttempts
   FROM UserData
   WHERE UserID = <CheckInput(typed user)>
   ```
   `CheckInput` is presumably an SQL-injection sanitizer for the user-typed string.

3. **Silent fail on unknown user.** If `RecordCount = 0` the form just closes the recordset and exits without telling the user — no "user not found" message. This is a deliberate security-through-ambiguity choice (don't reveal which half of the credentials was wrong) but the implementation is inconsistent because the wrong-password branch *does* show "Invalid Password, please try again", which leaks the same info anyway.

4. **Frozen account check.** If `Active = False`, shows "Your User ID has been frozen. Please contact Super User." and exits.

5. **Attempt-cap check.** If `LoginAttempts > 2`, shows the same frozen message and exits — **but this branch has a bug**: it tries to "freeze" the account by calling `UpdateField "UserData", "Active", "BOOLEAN", "TRUE", ...` which sets Active to TRUE, not FALSE. So the guard only displays the warning; the account never actually gets locked here. (In practice the next attempt-increment branch — see step 7 — is what eventually freezes it via `UserGroup > 1` logic, but the behavior is fragile.)

6. **Password verification** — interesting custom implementation:
   - Concatenates the typed password with the per-user `Salt` from the row.
   - Runs the result through `GoldFishEncode` (a custom hash/encode function, name suggests it's a homemade or low-strength algorithm — *not* a vetted KDF like bcrypt/PBKDF2/Argon2).
   - Compares the encoded result to the stored `UserPassword`.

7. **On success**: resets `LoginAttempts` to 0 via `UpdateField`, closes the recordset and DB, **and exits the sub.** That's it — `Unload Me` is *not* called, and the commented-out block that was supposed to actually perform the Void is disabled. So as the code currently stands, a successful admin login closes nothing, performs no action, and just dismisses control back to the caller. The caller (which would be `frmBooking` if it were invoking this modally) has no way to know whether authentication succeeded or the user cancelled.

8. **On wrong password**: increments `LoginAttempts` only if `UserGroup > 1` (the meaning of UserGroup here isn't visible in this file, but presumably > 1 means a non-super-user). Then if the local `intAttempt` counter exceeds 2, shows "Too many attempts." and unloads the form.

9. **Final fallthrough**: shows "Invalid Password, please try again!", refocuses the password box, and runs `SendKeys "{Home}+{End}"` — `Home` then `Shift+End` — to select all the existing text in the field so the user can immediately retype.

**Error handling** — standard `On Error GoTo CheckErr` template using the database-based `LogErrorDB` (since this dialog is only used after the DB is up and running).

## Notable points and quirks

- **Wingdings password masking is unusual.** Setting `PasswordChar = "n"` in a Wingdings font is a clever trick — `n` in Wingdings renders as a filled black square (■), giving a custom-looking masked-character display. Functionally equivalent to a normal `*`-mask, but cosmetically different.

- **The form is essentially a no-op stub.** The block that actually performs the privileged action is fully commented out:
  ```
  ' If UserAccessModule(MOD_BOOKING_VOID, strUser) Then
  '     frmBooking.VoidBooking CLng(lblBookingID.Caption)
  '     Unload Me
  ' Else
  '     MsgBox "Your have no access to this function!", ...
  ' End If
  ```
  Combined with the comments in `frmBooking` saying "VOID is not implemented since not a request", this confirms the whole Void feature is dormant. The form authenticates, then does nothing with the result.

- **Two parallel attempt counters.** There's `intAttempt` (a local variable that always starts at 0 every time `cmdOK_Click` runs) and the persisted `LoginAttempts` column on `UserData`. The local one resets per click, so the "Too many attempts" warning would never trigger from a single dialog session unless the user types the wrong password three times in one open instance. The persisted counter survives across attempts — this is what the admin-account-freeze logic relies on.

- **`UpdateField` is being called with what look like SQL fragments.** `"Loginattempts + 1"` is passed as the value for an attempted column update — meaning `UpdateField` must be building `SET Loginattempts = Loginattempts + 1` rather than `SET Loginattempts = 'Loginattempts + 1'`. That's a subtle helper-API contract worth noting.

- **`CheckInput` only sanitizes the User ID, not the Password.** The password is never put into a SQL string here — it's compared in VB after hashing — so this is fine, but worth flagging that the asymmetry is correct rather than a bug.

- **`MaxLength = 10`** on both User ID and Password is restrictive by modern standards. A 10-character password ceiling significantly weakens whatever protection `GoldFishEncode` provides.

- **`GoldFishEncode` is a red flag.** A hand-rolled or proprietary encoder (rather than a standard PBKDF2/bcrypt/scrypt/Argon2) plus salt-after-concatenation, plus a 10-char password limit, plus the system being VB6 era — this is not a credential storage scheme to rely on for anything sensitive. For an internal hotel-room booking app circa 2014 it's typical of the period; for a system being deployed today it's a significant weakness worth flagging to the maintainer.

- **`SendKeys "{Home}+{End}"`** — micro-UX touch to make retry easier. `SendKeys` is generally fragile (it sends to whatever has focus, which can race with focus changes) but here it follows immediately after `txtPassword.SetFocus` so it's reliable.

- **Inconsistent behavior on invalid user vs invalid password.** Unknown user → silent close. Wrong password → loud message. The intent was probably "fail silently to avoid revealing whether the user exists" but the second branch undermines that.

- **`OpenSQL` vs `OpenRS` inconsistency.** The other forms (Dashboard, Booking) use `OpenRS(gstrSQL)`. This form uses `OpenSQL(gstrSQL)`. Either there are two helpers with similar behavior, or this is a mismatch to investigate — a function name typo here would make the form crash at runtime, so presumably both helpers exist.

- **The hidden `lblBookingID` has a real purpose** that's currently disabled. The contract is: caller sets `frmAdmin.lblBookingID.Caption = <bookingID>` then `frmAdmin.Show vbModal`. You can see exactly that pattern in the *commented-out* lines in `frmBooking`:
  ```
  'frmAdmin.lblBookingID.Caption = lngBookingID
  'frmAdmin.Show vbModal
  ```
  So the entire Void → Admin step-up flow is wired but commented out at both ends. Activating it would mean uncommenting four blocks across two files and verifying the freeze-account logic in step 5 is fixed.
