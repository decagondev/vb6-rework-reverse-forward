# frmUserChangePassword — Change Own Password

A small dialog that lets the currently-logged-in user change their own password. Reachable from `frmDashboard` via F8 (Security button), and also forced on first login or after an admin reset (when the `ChangePassword` flag is true on the user's record).

## Layout

**Form shell**
- `BorderStyle = 0` (None) — frameless, like the splash
- `WindowState = 2` (Maximized) — fills the whole screen
- `ShowInTaskbar = False`, no min/max/control box
- `StartUpPosition = 1` (Center Owner)
- Background `#303030`, with a centered `fraLogin` PictureBox holding all the actual UI

**Centered panel `fraLogin`** (10575 wide, 5055 tall, dark gray):
- `lblCompanyProduct` (top-left, 18pt) — set at runtime from `COMPANY_PRODUCT_NAME`
- `lblProductName` (32.25pt bold, peach `#FF8080`) — large product banner
- `imgIcon` — a small icon to the left of the input fields (1095 × 1095, padlock-style)
- Three labeled password fields, each on top of a `Shape2/3/4` rectangle (the dark grey textbox-border simulation):
  - `txtPasswordOld` (Wingdings 24pt, **orange** `#FF7929` foreground) — Alt+O
  - `txtPasswordNew` (Wingdings 24pt, **green** `#76E600` foreground) — Alt+N
  - `txtPasswordConfirm` (Wingdings 24pt, **red** `#4417FF` foreground) — Alt+C
- All three fields use `PasswordChar = "n"` in Wingdings — same trick as `frmAdmin`, where `n` in Wingdings renders as a filled black square. MaxLength 10 on all three.
- `cmdOK` "OK (Enter)" + `cmdCancel` "Cancel (Esc)" at the bottom-right
- A horizontal line and `lblCopyright` at the very bottom

The three different field colors are a thoughtful touch: orange = "what you have now" (caution), green = "what you want" (positive), red = "confirm = critical" (urgency). All masked, but the color cue helps users glance and know which row they're typing in.

## Module-level state

```vb
strPassword     ' Declared but unused
intAttempt      ' Declared but unused
blnChange       ' Declared but unused
```

All three module-level variables are declared but never read or written. Dead state.

## Logic

**`Form_Load`** — sets the company-product label and the product name from `App.ProductName`. The commented-out `Me.BackColor = &HFFC000` lines suggest an earlier orange-background design that was switched to dark gray.

**`Form_Resize`** — re-centers the inner `fraLogin` panel whenever the form is resized. Since the form opens maximized, this fires once on load to position the centered panel correctly:

```vb
fraLogin.Left = (Me.Width - fraLogin.Width) \ 2
fraLogin.Top = (Me.Height - fraLogin.Height) \ 2 - fraLogin.Height
```

The `- fraLogin.Height` at the end pulls the panel up by its own height — so it's not actually centered, it's positioned in the upper-half of the screen. Looking at the geometry: the panel is 5055 high; on a 1080-pixel screen `Me.Height` is around 16000 twips, so the calculation gives `(16000-5055)/2 - 5055 = 472` — which is near the top of the screen. So the panel sits in the top portion, not the middle. That's a deliberate UX choice (probably to keep it above the keyboard on touchscreen kiosks) but the offset arithmetic is unusual.

The commented-out lines under that show an earlier resize design where Line1 and lblCopyright were repositioned relative to `Me.Height`. Now those positions are fixed to fraLogin coords.

**`Form_Unload`** — checks `UserAccessModule(MOD_DASHBOARD)`. If the user has Dashboard access, shows it; otherwise shows a "Your access has been disabled!" warning. Either way, unloads. The fallback "frmBooking.Show" is commented out — there's no longer a meaningful destination for users without Dashboard access; they just see a warning and the app effectively closes (since no other form is visible).

**`cmdOK_Click`** — the validation chain:

1. Old password not empty.
2. New password not empty.
3. **Verify old password** via `CheckPassword`. If wrong, error and exit.
4. New password length ≥ 4.
5. New password matches confirmation.
6. If all pass, calls `UpdatePassword(newPassword)`.
7. After save, navigates to Dashboard if the user has access, or shows "access disabled" warning otherwise.

The order is fine, but worth noting: the new-password length check happens *after* the old-password verification. So if you mistype your old password, you find out before you find out your new one's too short. Reverse order would let you fix the easier error first; this order is more secure (don't waste time on form validation if auth fails).

**`cmdCancel_Click`** — branches on `gblnUserChangePassword`:
- If True (user voluntarily opened this from F8/Security): unloads, shows Dashboard, clears the flag.
- If False: calls `End` — terminates the entire application immediately.

The False branch handles the **forced password change** case. When a user logs in with `ChangePassword = True` on their record (set by an admin or by first-time setup), the login form opens this dialog without setting the flag. Hitting Cancel on a forced password change kills the app — there's no way around it. You either change your password or quit. That's the security model: you cannot use the system without first resetting your password if forced.

The `End` statement is heavy-handed — it doesn't run cleanup, doesn't close database connections cleanly. For a forced-quit scenario it's adequate but not graceful.

**`CheckPassword(strPassword)`** — verifies the typed old password against the stored hash:

```sql
SELECT UserPassword, Salt FROM UserData WHERE UserID = gstrUserID
```

Concatenates `strPassword + Salt`, runs through `GoldFishEncode`, compares to stored `UserPassword`. Same scheme as `frmAdmin` — same hand-rolled hash with same caveats about modern security standards.

**`UpdatePassword(strPassword)`** — writes the new password:

1. **Generates a new 4-character salt** via `GenSalt(4)` — every password change gets a fresh salt. Good practice.
2. UPDATEs the user's row, setting `UserPassword = Encrypt(strPassword, mstrSalt)` (note: `Encrypt` here, not `GoldFishEncode` directly — presumably `Encrypt` is a wrapper that does the salt-concat-then-GoldFish dance).
3. Sets `Salt = newSalt`.
4. Sets `ChangePassword = False` — clearing the forced-change flag.
5. WHERE UserID = current user.
6. Confirmation MsgBox "Password is updated!".

## Notable points and quirks

- **The forced-password-change vs voluntary distinction.** `gblnUserChangePassword` is the global flag that distinguishes the two flows. Set to True by `frmDashboard` when the F8 button is pressed (you can see this in the Dashboard code: `gblnUserChangePassword = True; frmUserChangePassword.Show`). Not set by `frmUserLogin` when the login flow detects `ChangePassword = True`. So:
  - **Voluntary change**: Cancel → return to Dashboard, no harm done.
  - **Forced change**: Cancel → application terminates. Your only options are to change the password or quit.

  The `End` statement on the forced-change Cancel path is the security kill switch.

- **`MaxLength = 10`** on all three fields — same restrictive cap as `frmAdmin`. A 10-character password is weak by 2026 standards. Combined with `GoldFishEncode`, this is a clear "don't rely on this for real security" situation.

- **Three module-level variables (`strPassword`, `intAttempt`, `blnChange`) are declared but unused.** Dead state, probably copied from another form (likely `frmUserLogin` or `frmAdmin`) and never cleaned up.

- **No "passwords don't match" message until you try to submit.** A more responsive UI would compare the New and Confirm fields on focus-leave and color them red on mismatch. Here, you find out only when you click OK. Standard for the era.

- **Wingdings password masking** — same custom-glyph trick as `frmAdmin`. `n` in Wingdings is a filled square. The three fields each render their characters in different colors (orange/green/red), so even though the characters are identical glyphs, you can visually tell which field you're in.

- **No password-strength indicator.** The only check is "≥ 4 characters." No requirement for digits, capitals, special characters. Someone could change their password to "aaaa" and the form is fine with it.

- **Old password is verified before any new-password validation.** Subtle security choice — even revealing whether your typed new password was syntactically valid before you've authenticated would leak information. The current order is more conservative.

- **Maximize-on-show pattern.** The form fills the entire screen even though the actual UI is a small panel in the middle. This is deliberate — when the form is shown over the Dashboard or after login, it fully obscures everything else, preventing the user from seeing or interacting with anything until they've completed (or cancelled) the password change. A modal full-screen dialog. Combined with `End` on forced-cancel, this enforces the "you must change your password to proceed" workflow.

- **The off-center positioning** (panel sits in upper third of screen) is unusual. Probably so it's clearly visible above any taskbar/notification area, even on smaller screens. The arithmetic in `Form_Resize` is mildly suspect but produces a consistent layout.

- **No idle watchdog.** Unlike most other forms, this one doesn't set up a `tmrClock` or check `gintUserIdle`. Reasonable — the user is mid-authentication-task, you don't want them auto-logged-out partway through changing their password. They'll either finish or cancel.

- **Same `OpenSQL` (not `OpenRS`) inconsistency** as `frmAdmin`. Two different helpers in the codebase that look like they do similar things. Worth checking if these are aliases or distinct functions.

- **`Encrypt` vs `GoldFishEncode`.** `CheckPassword` calls `GoldFishEncode(strPassword & rst!Salt)` directly. `UpdatePassword` calls `Encrypt(strPassword, mstrSalt)`. Presumably `Encrypt` is a wrapper that does `GoldFishEncode(password & salt)`, making both code paths produce comparable hashes. But the asymmetric naming is confusing — one path uses the high-level wrapper, the other inlines. If `Encrypt` ever does anything different from raw concat-then-GoldFish (e.g., adds iteration count, or KDF rounds), the verify path will silently fail and no one will be able to log in after a password change. Worth verifying these are equivalent.

- **The "Your access has been disabled!" branch** is interesting — it could fire after a successful password change if the user's Dashboard permission was revoked while they were logged in. In practice this would only happen if an admin deactivated them between login and password change, but the code defensively handles it.

- **`lblCompanyProduct` and `lblProductName`** are populated from `COMPANY_PRODUCT_NAME` and `App.ProductName` — same pattern as `frmSplash`. The dialog is brand-aware, matching the rest of the system.

- **`ChangePassword = False` is explicitly set** on update. This is the flag that drives the forced-change flow. Once cleared, the user won't be forced into this dialog on next login (until an admin sets it back to True).

- **`GenSalt(4)`** generates a 4-character salt. That's short by modern cryptographic standards (typical recommendation is 16+ bytes) and means there are only ~2.8 million possible salts in the printable-ASCII range. For a small internal hotel app this is acceptable; for anything larger it would be a concern.

- **Tab order through the panel:** old password (0), new (1), confirm (2), OK button (3), Cancel (4). Logical flow, with Enter triggering OK by default and Esc triggering Cancel. The Wingdings labels have `&` accelerators (Alt+O, Alt+N, Alt+C) for keyboard shortcuts to each field.

- **`MaxLength = 10` enforced both at typing time and at field level.** No truncation logic in code — the textbox itself enforces it. So someone with an 11-character muscle-memory password literally can't type the 11th character.

- **Comment "Modified On : 19/05/2018"** at the top with no description. So this form has been touched since the original 2014 codebase, but the change log is empty. Quick examination suggests the maximize/centered-panel layout might have been the 2018 change — probably to support larger monitors or kiosk displays where a small dialog floating in the middle would feel lost.
