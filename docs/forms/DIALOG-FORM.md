# frmDialog — The Idle-Logout Warning Dialog

The sixteenth and final form in the project — the one we knew existed (referenced in every form's `tmrClock_Timer` handler) but never had source code for. Now we do. Versioned 1.2.22, **Added On : 27/12/2014** — the same date as the auto-logout timer additions across the other forms (`frmUserMaintain` and several others all share this date in their headers). So this form was created as part of the December 2014 "add idle-logout watchdog" feature rollout.

The form is small (~80 lines including the form designer block), has a single job, and reveals one important architectural detail about the application's logout flow that wasn't visible from the calling sites alone.

## Layout

**Form shell**
- `BorderStyle = 3` (Fixed Dialog) — the dedicated dialog style with a thinner border, no min/max buttons
- `ClientHeight = 3195`, `ClientWidth = 6030` — small, ~6000×3200 twips (about 400×220 pixels)
- `MaxButton = False`, `MinButton = False`, `ShowInTaskbar = False`
- `StartUpPosition = 1` (CenterOwner) — appears centered on whichever form invoked it
- Background `#303030` — matches the rest of the dark theme
- Caption: "Log out"

**Controls (top to bottom):**
- `lblMessage` (12pt bold, **peach `#FF8080`**) — "Alert: System will be log out automatically after"
- `lblSecondsLeft` (18pt bold, **red `#0000FF` — though the hex `&H000000FF&` is actually red in BGR notation**) — large countdown display, starts at "10 second(s)"
- `lblInstruction` (9.75pt) — "Click OK to log out now or Cancel to ignore."
- `cmdOK` "OK (Enter)" — `Default = True`
- `cmdCancel` "Cancel (Esc)" — `Cancel = True`
- `tmrClock` (Interval 1000 = 1 second) — drives the countdown

The grammar in the message ("System will be log out automatically") is awkward — should be "logged out" — but it's been there since 2014 and presumably no one's complained or noticed.

The peach color of `lblMessage` is intentionally alarming (matches the `lblUserID` color across other forms but used here for warning rather than identification). The bright red on `lblSecondsLeft` reinforces the urgency — countdown timers are conventional to color-warn.

## Module-level state

```vb
Dim intSecondsLeft As Integer
```

Just one variable: the countdown counter, initialized to 10 in `Form_Load` and decremented each timer tick.

## Logic

**`Form_Load`** — Initializes `intSecondsLeft = 10` and sets the label to "10 second(s)". Simple init.

**`tmrClock_Timer`** — Fires every second. Decrements `intSecondsLeft`, updates the label, and when the counter hits zero, performs the same "close everything" sequence as `cmdOK_Click`. So the user has 10 seconds to click Cancel; otherwise the auto-logout proceeds.

**`cmdCancel_Click`** — Just `Unload Me`. The dialog goes away. The calling form's idle counter (`intTick` in `tmrClock_Timer`) doesn't get reset by this — but every form's timer handler that invokes `frmDialog.Show vbModal` is structured so that *after* the modal call returns, normal form interaction continues. The user's continued interaction will eventually reset the timer state (any keystroke or click resets `intTick = 0` in well-behaved form designs, though we'd need to verify this in each form's code).

**`cmdOK_Click`** — Performs the immediate-logout sequence:

```vb
Unload Me
CloseAllPrintForms
Unload frmReportMaintain
Unload frmReport
Unload frmRoomTypeMaintain
Unload frmRoomMaintain
Unload frmUserMaintain
Unload frmModuleAccess
Unload frmBooking
Unload frmDashboard
'    CloseAllForms
```

Eight individual `Unload` calls in a specific order, plus a helper that closes all print forms first, plus a commented-out call to a more general cleanup function.

**The sequence is hand-curated, not algorithmic.** The form unloads exactly these forms by name, in this order:
- `frmReportMaintain` and `frmReport` — the report editor and viewer
- `frmRoomTypeMaintain` and `frmRoomMaintain` — the room admin pair
- `frmUserMaintain`, `frmModuleAccess` — the user/permission admin
- `frmBooking` — the booking editor
- `frmDashboard` — the central hub, last

The order matters because the parent-child relationships (`frmReportMaintain` is reachable from `frmReport`, `frmRoomTypeMaintain` is reachable from `frmRoomMaintain`) are unloaded child-first. Unloading the parent first while the child holds an active reference could cause issues; child-first is safer.

**What happens AFTER all this unloading?** Looking at the form-flow analysis: `frmDashboard.Form_Unload` re-shows `frmUserLogin`. So the unload chain ends with `frmDashboard` being unloaded, which triggers the login form to appear. The user is back at the authentication screen — exactly what "log out" means in this system.

## The two helper functions

**`CloseAllPrintForms`** — Iterates the global `Forms` collection and unloads any form named "frmPrint". Why a special function for this one form?

The answer: **`frmPrint` can have multiple instances.** Looking back at the print form's analysis, it's the universal Crystal Reports viewer used by Booking, FindCustomer, and Report. If a clerk opens one report, then opens another without closing the first, you'd have two `frmPrint` instances active simultaneously. The other forms (`frmDashboard`, `frmBooking`, etc.) are single-instance — they have `VB_PredeclaredId = True` and the codebase always uses the named instance. But Print could be multi-instance, hence the iteration.

**`CloseAllForms`** — Commented out (`'    CloseAllForms` after the unload list). The function is *defined* in the module but never called. It would have been a more generic version of the cleanup:

```vb
For Each frm In Forms
    If Not (frm.Name = "frmDialog" Or frm.Name = "frmDashboard") Then
        Unload frm
        Set frm = Nothing
    End If
Next
Unload frmDashboard
Set frmDashboard = Nothing
```

It would have closed **every** form that wasn't itself or the Dashboard, then closed the Dashboard last. More general than the hardcoded list, but never wired in. Probably an attempted refactor that was left as dead code in case someone wanted to enable it later.

The hardcoded list approach is **fragile but explicit**. If a future developer adds a new admin form (say, `frmAuditLog`), they'd have to remember to add `Unload frmAuditLog` to *both* the OK click handler and the timer handler in this dialog, otherwise the new form would be left orphaned after a logout. The commented-out `CloseAllForms` would have made this future-proof — but it isn't called.

## Notable points and quirks

- **The form fills in the missing piece of the idle-watchdog story.** Every analyzed form had `tmrClock_Timer` checking `If intTick > gintUserIdle Then ... frmDialog.Show vbModal`. The dialog is what they show. Now we know what happens next: 10-second countdown, then a coordinated multi-form unload that returns the user to the login screen.

- **The watchdog is currently inert because of the `gintUserIdle = 0` debug override in `frmUserLogin`.** So this dialog **never actually appears** in the running application as configured. The form exists, is wired up, has all the logic in place — but the gate that would invoke it (`If gintUserIdle > 0 Then ... If intTick > gintUserIdle`) never fires because `gintUserIdle` is always 0 after login. The dialog is dormant. Fix the login bug and the entire idle-logout feature comes alive instantly.

- **The 10-second grace period is hardcoded** in `Form_Load`. If a customer wanted "30 seconds to react" they'd need source code access. A more configurable design would read this from `Company` table or from `gintUserIdle / 6` or some derivation of the idle-out value. Hardcoded 10s is reasonable for most cases but not flexible.

- **The `cmdOK_Click` and `tmrClock_Timer` logout blocks are duplicated.** Same eight `Unload` calls, same `CloseAllPrintForms` helper call, same commented-out `CloseAllForms`. If a developer added a new form that needs cleanup, they have to remember to update **both** blocks. Easy to miss one. A simple refactor would extract a shared `Sub PerformLogout` and call it from both — that refactor wasn't done.

- **The countdown updates "10 second(s)" → "9 second(s)" → "8 second(s)" → ... → "1 second(s)" → logout at 0.** Note the parenthetical "(s)" — the form doesn't bother with singular/plural ("1 second" vs "2 seconds"), it just always says "second(s)". Lazy but practical. In a more polished UI you'd write `If intSecondsLeft = 1 Then "second" Else "seconds"`.

- **The countdown displays the same value briefly twice — once at Form_Load (set to "10 second(s)") and once at the first timer tick (set to "9 second(s)" after decrementing).** Wait — actually no, Form_Load sets it to "10", then the first timer tick decrements to 9 and displays "9". So the user sees 10 → 9 → 8 → ... → 1 → logout. Total elapsed time is 10 seconds. Correct timing.

- **`tmrClock` is enabled by default.** The Timer's `Enabled` property isn't shown in the .frm but the default is True, so the countdown starts immediately when the form is shown. No "click here to start" — the clock is running from the moment the dialog appears.

- **`vbModal` from the calling site means the calling form's UI is blocked while this dialog is showing.** So while the 10-second countdown runs, the user can't interact with the form they were on. They can only click OK or Cancel. This is appropriate — an idle-warning that you could ignore by just clicking elsewhere wouldn't be much of a warning.

- **The dialog's caption is "Log out"** but the message says "Alert: System will be log out automatically after". The caption is misleadingly direct — the dialog isn't *about* logging out, it's *warning* about an impending logout. A better caption would be "Auto-Logout Warning" or "Are You Still There?". Minor UX issue, but consistent with the somewhat awkward messaging throughout.

- **The dialog doesn't provide a way to *extend* the timeout** — only to cancel it (i.e., reject the auto-logout for now) or accept it. So a user actively working but who happens to step away for 5 minutes (triggering the warning) and clicks Cancel just to continue working still has the same 5-minute idle timer hanging over them. Click after click after click of Cancel each time the warning pops up. A "extend by 30 minutes" button would be friendlier — but again, hardcoded behavior.

- **The dialog never actually returns the user to login directly.** Looking at the unload sequence: the dialog unloads itself, then unloads child forms in order, ending with `frmDashboard`. When `frmDashboard.Form_Unload` runs, it calls `frmUserLogin.Show`. So the path back to login goes *through* the Dashboard's unload handler, not directly from this dialog. The dialog trusts the chain: "if I unload Dashboard, the system's normal unload behavior will route to login." This works because the Dashboard is the only form whose `Form_Unload` shows the login screen — every other form's `Form_Unload` shows the Dashboard. Walking up the chain naturally lands at login.

- **Session globals are NOT cleared by the logout.** After this dialog runs and the user is back at the login screen, `gstrUserID`, `gintUserGroup`, `gstrUserPassword`, `gstrUserSalt`, etc. are all still populated with the previous user's values. The next login overwrites them, but in the gap between logout and re-login (which could be long if no one logs back in), the previous user's session state persists in memory. For a single-user kiosk this is mostly harmless, but a more security-conscious design would explicitly clear these globals after logout — particularly `gstrUserPassword` and `gstrUserSalt`. Not done here.

- **`frmAdmin`, `frmDatabase`, `frmFindCustomer`, `frmSplash`, `frmUserChangePassword`, `frmUserLogin`, and `frmDialog` itself are NOT in the unload list.** Looking at why each is excluded:
  - `frmDialog` — that's `Me`, already unloaded at the start
  - `frmAdmin` — modal step-up dialog, doesn't usually persist
  - `frmDatabase` — only shown from splash or login Option button, not during normal use
  - `frmFindCustomer` — **conspicuously absent** despite being a major form launched from Dashboard. This is a bug — `frmFindCustomer` could be open at idle-logout time and would NOT be closed by this sequence. Worth flagging.
  - `frmSplash`, `frmUserChangePassword`, `frmUserLogin` — these are bootstrap/auth flow forms that wouldn't normally be open during idle-logout

  The `frmFindCustomer` omission is the most likely real bug — that form is reachable from the Dashboard via F3 and the user could plausibly be on it when the idle timeout fires. The dialog's hardcoded list misses it.

- **The `Set frm = Nothing` calls inside `For Each frm In Forms`** are technically incorrect — VB6's Forms collection doesn't update when you nil the iteration variable, and setting a local variable to Nothing doesn't release the form anyway (Unload does that). The Nothing-setting is harmless but cargo-cult code. Same pattern in `CloseAllForms` (which is commented out).

- **No error handling.** Unlike most forms in the codebase which use `On Error GoTo CheckErr` patterns, this dialog has no error handlers anywhere. If `Unload frmReport` fails (because frmReport has an OnUnload handler that itself errors), the whole logout sequence would propagate the error up. In practice this doesn't happen because none of those forms' Unload handlers do risky work, but it's a small completeness gap.

- **The button order is OK on the left, Cancel on the right** (TabIndex 0 = OK, TabIndex 1 = Cancel; positions cmdOK Left=1320, cmdCancel Left=3240). On a Windows-conventional dialog, the *destructive/affirmative* action is usually on the right and Cancel on the left. Here, the "destructive" action (log out, lose your work) is OK on the left, default-button. The Default = True on cmdOK means hitting Enter immediately logs out. Combined with the auto-countdown, this is moderately user-hostile — a user with their finger on the Enter key (perhaps from typing in the form they were just on) could trigger immediate logout without intent. A safer design would have Cancel as Default.

- **The form has no idle watchdog of its own.** No `tmrClock` checking `gintUserIdle`, no recursive call to `frmDialog.Show vbModal` while the dialog is showing. Sensible — the dialog *is* the idle-watchdog response, you don't want it warning the user about being idle while displaying an idle warning.

- **The form is in `vbModal`** which means the calling form's `tmrClock_Timer` is paused until the dialog returns. So the calling form's `intTick` doesn't keep incrementing during the 10-second grace period. Good — you don't want the countdown timer racing with the calling form's idle timer.

- **The icon comes from `frmDialog.frx`** (offset 0). Since we don't have the .frx, we don't know what it looks like, but given the rest of the codebase uses fairly standard Windows-application icons, it's probably a generic info or warning icon.

- **`Added On : 27/12/2014`** matches the modification date in `frmUserMaintain`'s header ("Use timer to count after an interval then auto log out") and similar comments in other forms. So this form was added as part of a December 2014 push to add the idle-logout feature across the entire app — a coordinated multi-form change.

- **The form represents the design pattern's weakest link.** The watchdog feature requires:
  1. `gintUserIdle` set per-user (works, except for the debug override)
  2. Every form's `tmrClock_Timer` correctly increment `intTick` and reset on user activity (mostly works)
  3. Every form to invoke `frmDialog.Show vbModal` when threshold exceeded (works)
  4. **This dialog** to coordinate the cleanup of all open forms (works for the hardcoded list, misses `frmFindCustomer`)
  5. The unload chain to land at login (works via Dashboard's unload handler)

  The whole feature is currently dormant due to bug #1, but if fixed, the missing `frmFindCustomer` in step 4 would manifest as "user is on FindCustomer when idle fires, FindCustomer stays open after logout, login screen appears underneath, confusion ensues."

- **The form is the canonical example of the codebase's design philosophy: explicit and concrete over abstract and reusable.** A `CloseAllForms` helper exists right here in the same module but isn't used; instead, eight individual `Unload` calls hardcode the cleanup order. This is consistent with the rest of the codebase — concrete naming over generic iteration, hardcoded constants over configurability, copy-pasted logic over extracted functions. It works, but it ages poorly because every new form requires updates in multiple places.

So the "forgotten dialog form" turns out to be a small but architecturally significant component — it's the missing keystone of the idle-logout feature, and analyzing it reveals both how the feature was *designed* to work and the specific gaps (hardcoded form list missing FindCustomer, no session-globals clearing, no extend option, default button is the destructive action) that would surface immediately if the inert feature were ever activated by fixing the login-time `gintUserIdle = 0` override.