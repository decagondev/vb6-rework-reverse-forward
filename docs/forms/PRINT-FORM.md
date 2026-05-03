# frmPrint — Crystal Reports Preview & Print

This is the universal report viewer used by every print path in the system — Booking receipts (Temporary/Official), Customer Transaction History, and presumably the Reports menu (F2). It's a generic shell around a Crystal Reports ActiveX viewer (`CRViewer1`).

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`.

**Header band** — same template as the rest: `imgLogo` left, `lblBusinessName` (24pt, **center-aligned** here rather than right — small inconsistency with the other forms), `lblCopyright` underneath, `shpCopyright` rectangle behind. The form title (`Me.Caption`) gets overwritten at load time with `gstrReportTitle`, so the form is contextually labeled by whatever report is being shown.

**Toolbar `tbrButton`** — 5 buttons (using its own `ImageList1`):
- CLOSE (F4) — note the unusual choice: F4 instead of Esc
- REFRESH (F5)
- SETUP (F6) — printer setup
- EXPORT (F7) — export to file (PDF, Excel, etc., via Crystal's built-in dialog)
- PRINT (F8)

`fraLogin` on the right with the live clock and `lblUserID` — same idle-watchdog as everywhere else.

**`Frame1`** below the toolbar contains the `CRViewer1` Crystal Reports ActiveX control — the viewer fills it. The frame and the viewer get resized together by `Form_Resize`.

**The CRViewer is heavily customized.** The form-designer config disables most of the viewer's built-in toolbar buttons (export, print, close, refresh, drill-down, animation, search-expert, select-expert, help) and instead re-implements them in this form's own toolbar. What's left enabled inside the viewer: navigation, zoom, stop, progress, search. So users get a clean preview area while the print/export/refresh actions stay under the form's control.

## Module-level state

```vb
CrReport       ' CRAXDRT.Report — the loaded Crystal Reports document
Con            ' Private ADODB.Connection (separate from the global one used elsewhere)
rs             ' ADODB.Recordset bound as the report's data source
F4_Pressed..F8_Pressed  ' Polled keyboard state flags
intTick        ' Idle-logout counter
```

The standalone `Con` and `rs` are **not** shared with the global `OpenDB`/`OpenRS` helpers used by the rest of the project. This form opens its own connection so the viewer keeps its data source live independently of whatever the rest of the app is doing.

## Logic

**Caller contract** — this form is created and shown by other forms (`frmBooking.PrintReceipt`, `frmFindCustomer.PrintHistory`, presumably `frmReport`). The caller sets three globals before showing:
- `gstrReportFileName` — e.g. `"Official Receipt.rpt"`
- `gstrReportTitle` — e.g. `"OFFICIAL RECEIPT"`
- `gstrSQL` — the query whose recordset will populate the report

**`Form_Load`** — does the heavy lifting:
1. Sets the title bar to `gstrReportTitle`, populates the clock and user ID.
2. Centers the form on screen manually (`(Screen.Width - Width) / 2`, top minus 680 — slight upward offset to compensate for the taskbar).
3. Permission gate: enables the EXPORT button only if the user has `MOD_REPORT_EXPORT` access. Print, refresh, setup, close are unconditional.
4. Loads the `.rpt` file from `App.Path\Report\`:
   ```
   Set CrReport = CrApplication.OpenReport(App.Path & "\Report\" & gstrReportFileName)
   ```
   `CrApplication` is presumably a globally-instantiated `CRAXDRT.Application`.
5. Opens its own ADO connection via `OpenRDB`, runs `gstrSQL` into `rs` as a static, pessimistically-locked recordset.
6. **Pushes the recordset into the report:** `CrReport.Database.SetDataSource rs, 3, 1`. This is the key trick — instead of letting Crystal Reports connect to the database itself, the VB code controls the SQL and hands the report the resulting rows. Means the `.rpt` file is just a layout template, and the SELECT statements live in the calling form (`frmBooking.PrintReceipt`, `frmFindCustomer.PrintHistory`).
7. Sets `ReportSource`, overrides the report's stored title with `gstrReportTitle`, calls `ViewReport` to render, sets zoom to `1` (which in CR's API is "fit to width" or similar, not 1%).

Wraps the whole thing in `Screen.MousePointer = vbHourglass` for the load delay.

**`OpenRDB`** — opens an ADO connection directly to the Access database:
```
Provider = Microsoft.Jet.OLEDB.4.0
Data Source = gstrDatabasePath
Jet OLEDB:Database Password = gstrPassword
```
So the database is password-protected at the Jet level (the password is held in a global), and this form has to know it to push data into the report. Worth flagging that `gstrPassword` exists as a plaintext global in memory.

**`CloseRDB`** — defensive close: checks `Con Is Nothing`, then `Con.State = adStateOpen` before closing, then sets to Nothing.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (idle watchdog → `frmDialog`).

**`Form_Resize`** — manually pins `Frame1` and `CRViewer1` to fill the available space below the toolbar:
- Top = 2040 (just below the toolbar)
- Height = Form.Height − 2500
- Width = Form.Width − 100

Several earlier iterations of this resize code are left commented out — looks like the developer iterated a few times to get the right spacing. The current version is the simplest.

**`Form_Unload`** — closes the recordset, closes the connection, releases the report. Notably, the `frmDashboard.Show` line is **commented out** — so closing this form doesn't auto-return to the Dashboard. Whatever form opened the print preview remains the active one underneath (which is correct, since `frmBooking.PrintReceipt` etc. show this form non-modally and stay open).

## The keyboard handling — an unusual approach

This is the most interesting part of the form. There's a 100ms `Timer1` that polls `GetKeyState` (declared via Win32 API at the top) for F4–F8 every tick:

```vb
F4_Pressed = (GetKeyState(vbKeyF4) < 0)
F5_Pressed = (GetKeyState(vbKeyF5) < 0)
...
If F4_Pressed Then Unload Me
ElseIf F5_Pressed And tbrButton.Buttons("REFRESH").Enabled Then CRViewer1.Refresh
...
```

Why? Because the Crystal Reports viewer is an ActiveX control that **swallows keyboard events** when it has focus. The form's own `Form_KeyDown` handler (currently commented out at the bottom) doesn't fire while the user is interacting with the report viewer. So a normal `KeyDown` shortcut won't work reliably.

The polling approach side-steps this: every 100ms the form asks Windows directly "is F4 currently pressed?" via the Win32 API, completely bypassing the focus/event-routing system. As long as the form is the foreground window, keystrokes are detected regardless of which child control has focus.

**Trade-off:** if the user holds a key down for longer than 100ms (which is normal), multiple ticks will fire. So pressing F8 once could trigger two or three Print dialogs. The code doesn't debounce. In practice the side effects of the actions (print dialog opens, focus changes) usually mask the issue, but it's not bulletproof.

The commented-out `Form_KeyDown` handler at the bottom uses Esc/F5/F6/F4/Ctrl+P with different mappings — clearly an earlier scheme that didn't work reliably for the same focus-stealing reason, hence the rewrite.

## `tbrButton_ButtonClick`

Standard toolbar dispatcher mapping the same five actions:
- CLOSE → `Unload Me`
- PRINT → `CRViewer1.PrintReport` (opens the system print dialog)
- EXPORT → `CrReport.Export` then `CRViewer1.Refresh` (triggers Crystal's export wizard, then redraws)
- REFRESH → `CRViewer1.Refresh` (re-renders without re-querying)
- SETUP → `CrReport.PrinterSetup(0)` (printer setup dialog, parent = 0 / desktop)

## Notable points and quirks

- **Reports are layout templates, data is pushed.** The `.rpt` files in `App.Path\Report\` define visual layout only; the SQL that fills them lives in the calling forms. This is a clean architecture choice — it means changing a query (e.g. adding a new filter) is a code change, not a Crystal Reports change. But it also means whoever edits an `.rpt` has to know which fields the calling code provides via the recordset, and the field names must match.

- **F4 to close instead of Esc.** Esc is the convention used in every other form. F4 here is unusual — F4 is the Windows convention for closing apps (Alt+F4) and might trip up users. The commented-out earlier handler used Esc; the polling-based replacement switched to F4. Possibly because Esc is consumed by the Crystal viewer for cancelling some internal operations.

- **Timer-based keyboard polling is rare.** This is a 2014-era VB6 workaround for an ActiveX limitation. It works, but ties the form to constantly running a 100ms timer which means battery and CPU never quite settle while this form is open. With modern hardware, irrelevant.

- **Two timers running simultaneously.** `Timer1` (100ms, keyboard polling) and `tmrClock` (1s, idle watchdog and clock). Each independent. The idle watchdog can fire `frmDialog` modally even while a Print dialog is up — which would be visually confusing but probably harmless.

- **`gstrPassword` global.** The Jet password is held in a global string in memory for the whole session. Anyone with debug access could read it. Standard for Access apps of this era.

- **Connection isolation.** This form opens its own `ADODB.Connection` rather than reusing the shared one. Smart for this use case — Crystal Reports keeps the recordset open for as long as the viewer needs it, and the rest of the app may want to issue queries in parallel without interfering.

- **`CrApplication` is implicit.** It's referenced but never declared in this file, so it lives in a standard module elsewhere — presumably a `Public CrApplication As New CRAXDRT.Application` in a `.bas` module that loads at startup.

- **Centering math has a magic offset.** `Top = (Screen.Height - Me.Height) / 2 - 680`. The `- 680` raises the form by ~12mm to compensate for the Windows taskbar. Hardcoded, so it'd misposition on screens without a taskbar at the bottom or with a thicker taskbar.

- **No "report has no data" handling.** If the SQL returns zero rows, the report renders empty. The calling forms (`frmBooking.PrintReceipt`, `frmFindCustomer.PrintHistory`) both have fallback queries that return the Company table with placeholder zeros to avoid this — that defensive logic lives in the callers, not here.

- **`Form_Resize` is wired up** but the form's `BorderStyle = 1` (Fixed Single) means users can't actually resize it. Either the resize handler is dead code from when it was resizable, or it fires on initial load to set the geometry once. Looking at the values it sets — `Frame1.Top = 2040`, etc. — those are just initial-positioning calculations. So the resize handler is effectively a one-shot layout pass triggered by the framework on form show.

- **`MOD_REPORT_EXPORT` is a separate permission from `MOD_REPORT_LIST`.** The export is gated independently, presumably so you can let staff view reports without letting them export raw data to Excel.

- **No print confirmation.** Hitting F8 or Print sends the report straight to the printer dialog. There's no "are you sure?" — sensible since the user has already chosen to print and the printer dialog itself is the confirmation.

- **The whole F-key set is unique to this form** (F4–F8). Every other form uses F2–F8 for a different purpose (Dashboard's CUSTOMER/ROOM/USER/ACCESS shortcuts). When this form opens on top of the Dashboard, those keys mean different things — fine because the Dashboard isn't receiving keystrokes while this is the active window, but worth noting for muscle-memory consistency.

- **Image alignment difference.** `lblBusinessName` here is `Alignment = 2 (Center)` while every other form uses `Alignment = 1 (Right Justify)`. Minor visual inconsistency.

- **The `Form_KeyPress` block at the top is commented-out scratch code** ("`If KeyAscii = vbKeyP Then MsgBox "P pressed"`") — debug remnants that should be cleaned up but don't affect anything.
