# frmFindCustomer — Customer Search & Booking History

This is the F3-from-Dashboard form that lets staff look up past guests by name/passport/origin/contact within a date range, see all of that customer's bookings, and print a transaction history report. Created 08/01/2015 per the header.

## Layout (top to bottom)

**Form shell**
- Same dark theme (`#303030`), fixed single border, 15375 × 9645 twips, no min/max, `KeyPreview = True`. Identical chrome to the Dashboard and Booking forms.

**Header band** — same pattern as the other main forms: `imgLogo` left, `lblBusinessName` (24pt) right-aligned, `lblCopyright` underneath, sitting on a `shpCopyright` rectangle.

**Toolbar `tbrMenu`** with 4 buttons (smaller than Dashboard/Booking — this is a focused tool):
- CLOSE (Esc), CLEAR (Ctrl+C), FIND (Ctrl+F), PRINT (Ctrl+P)
- `fraLogin` on the right carries the live clock and `lblUserID` — same idle-watchdog pattern.

**Main body — three panels in an L-shape**

*Top-left panel `Picture1`* — search criteria:
- `txtGuestName` + Label "Name"
- `txtGuestPassport` + Label "Passport / IC No" + `cboSQL1` (AND/OR connector to the row above)
- `txtGuestOrigin` + Label "Country / Origin" + `cboSQL2` (AND/OR connector)
- `txtGuestContact` + Label "Contact No" + `cboSQL3` (AND/OR connector)
- `dtpBookingDateFrom` + Label "Booking Date From"
- `dtpBookingDateTo` + Label "Booking Date To"

The three `cboSQL*` combos are an interesting choice — they let the user pick AND/OR between successive criteria rather than forcing strict-AND. Each combo controls the connector for *its own* row relative to whatever rows are already in the query.

*Top-right panel `Picture2`* — `lvCustomers` ListView:
- Columns: `ID` (zero-width, hidden — actually shows Passport text), Booking No, Date, Room Type, Amount.
- Wait — that's misleading. Looking closer, the column headers are "Passport / IC No", "Guest Name", "Country / Origin", "Contact No". The Picture3 ListView underneath has the Booking columns. So the master/detail layout is: top-right shows distinct customers found, bottom shows that customer's bookings.

*Bottom panel `Picture3`* — `lvBookings` ListView (full-width):
- Columns: `ID` (zero-width, hidden), Booking No, Date, Room Type, Amount.

So the visual flow is: enter criteria on the left → click Find → matching distinct customers appear top-right → click a customer → their bookings appear in the wide list at the bottom → optionally Print.

## Module-level state

```vb
intTick           ' Idle-logout counter
strGuestPassport  ' Snapshot of selected customer's passport for printing
strGuestContact   ' Snapshot of contact for printing
```

## Logic

**`Form_Load`** — sets the clock, user ID label, calls `ResetFields`. Wrapped in error handling.

**`Form_Activate` / `Deactivate`** — start/stop `tmrClock` (same idle-watchdog template as the other forms — fires `frmDialog` after `gintUserIdle` seconds).

**`Form_Unload`** — re-shows `frmDashboard`. Closing this form returns control there.

**`Form_KeyDown`** — Esc closes, Ctrl+C clears, Ctrl+F finds, Ctrl+P prints (only if there are results and a customer is selected). The Ctrl+P branch grabs `lvCustomers.SelectedItem.Text` (the passport) and the trimmed contact field.

**Each search textbox responds to Enter** — `txtGuestName_KeyDown`, `txtGuestPassport_KeyDown`, `txtGuestOrigin_KeyDown`, `txtGuestContact_KeyDown` all call `ListCustomers` on Enter. Nice UX touch.

**`ResetFields`** — clears all four text fields, sets all three AND/OR combos to index 0, and resets the date range to `YearDay1(Now)` through `MonthDay30(Now)`. Those are app helpers that presumably mean "Jan 1 of current year" and "30th of the current month" — so the default search window is "year-to-date through end of this month."

**`ListCustomers`** — the main search routine and the most code-dense part of the form. It builds a SELECT against `Booking` table dynamically:

1. Starts with `SELECT DISTINCT GuestPassport, GuestName, GuestOrigin, GuestContact FROM Booking`.
2. Walks the four search fields in order. The first non-empty field gets a `WHERE ... LIKE '%value%'` clause and sets `blnAND = True`. Each subsequent non-empty field reads its corresponding AND/OR combo and appends `AND/OR field LIKE '%value%'`.
3. Always appends a date-range filter (`BETWEEN #from# AND #to#`) and `Active = TRUE AND Temp = FALSE`. The `WHERE` vs `AND` keyword for the date filter depends on whether any text criterion was set — if not, the date range becomes the leading WHERE.
4. Iterates the recordset, populates the customer ListView with passport (as the main `Text`) plus name, origin, contact as sub-items.

**Two important gotchas in this code:**

- **`SELECT DISTINCT` on four fields** means that two visits by the same passport with even slightly different recorded contact numbers (e.g., different formats, typos, or the contact updated between bookings) will show up as two separate rows in the customer list. This is a real-world issue with denormalized data.

- **Operator precedence trap.** The dynamic clause builder generates raw `AND`/`OR` between criteria with no parentheses. Combined with the trailing `AND (BookingDate BETWEEN #...# AND #...#) AND Active = TRUE AND Temp = FALSE`, an `OR` between two text criteria can interact unexpectedly with the date filter. For example: `WHERE Name LIKE '%a%' OR Passport LIKE '%b%' AND BookingDate BETWEEN ...` evaluates as `Name LIKE '%a%' OR (Passport LIKE '%b%' AND BookingDate BETWEEN ...)` — so any "Name" match gets through *regardless of date*. There's a bug here whenever the user picks OR between criteria.

**`lvCustomers_Click`** — when a customer row is clicked (or focused via keyboard via `lvCustomers_KeyUp` which just delegates back), reads the four sub-item texts and calls `ListBooking`.

**`ListBooking(passport, name, origin, contact)`** — populates the bottom ListView with that customer's bookings:

1. `SELECT ID, Format(ID, '100000') AS BookingNo, CreatedDate, RoomType, Payment FROM Booking`.
2. Filter is `WHERE GuestPassport = <selected passport>`. The name and origin parameters are received but **not used in the query** — they're commented out.
3. If the contact field is non-empty, appends `AND/OR GuestContact LIKE '%contact%'` based on `cboSQL3`. (Note: `cboSQL3` controls connector to GuestContact specifically, which is consistent with how it's used in `ListCustomers`.)
4. Always appends the date-range filter and `Active = TRUE AND Temp = FALSE`.
5. The hidden first column actually receives the booking ID, with formatted BookingNo in column 2, CreatedDate in 3, RoomType in 4, and currency-formatted Payment in 5.

The fact that the function takes name+origin parameters but doesn't use them is a code smell — either the filter was simplified later, or it's a stub for future expansion.

**`PrintHistory(passport, [contact])`** — builds a Crystal Report query:
1. Sets `gstrReportFileName = "Customer Transaction History.rpt"` and `gstrReportTitle`.
2. Builds a SQL string by concatenation (same pattern as `frmBooking.PrintReceipt` — *not* using the typed SQL helper layer here either):
   - Joins Company and Booking.
   - Selects company info + booking details + computed Total = Payment - Refund.
   - Filters by `B.Active = TRUE AND B.GuestPassport = '<passport>'`.
3. **Bug here:** the contact filter block calls `SQL_AND_LIKE_Text "GuestContact", strGuestContact` and `SQL_OR_LIKE_Text` — but those helpers write to `gstrSQL` directly, not concatenate to the local string. Since `gstrSQL` was just set by string concatenation above, calling these helpers will likely either overwrite it or produce garbage. This branch looks broken.
4. Appends `AND B.BookingDate BETWEEN #from# AND #to#`.
5. If the query returns no rows (`QueryHasData` returns False), substitutes a Company-only fallback query with placeholder zeros so the report still renders headers without crashing.
6. Sets `rep.Caption = "CUSTOMER TRANSACTION HISTORY"` and shows the report form.

**`tbrMenu_ButtonClick`** — standard toolbar dispatcher. The PRINT branch checks `lvCustomers.ListItems.Count > 0` and uses `lvCustomers.SelectedItem.Text` as passport. There's an implicit assumption that `SelectedItem` is non-Nothing — if the user clicked Find but never selected a row, `SelectedItem` could be Nothing and this would crash. In practice clicking a row to display bookings makes selection automatic.

## Notable points and quirks

- **Master-detail UI pattern.** Search criteria → customer list → booking list → print. Clean and discoverable.

- **AND/OR per field.** Letting the user combine criteria with their choice of operator is unusual — most search UIs hard-code AND. It's flexible but the precedence bug noted above means OR can produce surprising results. Adding parentheses around the user-selected criteria block would fix it.

- **Three search modes get implicitly supported by the LIKE wildcards.** `'%value%'` means partial matches anywhere — so typing "John" finds "Johnson" and "Jonathan." For passport/IC numbers this is forgiving; for names it's appropriately fuzzy.

- **Default date range starts Jan 1.** `dtpBookingDateFrom.Value = YearDay1(Now)` means the search defaults to year-to-date. For a hotel that books a year ahead, `MonthDay30(Now)` as the To date may exclude future bookings; users who want next month's bookings have to widen the range manually.

- **Dynamic SQL with concatenation in two places** — the AND/OR builder in `ListCustomers` and the report query in `PrintHistory`. The user-typed text goes straight into `LIKE '%...%'` without sanitization. A guest name containing an apostrophe ("O'Brien") will break the query. `SQL_WHERE_LIKE_Text` (used at the top of the chain) presumably does the escaping, but the manual `SQLText "AND GuestPassport LIKE '%" & ... & "%'"` lines below it don't. This is a real injection/escaping issue, though for an internal staff-only tool the impact is limited.

- **`lvBookings.Refresh` / `lvCustomers.Refresh` inside the loop** — refreshing after every row added is expensive but visually shows progress. For a hotel-scale database this is fine; for a large dataset you'd build the items first and call `lvBookings.Refresh` once at the end.

- **Hidden first columns.** Both ListViews have an `ID` column with width 0. The `lvCustomers` first column actually carries the *passport* (because the SELECT doesn't include a separate row identifier — the passport is the natural key for distinct customers). The `lvBookings` first column carries the actual numeric ID. Slightly inconsistent but functional.

- **`strGuestName` and `strGuestOrigin` parameters are unused** in `ListBooking`. Dead code worth cleaning up.

- **`PrintHistory` has dead/broken code in the contact filter branch.** Either it was working at some point against a different SQL helper, or it was always broken and just hasn't been hit because the typical flow doesn't reach it (most prints happen with just a selected customer, no contact override). Worth verifying.

- **ListView icons reference `ImageList1`** but always pass icon index `0, 0` to `Add`, so no icons actually show on rows. Just a pattern carried forward.

- **No empty-state messaging.** If the search returns no rows, both ListViews are simply empty. A "No customers found" hint would help.

- **`OpenRS` consistency.** This form uses `OpenRS(gstrSQL)` like the Dashboard and Booking forms — so the inconsistent `OpenSQL` in `frmAdmin` does look like a one-off oddity in that file.

- **The commented-out `PopulateGroupName`** at the top is dead code from a copy-paste off `frmUserMaintain` or similar — references `cboUserGroup` which doesn't exist on this form.

- **`txtGuestContact.SelectedItem.Text` for the print path** vs `Trim(txtGuestContact.Text)` — there's mild duplication where the selected customer's contact in the list could differ from what's typed in the search box. The print uses what's in the textbox, not what's in the selected ListView row. So if the user typed "555" to search, found a customer whose recorded contact is "01-555-1234", and clicked print, the report's contact filter would use "555" (the typed value), not the actual stored value. Subtle but probably not what users expect.
