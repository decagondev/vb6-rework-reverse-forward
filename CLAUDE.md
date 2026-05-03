# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository purpose

This is a **reverse-then-forward** rework of a VB6 hotel reservation system ("StarHotel" by Computerise System Solutions, 2014–2021). The repo currently holds:

- `ORIGINAL-CODE/star-hotel-vb6-main/` — the unmodified upstream VB6 source, treated as a frozen reference. Do not edit files here.
- `docs/` — the reverse-engineering knowledge base. Each form and module has a long-form analysis written from reading the original source. This is the primary artifact so far.
- `README.md` — placeholder.

The "forward" half (a modernized reimplementation) has not started. Until that begins, the work in this repo is documentation: extending, correcting, and cross-linking the analyses in `docs/`.

## Build / run / test

There is **no build system in this repository**. The original code under `ORIGINAL-CODE/` is a Visual Basic 6 project (`StarHotel.vbp`) that only builds inside the VB6 IDE on Windows with these COM dependencies installed: ADO 2.8 (`msado15.dll`), ADOX (`msadox.dll`), Crystal Reports 8.5 ActiveX runtime (`craxdrt.dll`) + Viewer 8.0 (`crviewer.dll`), MSCOMCTL.OCX, MSCOMCT2.OCX. A pre-built `StarHotel.exe` plus all redistributables sit in `ORIGINAL-CODE/star-hotel-vb6-main/publish/`. Database is a password-protected Access `.mdb` (Jet 4.0).

Documentation changes do not need a build — they're plain Markdown. There is no linter, test suite, or CI configured.

## Documentation structure

Three parallel doc trees under `docs/`, each mirroring a layer of the original codebase:

- `docs/forms/` — one Markdown file per `.frm` form (16 forms total), plus `README.md` (system overview + form catalog) and `UI-FLOW.md` (lifecycle phases + UX flow). `assets/UX-FLOW-DIAGRAM.svg` is the rendered diagram. Note: `DIALOG-FORM.md` (covering `frmDialog`, the idle-logout warning) was added after the forms `README.md` and `UI-FLOW.md` were written, so those overview files don't yet list it in their catalogs — worth folding in if either is revised.
- `docs/modules/` — one file per `.bas` module (7 modules), plus `DATA-FLOW-DIAGRAM.md` (layer-cake architecture write-up) and `assets/DATA-FLOW-DIAGRAM.svg`.
- `docs/project/` — files describing the build/packaging artifacts (`PROJECT-VBP.md`, `RES-FILE.md`, `VBW-FILE.md`, `MANIFEST-FILE.md`).
- `docs/BUGS-MITIGATIONS.md` — cross-cutting catalog of ~150 bugs, security flaws, half-implemented features, and architectural smells synthesized from the per-form/per-module analyses, with severities (Critical/High/Medium/Low) and concrete mitigations. Use this as the triage list for the eventual rewrite. When extending a form/module analysis with a new finding, also append it to the appropriate category here and bump the summary count.

When adding new analysis files, follow the existing voice: long-form prose, deep dives into quirks, named code citations (function names in backticks), tables for catalogs, and explicit calls-out of bugs/smells/legacy artifacts. The docs are written for a reader who will eventually rewrite this system — flag anything that should/shouldn't survive the port.

## Application architecture (essentials)

These are the load-bearing facts about the original system that any documentation or future port must preserve. Read the full files in `docs/` for nuance.

**Entry point.** `Sub Main` in `modMain.bas` is the startup (set in `StarHotel.vbp` as `Startup="Sub Main"`, not a startup form). It initializes Win32 Common Controls, then shows `frmSplash`, which validates `Config.txt`, opens or migrates the `.mdb`, instantiates Crystal Reports, and hands off to `frmUserLogin`.

**Hub-and-spoke UX.** `frmDashboard` is the central form. Every other operational form launches from it and returns to it on close — no cross-traffic between leaves. Two nested sub-hubs exist: `frmReport`→`frmReportMaintain` (gated by hardcoded `expert` developer password) and `frmRoomMaintain`→`frmRoomTypeMaintain`. F-keys F2–F8 are the navigation primitive.

**The SQL builder is the architectural defining choice.** `modCommon` exposes a family of `SQL_SELECT`, `SQL_FROM`, `SQL_WHERE_*`, `SQL_SET_*`, `SQLData_*` subs that **mutate a global `gstrSQL` string**. Forms call this sequence to build a query, then pass `gstrSQL` to `modDatabase` for execution. This global is also how `frmPrint` receives its query from callers. Consequence: no parallel queries, no thread safety, and comma-management bugs are possible.

**Database access.** `modDatabase` owns two ADO connections — `ACN` (primary) and `DAT` (parallel, used during migrations to read the old DB while writing the new). Reads go through `OpenRS`/`OpenSQL` returning recordsets; writes go through `QuerySQL` (per-statement transaction). **Multi-step business operations are not wrapped in a transaction anywhere** — partial-save risk is accepted. `CreateDB` defines the entire schema (12 tables) inline and `CreateSampleData` seeds it.

**Globals as session state.** `modGlobalVariable` declares ~18 publics that hold the user session (`gstrUserID`, `gintUserGroup`, `gstrUserPassword`, `gstrUserSalt`, `gintUserIdle`, `gblnUserChangePassword`), company info, current report context, and the SQL buffer. Permission checks read these globally; audit-trail columns (`CreatedBy`, `LastModifiedBy`) read `gstrUserID`.

**Security model.** Hand-rolled `GoldFishEncode` hash + 4-char salt in `modEncryption` — passwords capped at 8–10 chars. Default credentials `admin`/`admin` are hardcoded into the login screen text and auto-filled when the user clicks the copyright label. Permissions live in a `ModuleAccess` cell-by-cell matrix; only Groups 1 (Admin) and 4 (Clerk) are used in practice though the schema reserves four. Several disabled-but-wired features exist: `gintUserIdle = 0` debug override kills the idle-logout watchdog; the Void operation in `frmBooking` and `frmAdmin` is intentionally commented out at both ends. The full enumeration — 16 auth/credential issues, 9 SQL-injection issues, 12 authorization issues, plus disabled features and other categories — lives in `docs/BUGS-MITIGATIONS.md` with severities and mitigations; consult it before recommending fixes or porting decisions.

**Reports are metadata-driven.** Adding a report means inserting a row into the `Report` table (with `ReportQuery`/`SubQuery`/`NullQuery`) and dropping a `.rpt` file in `App.Path\Report\`. `frmReport` assembles the query at runtime from the row + date filters + variable substitution (`$UserID$`, `$BookingID$`). Crystal Reports opens its own DB connection, parallel to the main `ACN` stack.

**Soft delete and audit trails.** Every editable entity has an `Active` flag; hard-delete buttons are universally disabled. `Booking` and `Room` writes get full-snapshot history rows in `LogBooking` / `LogRoom`. `Booking` is denormalized — it stores a snapshot of room details at booking time so old receipts stay accurate after rooms are repriced.

**Error logging has two paths.** Most code calls `LogErrorDB` (writes to `LogError` table). Code that runs before the DB is open (connection helpers, encryption, splash) falls back to `LogErrorText` which writes a flat `Error.txt` in `App.Path`.

**Theming conventions.** Every form uses `PictureBox` instead of `Frame` (Frames don't repaint correctly under XP+ themes). Buttons use `Style = 0 (Standard)` not `Style = 1 (Graphical)` so the OS theme applies. These conventions are documented at the bottom of `modMain.bas`.

## Repo conventions

- The `ORIGINAL-CODE/` tree is a frozen reference copy. Do not modify it. New work goes under `docs/` (or a future `src/` for the rewrite).
- Each doc file has been committed with a focused message ("Document X module functionality"). Match that style for new docs.
- The docs already cite the in-source dates (`Modified On : DD/MM/YYYY`) and version numbers (`Version : 1.2.22`) from the original module headers — preserve those references when extending analyses, since they're load-bearing for understanding what changed when.
- Diagrams live as `.svg` under each section's `assets/` folder, referenced from a `.md` companion file. Keep that pairing.
