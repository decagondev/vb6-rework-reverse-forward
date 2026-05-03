# StarHotel.vbp — Project Manifest Breakdown

The .vbp file is the VB6 equivalent of a project file (.csproj, package.json, pom.xml, etc.). It's a plain-text INI-format manifest that tells the VB6 IDE and compiler what to build, what to reference, and how. Reading this file gives a top-down view of the entire codebase as a buildable unit — and confirms (or contradicts) several inferences from the form/module analyses.

## Build target

```ini
Type=Exe
Startup="Sub Main"
ExeName32="StarHotel.exe"
Command32=""
ResFile32="StarHotel.res"
IconForm="frmBooking"
```

**It's a standalone .exe**, not an ActiveX DLL or OCX. **Startup is `Sub Main`** — confirms what we inferred from `modMain` (no startup form; entry is `modMain.Main`). **ExeName is `StarHotel.exe`** — the actual binary the user double-clicks.

**`IconForm = "frmBooking"`** is interesting. The .exe's icon (the one Windows shows in Explorer, the taskbar, and Alt-Tab) is taken from whichever form has an Icon property and is named here. Curious that it's `frmBooking` and not `frmDashboard` or `frmSplash` — `frmBooking` isn't the entry point or the most-used form, just one form in the middle of the workflow. Probably a development-time choice (someone set it once early on and it stuck) rather than a deliberate branding decision.

**`Command32=""`** means no default command-line arguments. The app expects to be launched with no args, consistent with `modMain.Main`'s lack of argument parsing.

**`ResFile32="StarHotel.res"`** points to a Windows resource file. Resource files in VB6 hold version info, embedded icons, string tables, and binary blobs. The version metadata block at the bottom of the .vbp populates this resource file at compile time.

## References (COM libraries the project depends on)

```ini
*\G{00020430-0000-0000-C000-000000000046}#2.0#0  OLE Automation
*\G{2A75196C-D9EB-4129-B803-931327F72D5C}#2.8#0  Microsoft ActiveX Data Objects 2.8 Library
*\G{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0  Crystal Report Viewer Control
*\G{B4741C00-45A6-11D1-ABEC-00A0C9274B91}#8.5#0  Crystal Reports 8.5 ActiveX Designer Run Time Library
*\G{00000600-0000-0010-8000-00AA006D2EA4}#2.8#0  Microsoft ADO Ext. 2.8 for DDL and Security
```

Five COM/ActiveX references, each with its GUID, version, and library file. Worth working through:

**OLE Automation** is implicit in every VB6 project — the base layer for all COM interop. Always present, never optional.

**ADO 2.8 (`msado15.dll`)** is the database-access library. This is what gives us `ADODB.Connection`, `ADODB.Recordset`, `ADODB.Field`, etc., used everywhere in `modDatabase`. ADO 2.8 was the last released version (2002) and shipped with Windows XP/2003 and later. It's frozen — Microsoft never released ADO 2.9 or 3.0; the line ended here.

**Crystal Report Viewer Control 8.0 (`crviewer.dll`)** is the visible widget — the toolbar, scrollbar, and rendering surface that `frmPrint` embeds. Version 8.0.

**Crystal Reports 8.5 ActiveX Designer Run Time Library (`craxdrt.dll`)** is the engine — what `CrApplication As CRAXDRT.Application` in `modGlobalVariable` refers to. **Version 8.5**, slightly newer than the viewer.

The viewer/runtime version mismatch (8.0 viewer + 8.5 runtime) is the kind of cross-version mixing that *usually* works because Crystal Reports kept binary compatibility within v8.x, but it could surface odd rendering bugs on certain Windows installs. A more disciplined project would match the versions; this one doesn't.

**Crystal Reports 8.5 was released in 2001, end-of-lifed in ~2005.** The fact that the project still uses it in 2021 tells us a few things:

- The deployment environment is locked down — customers must have CR 8.5 installed (typically via the merge modules that came with the original Crystal Decisions installer, or via the Crystal runtime redistributable)
- New installs would need to source these now-ancient runtimes, which Microsoft and SAP no longer distribute
- The reports themselves (`.rpt` files in the `App.Path\Report\` directory) are saved in CR 8.5 format
- Migrating to a newer reporting engine (Crystal XI, BusinessObjects, or replacing it entirely) would be one of the larger modernization efforts

**ADOX (`msadox.dll`)** is the schema-definition library — what `ADOX.Catalog` in `modDatabase.CreateData` uses to create empty .mdb files at first run. ADOX is the "X" (extension) library to ADO, providing programmatic access to schema operations (create table, alter, drop) that ADO itself doesn't directly expose. It's used in exactly one place: creating the initial database file.

## Components (ActiveX controls placed on forms)

```ini
Object={C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0; crviewer.dll  Crystal Viewer
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX  Common Controls
Object={86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCT2.OCX  Common Controls 2
```

The difference between References and Objects: References are libraries you call from code (`Dim x As ADODB.Connection`); Objects are controls you place visually on a form (the toolbox-and-drag-onto-a-form kind).

**`crviewer.dll`** is also listed as an Object because it's both a code library and a visual control — `frmPrint` has a Crystal Viewer control on it.

**`MSCOMCTL.OCX` (Common Controls 2.0)** provides the ListView, TreeView, Toolbar, ImageList, Slider, ProgressBar, StatusBar, TabStrip controls used throughout the application. Every form's toolbar (`tbrMenu`), every ListView (`lvUsers`, `lvBookings`, etc.), every ImageList (`ImageList1`) comes from this OCX. This is the single most-used component in the codebase visually.

**`MSCOMCT2.OCX` (Common Controls 2)** provides the DateTimePicker, MonthView, UpDown, and Animation controls. Used by date-input fields in `frmBooking` (check-in/check-out date pickers) and `frmReport` (report-as-on date picker).

The combination of these three OCXes plus the intrinsic VB6 controls (TextBox, Label, CommandButton, PictureBox, Image, Shape, Line, Frame, Timer, ComboBox, CheckBox) covers every visual element in the application. No third-party controls.

## Forms (in the order added to the project)

```ini
Form=Form\frmUserLogin.frm
Form=Form\frmUserChangePassword.frm
Form=Form\frmSplash.frm
Form=Form\frmFindCustomer.frm
Form=Form\frmBooking.frm
Form=Form\frmDashboard.frm
Form=Form\frmModuleAccess.frm
Form=Form\frmReport.frm
Form=Form\frmPrint.frm
Form=Form\frmReportMaintain.frm
Form=Form\frmRoomMaintain.frm
Form=Form\frmRoomTypeMaintain.frm
Form=Form\frmDatabase.frm
Form=Form\frmDialog.frm
Form=Form\frmUserMaintain.frm
Form=Form\frmAdmin.frm
```

**Sixteen forms total.** This is one more than the fifteen we analyzed — `frmDialog.frm` was referenced in form code (called via `frmDialog.Show vbModal` in the idle-watchdog timer in many forms) but never analyzed because we never received its source. Based on usage patterns, `frmDialog` is likely the auto-logout confirmation dialog or a generic message-dialog form.

The order in the .vbp tells the IDE and compiler the order to load and consider the forms — usually it reflects the chronological order forms were added to the project. Notable observations from the order:

**`frmUserLogin` is first.** This was probably the first form created in the project. Then `frmUserChangePassword`, then `frmSplash`. So the developer started with the login flow, added the password-change flow, then built the splash. Reasonable: get auth working, then add bootstrap.

**`frmFindCustomer` is fourth, before `frmBooking`.** Customer search came before booking creation. Either the developer prototyped lookup before write, or the search was the harder form to figure out and got built first.

**`frmDashboard` is sixth**, after both Booking and FindCustomer. So the central hub was added once the leaf forms it would launch existed. Backwards from how we'd document the system, but reasonable bottom-up development.

**`frmDatabase` is thirteenth (very late).** The first-run database picker was added much later — consistent with the form's `Added On : 13/07/2019` comment. This was a 2019 addition to support multiple databases per install.

**`frmDialog` is fourteenth, after `frmDatabase`.** Likely the last UI form added before `frmUserMaintain` and `frmAdmin`. Without the source we can't say more.

**`frmAdmin` is sixteenth (last).** The Void-authentication form was the final addition. Matches its status as a never-fully-enabled feature — added late, never wired through.

## Modules (in the order added)

```ini
Module=modCommon; Module\modCommon.bas
Module=modDatabase; Module\modDatabase.bas
Module=modGlobalVariable; Module\modGlobal.bas
Module=modFunction; Module\modFunction.bas
Module=modTextFile; Module\modTextFile.bas
Module=modEncryption; Module\modEncryption.bas
Module=modMain; Module\modMain.bas
```

**Seven modules.** The order tells us the development sequence:

1. **`modCommon`** — utilities and SQL builder (the foundation everyone depends on)
2. **`modDatabase`** — connection and schema (depends on modCommon's SQL builder)
3. **`modGlobalVariable`** — globals (added once enough variables had accumulated to warrant centralization)
4. **`modFunction`** — domain helpers (depends on the above)
5. **`modTextFile`** — file I/O (independent of the database stack)
6. **`modEncryption`** — password hashing (independent of everything except modCommon's helpers)
7. **`modMain`** — bootstrap (added last because the entry point is conceptually first but practically the last thing you write)

Note **`modGlobalVariable; Module\modGlobal.bas`** — the module name (`modGlobalVariable`) doesn't match the filename (`modGlobal.bas`). The filename was abbreviated, the in-source `Attribute VB_Name` is the longer name. VB6 allows this; the IDE shows the Attribute name, the filesystem shows the abbreviated name. Mildly confusing for anyone navigating the source folder.

The modification dates in each module's header confirm this order:
- `modEncryption`: 01/10/2014 (oldest)
- `modCommon`: 28/10/2014
- `modDatabase`: 17/12/2014
- `modTextFile`: 17/12/2014 (same day as modDatabase)
- `modFunction`: 04/01/2015
- `modGlobalVariable`: 19/05/2018 (much later — kept getting touched as new globals were added)
- `modMain`: undated (probably original)

So the actual chronological order of *first creation* (modEncryption → modCommon → modTextFile/modDatabase → modMain → modFunction → modGlobalVariable) doesn't quite match the .vbp order (which reflects when they were added to the project, not necessarily when written). The .vbp order suggests `modMain` was added to the project last, after all other modules existed and the developer was ready to wire up the entry point.

## Versioning

```ini
MajorVer=1
MinorVer=2
RevisionVer=22
AutoIncrementVer=0
```

**Version 1.2.22.** This is the value displayed by `frmUserLogin` ("Software Version 1.2.22") and stamped on every form's header comment. **`AutoIncrementVer=0`** means the revision doesn't auto-bump on every compile — the developer set it manually to 22, and it stays at 22 until manually changed. So 22 represents 22 deliberate revision releases since the 1.2.0 baseline, not 22 builds.

Looking at the form modification dates (oldest 01/10/2014, newest 19/05/2018), the project saw active development from late 2014 through mid-2018, with no recorded modifications after that. The "2014-2021" copyright range suggests light maintenance through 2021 (probably the 1.2.22 revision itself happened in 2021), but the bulk of the work was done by 2018.

## Version metadata (embedded in the .exe)

```ini
VersionComments="Full Version (Dark Mode)"
VersionCompanyName="Computerise System Solutions"
VersionFileDescription="Star Hotel"
VersionLegalCopyright="Copyright 2014-2021"
VersionProductName="Star Hotel"
```

This is the metadata Windows shows when you right-click the .exe → Properties → Details. Notable points:

**`VersionComments="Full Version (Dark Mode)"`** is the most informative line. It tells us:

1. There exists or existed a "Lite Version" or "Demo Version" or "Trial Version" — the "Full" qualifier wouldn't be there otherwise. The commented-out kill-switch in `frmSplash` (the one that would have expired the system on Jan 1, 2021) supports the existence of a separate trial build. The commented `lblDemo` text in `frmUserLogin` ("This free demo is valid until ...") also supports this. So the codebase has at least two compile profiles: the full version and a time-limited demo.

2. There exists or existed a "Light Mode" version. The dark gray `#303030` theme we've seen across every form is the "Dark Mode" variant. Earlier commented-out `Me.BackColor = RGB(184, 209, 255)` and `Me.BackColor = &HFFC000` lines in various forms hint at a light-themed variant that was eventually replaced or co-exists in a separate build.

So `VersionComments` reveals that what we've been analyzing is the full, dark-themed release — and that the codebase at one point or still does support a demo build and a light theme through compile-time configuration (probably `#If` directives that we haven't seen in the analyzed code).

**`VersionCompanyName="Computerise System Solutions"`** confirms the developer/vendor. Same name as the GenWord-obfuscated database password ("Computerise"). Same as the copyright line at the bottom of every form.

**`VersionProductName="Star Hotel"` vs `VersionFileDescription="Star Hotel"`** — both say "Star Hotel" matching the seed Company.CompanyName. So this build is **branded for a specific customer** (Star Hotel in Kuala Lumpur), not a generic "Hotel Booking System". Interesting because the source code's `COMPANY_PRODUCT_NAME` constant says "Hotel Booking System" — the *product* is generic but the *executable* is customer-branded. So either:

- This particular .vbp is a customer-customized build, and there are other .vbp files for other customers
- Or "Star Hotel" was the development codename and it just happened to ship that way

Combined with the `IconForm = "frmBooking"` choice (an arbitrary-feeling default), the app feels like it was originally built for one customer and then has occasional generic touches added when reusing the code for others.

## Compilation settings

```ini
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
```

**`CompilationType=0`** means "compile to native code" (P-code is `1`). So the .exe contains compiled x86 machine code, not interpreted VB pseudo-code. Faster runtime, larger executable, and (relevantly) **harder to decompile back to source**. P-code can be reverse-engineered to readable VB6 with tools like P32Dasm; native-compiled executables can be disassembled but you only get assembly, not source. The choice means the `GenWord` obfuscation of the database password is somewhat better hidden than it would be in a P-code build — though still recoverable with effort.

**`OptimizationType=0`** is "Optimize for fast code" (1 is "small code", 2 is "no optimization"). Standard release-build choice.

**`BoundsCheck=0`, `OverflowCheck=0`, `FlPointCheck=0`, `FDIVCheck=0`, `UnroundedFP=0`** — all the runtime safety checks are **disabled**. Array out-of-bounds, integer overflow, floating-point underflow, the Pentium FDIV bug check — all off. This is the standard "release build" configuration in VB6: turn off the checks for performance, accept that any latent bug will manifest as a crash or silent corruption rather than a controlled error.

For a VB6 desktop app of this size the impact is negligible — none of these checks are common bug causes — but it does mean that a programming error like `ConvInt` returning a value larger than 32,767 would silently overflow rather than raise a runtime error.

**`CodeViewDebugInfo=0`** — no debug symbols in the .exe. So crash reports or memory dumps from production wouldn't show useful function names. Standard release build.

**`NoAliasing=0`** means VB6 won't assume that two ByRef parameters might point to the same memory. Allows certain optimizations but means a pathological caller passing the same variable twice could see undefined behavior. Edge case that doesn't matter for this codebase.

## Threading and runtime

```ini
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
DebugStartupOption=0
```

**`StartMode=0`** is "Standalone" (vs `1` "ActiveX Component"). Confirms it's an .exe meant to be launched by a user, not consumed by another app.

**`Unattended=0`** means the app expects user interaction (vs unattended/server-mode). Confirms it's a desktop UI application.

**`MaxNumberOfThreads=1`** is the most architecturally significant line in this section: **the application is single-threaded**. Every form load, every database query, every button click runs on the same UI thread. Long operations freeze the UI — there's no background worker, no async pattern. This matches the synchronous-blocking nature visible across the form code.

**`ThreadPerObject=0`** is moot when MaxNumberOfThreads is 1.

## MTS section

```ini
[MS Transaction Server]
AutoRefresh=1
```

Microsoft Transaction Server settings. Inactive for a Standalone .exe — MTS was a middleware layer for distributed transactions across server components. The `AutoRefresh=1` is a default IDE setting, not a runtime behavior. Effectively dead config for this project type.

## Notable points and quirks

- **Title is "Star Hotel"** but `COMPANY_PRODUCT_NAME` constant in the source is "Hotel Booking System" and the form designer's `lblBusinessName` defaults to "Room Booking System". **Three different product names in the same codebase**, surfaced in different places. The .exe will display "Star Hotel" in window title bars (set from `App.Title`); the in-form labels will display whichever string they were assigned at design time or load time. A user looking at the running app sees a mixture.

- **`IconForm="frmBooking"`** is non-obvious. The .exe icon comes from this form. Typically you'd pick the splash or dashboard form. Probably a leftover from early development.

- **Crystal Reports 8.5 dependency is a deployment liability.** This runtime has been EOL for 20 years. New installs require sourcing it from archived installers, which raises legal-licensing questions (Crystal Decisions/Business Objects/SAP have changed hands multiple times, and CR 8.5 redistribution rights are murky). Modernization candidate #1 if this app needs to be deployed to new sites in the next few years.

- **ADO 2.8 is the last version.** Forever-frozen. Works on every Windows from XP forward. Will continue to work on future Windows as long as the legacy ODBC/OLEDB stack is preserved. Microsoft has occasionally hinted at deprecating Jet, but as of writing it's still shipped with Windows. So the database stack itself is durable; the reporting stack is not.

- **No third-party controls.** Zero dependencies on commercial control suites (no FarPoint, no ComponentOne, no Sheridan, no DBI Tech). Everything is Microsoft-stack. Good for portability and licensing simplicity; means the UI is constrained to what stock controls can do (which is why the codebase invents tricks like the Wingdings password masking and the PictureBox-as-Frame substitution).

- **Sixteen forms but only fifteen analyzed.** `frmDialog.frm` was never received. Based on call sites in the form code (`frmDialog.Show vbModal` from `tmrClock_Timer` handlers when `intTick > gintUserIdle`), it's the auto-logout warning dialog — the form that would have appeared after the idle timeout and prompted the user to confirm continued session or log out. Currently inert because of the `gintUserIdle = 0` debug override in `frmUserLogin`. If you have the `.frm` file for it, that would complete the analysis.

- **Seven modules, all .bas (standard modules), no class modules (.cls).** The codebase is procedural, not object-oriented. Even the SQL builder is a sequence of subs rather than a builder class. Consistent with the era and the team size — class-based design would add complexity without obvious benefit at this scale.

- **`AutoIncrementVer=0`** means the version stays at 22 across builds until manually changed. So if a developer recompiles without bumping the version, two builds with different binaries can both report 1.2.22. For an internal tool this is fine; for a distributed product it would be a pain to support ("which 1.2.22 do you have?").

- **All checks disabled in compile settings.** Standard for a release build. Means any crash in production yields a generic Windows error dialog rather than a controlled VB6 runtime error. The form-by-form `On Error GoTo CheckErr` handlers catch most issues at the application level; what slips through becomes a Windows-level crash.

- **Native-compiled (CompilationType=0).** Faster runtime, harder to decompile. Means the obfuscated `GenWord` password is reasonably hidden — recoverable by a skilled reverse-engineer with IDA Pro, but not by a casual snoop with strings.exe.

- **The .vbp is the closest thing to a project-level "table of contents".** With it, you can build the entire codebase from source. Without the matching `.frx` files (binary form resources), you can't — the icons, mouse cursors, embedded images, and ListView columns are stored in those binary side-cars. So a complete checkout needs all `.frm`, `.frx`, `.bas`, `.res`, and `.vbp` files together.

- **`VersionComments="Full Version (Dark Mode)"`** is the single most informative line in the file. It tells us this is one of multiple builds of the same source, with the variants distinguished by feature set ("Full" vs trial/demo) and by visual theme ("Dark" vs implied "Light"). The codebase as analyzed represents one specific build configuration; the full project may have conditional compilation directives we haven't seen that produce different outputs.

- **The `2014-2021` copyright** in `VersionLegalCopyright` is the upper-bound date we have for any active modification. Combined with the form modification dates (latest 19/05/2018) and the absence of any post-2021 markers in the seed data or kill-switch logic, the system was effectively in maintenance mode by 2018-2019 with the 2021 date representing a final compile-and-stamp rather than substantive new work.

- **Customer-specific build (`Star Hotel`) vs generic source.** The source code uses generic constant names (`COMPANY_PRODUCT_NAME = "Hotel Booking System"`) but the executable is branded for one customer. This is a classic ISV pattern — one source codebase, multiple customer builds, each with its own .vbp and resource file. If "Star Hotel" was the only customer this code was ever built for, the duality is just historical clutter; if there were five customers, there are five .vbp files somewhere with five different `Title=` and `VersionProductName=` lines.

- **Module file `modGlobal.bas` vs internal name `modGlobalVariable`.** The mismatch suggests a rename mid-project — the file was first named `modGlobal.bas`, then renamed inside the file to `modGlobalVariable` (probably for clarity), but the filesystem name was never updated to match. VB6 doesn't enforce consistency. Cosmetic issue; doesn't affect compilation.

If you want, the natural next step would be to look at the **`StarHotel.res` resource file** (if you have it) — it would contain the version metadata, embedded icons, and any string tables. Or, looking further, the `.frx` binary files for the forms — those contain the embedded icons and cursors and would round out the picture. But the .vbp by itself confirms the architecture we've inferred from the analyzed source: native VB6 .exe, single-threaded, Sub Main entry, ADO 2.8 + Jet 4 backend, Crystal Reports 8.5 for printing, MSCOMCTL/CT2 for UI controls, customer-branded build (Star Hotel) of a generic Hotel Booking System.
