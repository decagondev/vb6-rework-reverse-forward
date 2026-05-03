# modMain — Application Bootstrap Module

The system's actual entry point. Versioned 1.2.22 with no other modification dates recorded — making this likely one of the original modules. Tiny by line count (one Sub, a handful of API declarations) but architecturally important: **this is what runs first when the user double-clicks the executable.** Everything else in the system flows from here.

The VB6 project is configured (in the project file) to use `Sub Main` as its startup object rather than a startup form. So when Windows launches the .exe, control begins inside `modMain.Main()`. From there, `frmSplash` is shown, `frmSplash` shows `frmUserLogin`, login shows `frmDashboard`, and the rest of the system follows.

## What this module provides

The module has exactly two responsibilities:

1. **Initialize the Windows Common Controls library** with appropriate flags so that ListView, Toolbar, ImageList, and other ComCtl32 controls render correctly under modern Windows themes
2. **Show the splash form** to begin the application

That's the entire job. It's the bootstrap shim between Windows and the application proper.

## API declarations

```vb
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
```

Four Win32 API functions and a struct definition. All `Private` to the module — they don't leak out to the rest of the codebase. This is the only place in the system that talks directly to Windows DLLs.

- **`InitCommonControlsEx`** — the modern (IE3.0+) Common Controls initializer. Takes a struct specifying *which* control classes to initialize via a bitmask.
- **`InitCommonControls`** — the older (Win95) version, no parameters, initializes everything with defaults. Used as a fallback.
- **`LoadLibrary("shell32.dll")`** — manually loads shell32 to work around an XP-era crash bug with VB UserControls.
- **`FreeLibrary`** — releases the shell32 handle when done.

The duplicate declarations at the top (commented-out `LoadLibraryA` declared by alias, vs the live `LoadLibrary` aliased to `LoadLibraryA`) show the developer working through the API binding mechanics. The active form is the standard VB6 idiom.

## The Main subroutine

The function is short but does several things:

```vb
Private Sub Main()
    Dim iccex As InitCommonControlsExStruct, hMod As Long
    Const ICC_ALL_CLASSES As Long = &HFDFF&
    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_ALL_CLASSES
    End With
    On Error Resume Next
    hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls
        Err.Clear
    End If
    On Error GoTo 0
    frmSplash.Show
    If hMod Then FreeLibrary hMod
End Sub
```

Step by step:

**1. Build the InitCommonControlsExStruct.** The struct has two fields: `lngSize` (sizeof the struct, used by the API to validate it understands this version) and `lngICC` (the bitmask of control classes to initialize). The code uses `ICC_ALL_CLASSES = &HFDFF` — the combined mask for every known control class. The commented-out section above shows individual constants for animate, bar, cool, date, hotkey, internet, link, listview, native font, page-scroller, progress, tab, treeview, updown, userex, standard, and Win95 classes. The developer chose "everything" rather than picking and choosing.

**2. Pre-load shell32.dll.** The `LoadLibrary("shell32.dll")` call is described in the comments as a "patch to prevent XP crashes when VB usercontrols present." This is a real Windows quirk from the XP era — certain VB6 user controls would crash on initialization if shell32 wasn't already loaded into the process. By force-loading it before any forms initialize, the bug is sidestepped. The handle is held until just before the form shows, then released via `FreeLibrary`.

**3. Initialize Common Controls.** First try `InitCommonControlsEx` (modern API). If that fails (older Windows or missing IE 3+), fall back to `InitCommonControls` (Win95-era API with no parameters). The `On Error Resume Next` guard ensures the failover happens silently — no MsgBox, no LogError. If Windows is so old that neither succeeds, the app will probably crash later when it tries to instantiate a ListView, but at this stage there's no useful recovery anyway.

**4. Show the splash form.** This is the moment control hands off to `frmSplash`, which then drives the rest of the boot sequence. Because `frmSplash.Show` is non-modal but `Sub Main` doesn't have an explicit message loop, VB6's runtime does the right thing — when the last form closes, the application terminates.

**5. Free the shell32 handle** if it was successfully loaded. Cleanup hygiene, though the OS would reclaim it on process exit anyway.

## The commented-out code

About half the file is a commented-out earlier version of `Sub Main`. It's mostly identical to the live version with two differences:

- The earlier version used `ICC_STANDARD_CLASSES` (just intrinsic controls — buttons, textboxes) rather than `ICC_ALL_CLASSES`. The promoted version is more inclusive.
- The earlier version used `LoadLibraryA` (the ANSI alias) directly instead of the cleaner `LoadLibrary` alias declaration.

Both versions are functionally equivalent for this app's needs. The earlier version's commented code includes the full list of `ICC_*` constants (with MSDN URL) — kept in source for reference even though only the combined mask is used.

## The hyperlinked text in the source

Two oddities worth flagging:

```vb
'... show your main form next (i.e., [Form1.Show](http://Form1.Show))
     [frmSplash.Show](http://frmSplash.Show)
```

These look like Markdown-formatted hyperlinks: `[text](url)`. Inside a VB6 source file, this is **not** valid syntax — VB6 would treat the brackets as nothing meaningful and the parens as something else. So either:

- This file has been rendered through a Markdown formatter at some point that wrapped naked tokens in link syntax (could be from a code-review tool, a documentation generator, or even an editor that did automatic link detection)
- Or the file was hand-edited to add Markdown but somehow still compiled

Looking carefully: VB6 does compile `[frmSplash.Show](http://frmSplash.Show)` because `[frmSplash.Show]` is interpreted as a bracket-escaped identifier (a feature for using reserved words as identifiers), and `(http://frmSplash.Show)` would be a parameter list — but `http:` isn't a valid expression. So this likely doesn't compile as shown. **The file as pasted has been corrupted by a Markdown-aware viewer or paste operation somewhere in transit** — the actual file in the project almost certainly reads `frmSplash.Show` without the bracket/URL noise.

If this corrupted text was committed to the actual source file, the project would fail to compile. The fact that the application runs means the source-of-truth is clean. The corruption is in the version pasted here.

## Tips in trailing comments

Two tips at the bottom worth noting because they reflect lessons learned during development:

```vb
'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.
```

This explains design choices observed across the form analyses:

- **Tip 1** is why every form uses `PictureBox` as a container (`Picture1`, `Picture2`, `fraLogin` in many forms) instead of the more obvious `Frame` control. VB6 Frames don't repaint correctly under themed Windows (XP and later) — they paint over their child controls in old-style gray. PictureBoxes work around this. The whole codebase honors this convention.

- **Tip 2** is why almost every button uses `Style = 0 (Standard)` rather than `Style = 1 (Graphical)`. Graphical buttons let you set an image on the button, but they bypass the OS theme engine and render flat. The codebase uses standard buttons throughout, which means the buttons render with the user's current Windows theme (XP Luna, Vista Aero, Win7 Aero, Win10 flat, etc.). The mouse-cursor customization (`MouseIcon`, `MousePointer = 99 'Custom'`) is the codebase's preferred way to add visual flair without sacrificing themeability.

These tips are essentially in-code documentation of the visual-design conventions the rest of the codebase follows.

## Notable points and quirks

- **The module's job is over in ~15 lines.** Everything else in the file is either commented-out reference material or trailing tips. The actual work is minimal.

- **No error logging.** Unlike every other module in the codebase, `modMain` has no `LogErrorDB` or `LogErrorText` calls. This is *correct* — at this stage of startup, no database connection is established, no globals are initialized, nothing is set up. If the Common Controls init fails, there's nowhere to log it. The `On Error Resume Next` is the right tool for this stage.

- **`mstrModule` is not declared.** Every other module has `Private Const mstrModule As String = "mod___"`. This one doesn't — because the module has no error-handling blocks that would reference it. Consistent with the no-logging stance.

- **Called once, never again.** `Sub Main` runs at process start and never again for the life of the application. The `hMod = LoadLibrary("shell32.dll")` is paired with `FreeLibrary hMod` at the end — but the FreeLibrary runs *after* `frmSplash.Show`. Since `Show` is non-modal, this means FreeLibrary runs immediately after the form is *shown*, not after it's closed. The shell32 handle is held only during the brief window when forms are initializing. Reasonable — by the time the user is interacting with anything, the shell32 reference count Windows manages through `LoadLibraryA` doesn't matter (the DLL stays loaded because other components hold it).

- **`On Error Resume Next` is intentional and bounded.** It wraps only the ICC initialization and the LoadLibrary call. The `On Error GoTo 0` immediately after restores normal error handling. This is the right pattern — silence the bootstrap, but don't let the silence leak into the rest of the application.

- **`frmSplash.Show` is non-modal.** Compare to `frmDatabase.Show vbModal` from the splash code itself. Non-modal here is correct — the splash needs to do work (database checks, migrations) and then close itself before login begins. If it were modal, `Sub Main` would wait for it to close before reaching `FreeLibrary`, which is fine, but the current design also works: `Show` returns immediately, `FreeLibrary` runs, the splash continues its work, and when `frmSplash.Unload` happens, control returns to whatever's next on the form stack.

- **The "show form / FreeLibrary" ordering is subtle.** Looking at it again: the line `If hMod Then FreeLibrary hMod` runs immediately after `frmSplash.Show` returns, which is essentially immediately. So the shell32 patch is in effect for only milliseconds — long enough for the splash form to start initializing controls. The actual XP UserControl crash bug happens during *initial* control instantiation, so this brief window is exactly when shell32 needs to be loaded. Once forms have started instantiating controls, they hold their own DLL refs.

- **The `ICC_ALL_CLASSES` choice favors completeness over surgical optimization.** The commented-out version's code shows individual flags being commented in/out for the project's actual control usage. The promoted version doesn't bother — load everything, accept the small startup cost, ensure no unloaded class causes trouble. For a desktop app, this is a correct trade-off.

- **There's no `Sub_Initialize` or pre-Main hook.** VB6 lets you put initialization in a class's `Class_Initialize` event, but for a `Sub Main` startup, this `Sub Main` itself *is* the earliest hook. No other code runs before this.

- **No version check, no "is this OS supported" guard.** The fallback from `InitCommonControlsEx` to `InitCommonControls` is the only OS-version handling. The application implicitly assumes Windows 95+ with at minimum Internet Explorer 3.0 installed (for the comctl32.dll v4.7+ that supports `InitCommonControlsEx`). Realistic for any Windows machine made after about 1998 — and the developer's note about IE 3+ shows they considered the lower bound.

- **No splash-or-other-form choice logic.** `Sub Main` always shows `frmSplash` first. There's no command-line argument parsing, no "/silent" flag, no "/database=path" flag, no "/user=admin" flag. The app starts the same way every time.

- **`Modified On` is missing** from this module's header. Every other module has at least one modification date. The implication is that `Sub Main` was written once and never touched again — which makes sense for code that does one well-defined thing.

- **`Option Explicit` is *not* set** in this module. This is the only module we've analyzed without it. Probably an oversight — there are no undeclared variables in the actual code, so it doesn't matter, but it breaks the pattern.

- **The hyperlinked-Markdown text issue** (the `[frmSplash.Show](http://...)` lines) is a transmission artifact and not real code. If you opened the actual `.bas` file in the VB6 IDE, it would read just `frmSplash.Show`. Worth being aware of if the codebase ever gets re-exported through a Markdown-aware tool.

- **The module's true value is preserving the bootstrap conventions.** The commented-out section, the inline tips about XP themes, and the careful API ordering are all knowledge that would be hard to reconstruct without this file. Anyone modifying it should preserve the comments — they document non-obvious Windows behavior the application depends on.

- **No teardown.** `Sub Main` doesn't have a cleanup pass. When the application exits (last form closed, or `End` called), VB6 runs its standard shutdown — releasing COM references, unloading forms, closing DB connections (if still open). The `FreeLibrary hMod` is the only explicit cleanup, and it's done early. A more robust design might have a `Sub Cleanup` called from a centralized exit handler, but the codebase doesn't have one — `End` is called from various places (cancel buttons, login errors, frozen account) which immediately terminates without orderly cleanup. For a single-user desktop app, this is acceptable.

- **The module exemplifies the "do one thing, do it well" principle** at the smallest scale in the codebase. Every other module has multiple responsibilities (modCommon especially, blending utilities and SQL building). This one has a single, narrow job that takes 15 lines of actual code.
