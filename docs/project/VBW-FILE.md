# StarHotel.vbw — Window State File Breakdown

The `.vbw` file is the **window-state companion** to the `.vbp` project file. Where `.vbp` describes *what* the project contains (forms, modules, references, build settings), `.vbw` describes *how the IDE displayed those things* the last time the developer closed the project — window positions, sizes, and minimize/maximize states.

This file is purely a development-time artifact. It has zero impact on the compiled `.exe` and is irrelevant at runtime. Its job is to restore the developer's IDE workspace exactly as they left it: same windows open, same positions on screen, same maximized/minimized states. If you delete the .vbw and reopen the .vbp, the IDE just opens with default window placements — nothing breaks.

It's small (one line per project file) and easy to ignore, but it leaks interesting clues about the developer's workflow, monitor setup, and which files they were actively editing.

## File format

Each line follows one of two formats:

**For modules (single window per file):**
```
<name> = X, Y, Width, Height, State
```

**For forms (two windows per file — designer + code):**
```
<name> = X1, Y1, W1, H1, S1, X2, Y2, W2, H2, S2
```

Forms have *two* sets of coordinates because each form has both a **visual designer** (the WYSIWYG form editor) and a **code window** (the event-handler code-behind). Modules have only a **code window**, so one set.

The state codes:
- **`C`** = Child window (docked or floating inside the IDE's MDI client area). The vast majority.
- **`Z`** = Maximized (the window is maximized within the IDE).
- **`I`** = Iconic (minimized to a thumbnail). None present here.
- **`R`** = Restored (normal windowed state). None present here either — the file uses C and Z exclusively.

The X/Y/Width/Height numbers are in pixels relative to the IDE's MDI client area.

## What the actual file tells us

Walking through the entries:

```
frmUserLogin = -124, 61, 796, 517, C, -7, 4, 1072, 619, C
```

The form has two windows. The **designer** is at X=-124, which is **off the left edge of the IDE's client area**. The width is 796, so the right edge is at x=672. The form designer is positioned mostly off-screen to the left, with only the right portion visible.

The **code window** is at X=-7 (essentially the left edge), 1072 wide, 619 tall. So the code window is *much larger* than the designer — taking up most of the screen. This is the pattern of a developer who spends most of their time looking at code and only occasionally peeks at the designer.

```
frmUserChangePassword = 22, 29, 942, 485, C, 66, 87, 986, 543, C
```

Both designer (942×485) and code (986×543) at modest sizes near the top-left of the IDE. Standard layout — neither dominates.

```
frmSplash = -7, 40, 1005, 427, C, -12, 87, 908, 543, C
```

Designer at near-full width (1005), positioned at the very left edge. Code window slightly off-screen left (-12). Both windows fairly large. This was a form actively being edited — the developer wanted lots of canvas to see the splash's big visual layout.

```
modCommon = 43, 50, 963, 506, C
modDatabase = 1, 76, 921, 532, C
modGlobalVariable = -101, 17, 819, 473, C
modFunction = 0, 0, 952, 615, C
modTextFile = 4, 129, 924, 585, C
modEncryption = 44, 59, 963, 514, C
```

All six modules have similar code-window sizes (920-960 wide, 470-615 tall) clustered in the top-left of the IDE. Notable: `modGlobalVariable` is at X=-101 (substantially off-screen left). `modFunction` is at the absolute origin (0,0). The clustering suggests the developer kept all modules in a similar position so that switching between them didn't require eye refocusing — open one, see code, switch to the next, see code in the same screen region.

```
frmFindCustomer = 89, 25, 1009, 481, C, 73, 40, 993, 496, C
frmBooking = 27, 68, 947, 524, C, 44, 58, 964, 514, C
frmDashboard = 72, 23, 992, 479, C, 22, 29, 942, 485, C
frmModuleAccess = 35, 144, 955, 600, C, 72, 36, 992, 608, C
frmReport = 22, 29, 942, 485, C, 34, 35, 954, 491, C
frmPrint = 22, 29, 942, 485, C, 59, 53, 979, 509, C
frmReportMaintain = -5, 149, 915, 605, C, 56, 35, 976, 491, C
frmRoomMaintain = 24, 33, 944, 489, C, 63, 9, 983, 465, C
frmRoomTypeMaintain = 44, 58, 955, 514, C, 56, 24, 967, 480, C
frmDatabase = 83, 36, 963, 384, C, 132, 174, 1012, 522, C
frmDialog = 66, 87, 998, 474, C, 55, 86, 987, 473, C
frmUserMaintain = 66, 87, 904, 543, C, 45, 56, 977, 443, C
frmAdmin = 37, 53, 945, 509, C, 66, 58, 974, 514, C
```

The transactional and admin forms all sit in similar positions with similar sizes — between 900-1010 wide, 380-610 tall. Several share *identical* coordinates:

- `frmDashboard`, `frmReport`, `frmPrint` all have designer at `22, 29, 942, 485` — same exact rectangle. These three forms were last opened together or last positioned together.

The repeated coordinates are the IDE's default position for newly-opened windows. So forms with default coords were opened but then closed without much repositioning — the developer didn't customize their layout.

The forms with *unique* positions (`frmUserLogin`, `frmSplash`, `frmModuleAccess`, `frmReportMaintain`, `frmDatabase`) are likely the ones the developer was actively working on near the end of the development session.

```
modMain = 7, 28, 887, 376, Z
```

**The standout entry.** Position `7, 28` (near top-left) with size `887, 376`, and state `Z` — **maximized**. This is the only file in the whole list with state `Z`. So when the developer last saved the project, **modMain was the active maximized window** in the IDE.

This is the single most informative line in the .vbw. It tells us:

The developer's *last action* before closing the project was looking at `modMain.bas`. The bootstrap code — the `Sub Main`, the Common Controls initialization, the corrupted-Markdown `frmSplash.Show` line. Maybe they were debugging a startup issue, or making the final tweak before the 1.2.22 release build, or perhaps they were responding to a user-reported issue about app launch behavior. Whatever the reason, modMain was foreground at session close.

## Notable points and quirks

- **The off-screen positioning of multiple windows** (`frmUserLogin`'s -124, `frmSplash`'s -7 and -12, `modGlobalVariable`'s -101) suggests the developer's IDE was on a multi-monitor setup with the IDE spanning across screens, and these windows were positioned on the secondary monitor (which would have *negative* coordinates relative to the primary). When the .vbw was saved, those positions were captured even though they're "off-screen" from a single-monitor viewpoint. Anyone reopening this project on a single-monitor setup would find some windows initially invisible until they manually drag them back into view.

- **No iconified (minimized) windows.** State `I` doesn't appear. So when the project was last saved, every form/module was either visible (C) or maximized (Z), not minimized. The developer didn't habitually iconify windows for later use — they just opened, edited, and either closed or left visible.

- **Most windows are between 900-1000 pixels wide.** Combined with the consistent ~500 pixel heights, this suggests the developer's IDE was running on a monitor of roughly 1280-1600 pixel width. Not 4K, not 1080p stretched — typical of an early-2010s development machine, consistent with the project's 2014-2018 active-development window.

- **The repeated coordinates `22, 29, 942, 485`** appear three times across `frmDashboard`, `frmReport`, and `frmPrint`. This is the VB6 IDE's default "place new window here" position. Forms with these coordinates were opened-and-closed without the developer customizing their layout — probably opened briefly to look at something then closed.

- **`modFunction` at exactly `0, 0`** is also a default position. A developer who opened modFunction without moving it would get this.

- **The largest code window is `frmUserLogin` at 1072×619.** This was the form the developer needed most screen real estate for. Combined with its position (-7, 4, near top-left and almost full-screen), this was the focus form — the one the developer kept maximally visible for editing. Matches its 19/05/2018 modification date as one of the latest-touched files.

- **`modMain` is the smallest maximized window at 887×376.** That's an odd combination — small but maximized. The interpretation: the IDE's MDI client area was 887×376 pixels at the moment of save (perhaps the developer had docked the toolbox, properties window, and project explorer to take up most of the screen, leaving modMain a small maximized region). Or alternatively, modMain's size was set first, then the IDE was resized smaller, and the maximized state preserved through the resize.

- **Designer + code window ratio per form.** For most forms, the code window is similar size or larger than the designer. For `frmUserLogin` specifically, the code window (1072×619) is *much* larger than the designer (796×517). This pattern says "I'm working on the code, not the layout." Useful inference: the form's visual design was settled, but the login flow's logic (the `cmdOK_Click` validation, the `gintUserIdle = 0` debug override, etc.) was being actively worked on.

- **Reverse pattern for `frmDatabase`:** designer is 963×384 (wider but shorter), code window is 1012×522 (larger overall). Both are sized larger than most other forms, suggesting the developer was doing substantial work on both the layout and the logic. Matches `frmDatabase`'s 13/07/2019 addition date as a relatively recent form.

- **No `gridx`/`gridy` settings.** Some VB6 IDEs save grid-snap state in the .vbw. This file doesn't — either the developer disabled grid snapping or the version of VB6 in use doesn't write those values.

- **No window-Z-order information.** The .vbw doesn't tell you which window was *on top*, just where each was positioned. So we know modMain was the active maximized window, but if multiple non-maximized windows overlapped, we can't tell which was foreground.

- **The file has 23 lines for 23 project items** (16 forms + 7 modules). Every file in the .vbp got an entry — none were filtered out for never-opened. So the developer opened every file at some point during the project lifetime, and the IDE recorded its last-known position. If a file had never been opened, it would simply be absent from the .vbw.

- **The single-line-per-file format is unusually compact** for an IDE state file. Modern IDE state files (Visual Studio's `.suo`, IntelliJ's `.idea/workspace.xml`) can run to megabytes with detailed state. VB6's .vbw is barely a kilobyte. The lower fidelity reflects an era when "what window is open and where" was the entire scope of restorable state — no expanded code-tree state, no breakpoints, no debugger watch expressions, no recent-file lists, none of that lives here.

- **The .vbw is `.gitignore` material in modern parlance.** It's a per-developer artifact — if two developers worked on the same project from different machines, they'd each have a different .vbw reflecting their own monitor setup and window preferences. Committing it to source control means every developer's window layout fights for dominance with every commit. Most teams either ignore it or accept the noise. This project's .vbw being included alongside the .vbp suggests it's either single-developer (no collaboration friction) or the developer didn't consider this kind of thing.

- **Combined with the form analyses, the .vbw confirms the active-development pattern**:
  - `frmUserLogin`: dominant code window (1072×619) — actively worked
  - `frmSplash`: large designer + slightly off-screen large code — both being edited
  - `frmDatabase`: above-average sized both windows — recent addition being polished
  - `modMain`: maximized at session close — the last thing looked at
  
  Everything else (the transactional forms, the admin forms, most modules) sits at default-ish coordinates, suggesting they were stable code that wasn't being actively touched in the final edit session.

- **What the .vbw doesn't tell us:** which file was "the project's startup file" being run when the developer pressed F5 (that's in the .vbp's `Startup="Sub Main"`); which lines had breakpoints (that state is per-session and not persisted); what the developer was about to commit (no version control state); what tasks they had in progress (no TODO tracking). For all that, you'd need supplementary files that VB6 doesn't generate — the .vbw genuinely is just window-positions.

- **The developer's monitor was probably ~1366×768 or 1280×720.** Most code windows are around 900-1000 wide and 500-600 tall, with the IDE chrome (toolbars, project explorer, properties window) taking the remaining space. That width-height profile fits a netbook or budget laptop of the 2014-era — common for in-house development tools or kiosk-PC vendors. Consistent with the project's positioning as a desktop POS app for Malaysian hotel operations.

In summary: the .vbw is a small, lossy snapshot of the developer's last IDE state. The most actionable insight is that **`modMain` was the maximized active window** at project save — meaning the bootstrap code was the developer's last focus. Everything else is forensic detail about which forms were actively being edited (UserLogin, Splash, Database) versus which had stable layouts (Dashboard, Report, Print, etc., all at default coordinates). It's the kind of file you can safely ignore for understanding the application but read with interest if you're building a picture of how the developer actually worked.
