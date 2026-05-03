# The Manifest XML — Source-Form Breakdown

This is the **source form** of the application manifest that was compiled into `StarHotel.res` and embedded in the `.exe`. We saw the binary representation when the .res file was extracted; this is the human-authored XML that produced it.

The two are functionally identical — same content, same effect on Windows when the .exe runs. Comparing them side by side gives some interesting details about the build process and the development workflow. The previous breakdown of `StarHotel.res` covered what each section *does* in the operating system; this breakdown focuses on what's different between the source XML and the binary, what that tells us, and what new observations the source form enables.

## What's the same

The structural and semantic content matches the .res payload exactly:

- **Assembly identity:** "Star Hotel", version 1.2.0.22, win32, x86
- **Description:** "Star Hotel"
- **Dependency:** Microsoft Windows Common Controls 6.0.0.0 (the themed comctl32, with the standard Microsoft public-key token)
- **Trust info:** asInvoker privilege level, no UI access
- **Compatibility:** Vista, 7, 8, 8.1, 10/11 GUIDs

So this XML, when fed through `mt.exe` (Microsoft's Manifest Tool), produces the binary resource we already analyzed. The transformation is mechanical — XML in, RT_MANIFEST resource out, Windows resource framing wrapping it.

## What's different (or revealing)

A few details that the source XML exposes which the binary version flattened or normalized:

### Tab indentation

The source XML uses **hard tabs** for indentation (the `\t` character), while the binary form had whatever spacing came out of the build tool — turned out to be CRLF line endings with a final pad of two spaces. Tab indentation is a stylistic choice that matters for one practical reason: if anyone diff'd this manifest against the .res-extracted version using a tool that's whitespace-sensitive, the indentation differences would create false-positive change indicators. The XML semantics are identical.

### The Markdown-corrupted line

```xml
<assemblyIdentity name="[Microsoft.Windows](http://Microsoft.Windows).Common-Controls" ...
```

The dependency's name field reads `[Microsoft.Windows](http://Microsoft.Windows).Common-Controls` — but in the actual .res-extracted binary, this same line read `Microsoft.Windows.Common-Controls` (clean, no brackets, no URL).

This is the **same Markdown-corruption pattern we saw in `modMain.bas`** when it was pasted with `[frmSplash.Show](http://frmSplash.Show)`. Some viewer or paste pipeline along the way detected `Microsoft.Windows` as something link-like and auto-wrapped it in Markdown link syntax. The actual on-disk file almost certainly reads `Microsoft.Windows.Common-Controls` — because:

- The `.res` file we extracted contained the clean form
- If the actual XML had the bracketed/URL form, the manifest would be invalid (the publicKeyToken check would fail to resolve "[Microsoft.Windows](http://Microsoft.Windows).Common-Controls" as an assembly name)
- If the manifest were invalid, the .exe would either fail to load Common Controls 6 themes or fail to start altogether
- The application demonstrably runs and demonstrably has themed controls

So the source-of-truth is clean. The corruption is in the paste, not in the file.

This is now the **second time** the same Markdown-link-injection corruption has affected files passed through the analysis pipeline (modMain.bas was the first). If you're using a tool that auto-detects what looks like dotted identifiers and wraps them in fake links, it's worth knowing about — it changes the apparent content but (so far) hasn't broken the actual on-disk artifacts.

### XML declaration encoding attribute

Source:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
```

The .res-embedded version had:
```xml
<?xml version="1.0" standalone="yes"?>
```

The source declares UTF-8 encoding explicitly; the embedded version omits it (presumably because `mt.exe` stripped it as redundant — XML defaults to UTF-8 anyway). Same effect, fewer bytes in the compiled output. A minor optimization the manifest tool performs automatically.

### Line endings

The source file (as pasted) appears to use clean line endings without the trailing-whitespace artifacts seen in the binary form. The .res had `0x20 0x20` padding bytes at the end and CRLF throughout; the source XML is cleaner. This is normal — the manifest tool serializes the XML with its own conventions, not the source file's.

## What having the source XML tells us about the build workflow

Now that we have both the source XML and the compiled .res, we can infer the build pipeline:

1. The developer maintains the manifest as a hand-edited XML file (probably `StarHotel.exe.manifest` or `StarHotel.manifest` in the source tree)
2. As part of the build, `mt.exe` (or an equivalent tool) compiles the XML into a Win32 resource and writes it as `StarHotel.res`
3. The VB6 IDE picks up `StarHotel.res` via the .vbp's `ResFile32="StarHotel.res"` line
4. VB6 compiles the .res into the .exe alongside its own auto-generated resources (version block, form icons, etc.)

So the build has at least two preprocessing steps before VB6 even runs: the manifest compilation and probably icon embedding. The fact that `StarHotel.res` exists in the source tree (rather than being a build artifact) means the developer compiles it once when the manifest changes, then commits the .res and forgets about it. The .res becomes the "source of truth" for the build, even though it's derived from the .xml.

This is a common pattern for VB6 projects with manifests — VB6 itself can't author manifests, so the workflow involves an external manifest tool, and the .res is the bridge between that external tool and VB6's own build system.

## A few additional observations the source form exposes

- **The XML is well-formatted** — tab indentation, attributes ordered consistently, no extra whitespace. This suggests it was either authored in a real XML editor (not just Notepad) or run through a pretty-printer. Hand-typing this much XML this consistently would be error-prone.

- **No comments in the XML.** A more documented manifest might include `<!-- Vista -->`, `<!-- Windows 7 -->` etc. comments alongside the OS GUIDs to make them human-readable. Without comments, anyone maintaining this needs an external GUID-to-OS lookup. The lookup is well-documented (Microsoft publishes the canonical list), but it's a small barrier to maintainability.

- **The XML namespaces are explicit** (`urn:schemas-microsoft-com:asm.v1` for the root, `asm.v3` for trustInfo, `compatibility.v1` for the compatibility section). This is required syntax — Windows uses the namespaces to know which schema version each section conforms to. asm.v1 is the original 2001-era manifest schema; asm.v3 is the Vista/UAC-era extension; compatibility.v1 is the Windows 7+ extension. Three namespaces in one file, each from a different Windows era, all working together. Microsoft has been disciplined about backward compatibility here.

- **Attribute order is significant for Windows but not for XML.** The order of attributes inside `<assemblyIdentity>` is `name`, `version`, `type`, `processorArchitecture` — which is the canonical Microsoft order seen in their own samples. If a developer reordered them (e.g., putting `type` before `name`), the manifest would still validate as XML but might trigger different behavior in some Windows code paths that key off attribute position. The source preserves the canonical order.

- **`processorArchitecture="x86"` on the dependency** matches the `processorArchitecture="x86"` on the assembly identity. Both must agree — a 32-bit app must depend on the 32-bit version of Common Controls. If you ever ported this app to 64-bit (which would require leaving VB6 entirely, since VB6 only produces 32-bit code), you'd need to change both occurrences to `amd64`.

- **The `language="*"` wildcard in the dependency** means "any language pack of Common Controls is acceptable." The app doesn't ship a language-specific version; it accepts whatever's installed on the target Windows. This is the right call for an app that doesn't render localized strings from system resources — the controls themselves come pre-localized by the OS install language.

- **No conditional sections, no compile-time substitutions.** The XML is static. If "Star Hotel" needs to become "Sunset Resort" for a different customer, the developer hand-edits this file and rebuilds. There's no template variable like `{{CUSTOMER_NAME}}` that gets filled in from a config. This matches the pattern observed in the .vbp — customer branding is per-build, manually maintained.

## Notable points and quirks

- **The Markdown-corruption issue confirms a transcription artifact, not a source corruption.** As with `modMain.bas`, the actual file on disk is clean — what you see in the paste has been modified by some intermediate tool that auto-detected `Microsoft.Windows` as link-like and wrapped it. If you ever export this file through that same pipeline again, expect the same corruption.

- **Having both .xml and .res in the project means the manifest can be re-edited.** If the developer later wants to add a new supported OS GUID (say, when Microsoft eventually publishes one for Windows 12), they edit this XML, run `mt.exe` to regenerate `StarHotel.res`, then rebuild the VB6 project. The .res alone wouldn't be editable in any practical way; the .xml is the maintenance interface.

- **The XML is small enough to fit on one screen** (~30 lines, ~20 if you collapse whitespace). Easy to audit, easy to modify. No `<schemaLocation>` declarations, no namespace abbreviations beyond what's required. Clean and direct.

- **Comparing this to a typical .NET app's manifest:** .NET apps often have manifest sections for `<assemblyBinding>` (assembly version redirects), `<runtime>` (GC mode, configuration), and various publisher-policy sections. This Win32 app's manifest is much simpler — it doesn't need any of that, since there's no managed runtime to configure. The only thing the manifest does for this app is enable themed controls and declare OS support.

- **The `description` element is "Star Hotel"** — duplicated from the assembly identity's name. Different shells can show description vs name differently (the description sometimes shows as a tooltip or in property dialogs). For this app they're identical, so both surfaces show the same string.

- **The file name matters at runtime.** Windows looks for an external manifest at `<exename>.manifest` (e.g., `StarHotel.exe.manifest`) in the same directory as the .exe. If found, it overrides the embedded one. So a deployment could potentially include an external .manifest file to override the embedded settings without rebuilding the .exe — useful for last-minute compatibility tweaks. The build workflow here uses embedded manifests (compiled into .res), which is the safer choice — no risk of a missing or modified external file.

- **The `mt.exe` toolchain dependency is invisible from the source.** Nothing in the .vbp or this XML names the tool used to compile the .xml into the .res. So a developer cloning this project would need to know — out of band — that the .res was generated from this .xml using `mt.exe` (or equivalent), and would need that tool installed to regenerate the .res after editing the manifest. This is the kind of tribal knowledge that lives outside source control but is essential for project maintenance.

- **A modernization pass would consider adding:**
  - `<windowsSettings><dpiAware>True/PM</dpiAware></windowsSettings>` for high-DPI awareness (improves rendering on 4K monitors)
  - `<windowsSettings><activeCodePage>UTF-8</activeCodePage></windowsSettings>` for UTF-8 file path support (Win10 1903+ feature)
  - `<windowsSettings><longPathAware>true</longPathAware></windowsSettings>` to lift the 260-character path limit
  - Win11-specific `<supportedOS>` GUID once Microsoft publishes one
  
  None of these would change the application's behavior dramatically, but each would improve the experience on modern Windows. They're all opt-in — the absence is forward-compatible, just suboptimal.

- **The presence of asm.v3 and compatibility.v1 namespaces in addition to asm.v1** demonstrates that this manifest was authored (or last edited) by someone aware of post-Vista Windows requirements. asm.v3 is the namespace UAC introduced; compatibility.v1 is the namespace Windows 7 introduced for the supportedOS list. A pre-Vista manifest would use only asm.v1. So this manifest was authored or revised in the Vista-or-later era — consistent with the project's 2014-2021 timeline.

- **`uiAccess="false"`** is the default and could have been omitted. Including it explicitly is a minor act of documentation — making it clear that the developer considered UI accessibility privileges and chose not to request them. uiAccess=true is for tools that need to bypass UAC's input integrity isolation (screen readers, remote-assistance helpers) and would require code-signing the .exe. For a normal app, false is correct.

- **No `<fileAssociation>` or `<protocolHandler>` registrations** in the manifest. So the app doesn't claim to handle any particular file extension or URL protocol — double-clicking a `.mdb` file won't launch StarHotel.exe (that's correctly Microsoft Access's job). The app is launched only by direct user action on its own .exe.

- **The manifest is the same regardless of customer.** If this codebase produces builds for multiple customers (Star Hotel, Sunset Resort, Mountain Lodge), each would presumably have its own manifest with the customer name in `name` and `description`. So this XML is one of multiple per-customer files — alongside per-customer .vbp, per-customer .res, per-customer compiled .exe. The build process for a new customer involves at minimum: clone the .xml, change "Star Hotel" to the new name in three places (assembly identity name, description, and the version line if rev-numbered separately), regenerate the .res, update the .vbp's branding fields, and recompile. Five or six places to coordinate per customer, no automation visible — pure manual ISV workflow.

- **The version 1.2.0.22 in this file** matches the binary version we saw in the .res-extracted version. Both differ from the .vbp's 1.2.22 format. If you ever rev the version, you have to update it in: this XML (`assemblyIdentity version=`), the .vbp's `MajorVer`/`MinorVer`/`RevisionVer`, and (depending on workflow) any installer scripts and release notes. Three or four places to coordinate per release. Not unusual for VB6-era projects, but not great hygiene either.

In summary, the source XML adds two things on top of what we already knew from the binary .res: it confirms the build workflow (hand-edited XML compiled to .res by an external tool), and it reveals (again) the Markdown-link-injection corruption pattern affecting the paste pipeline. The actual content of the manifest is identical to what we analyzed before — same dependencies, same trust level, same supported-OS list. The XML is the human interface; the .res is the binary interface; they're two views of the same statement.
