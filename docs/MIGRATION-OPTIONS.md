# Migration & Port Options

Five plausible routes for taking the original `StarHotel` VB6 codebase under `ORIGINAL-CODE/` to a maintainable modern system, with pros/cons, phased roadmaps, and concrete tooling recommendations for each. The aim of this doc is **route selection** — picking the path that gets followed up with deeper per-route design docs (e.g., `MIGRATION-DOTNET.md`, `MIGRATION-PYTHON.md`).

The original is sized to make all five routes feasible: ~2,900 lines of module code across 7 `.bas` files, ~10–11K lines of actual code in 16 `.frm` files (the rest is form-designer markup), 12 database tables, 7–9 Crystal Reports, and a handful of distinct business workflows. That puts the total port surface area at roughly **13K lines of behavioral VB6 code**, plus database schema, plus reports — a small-to-medium codebase.

For everything below, severity context comes from `BUGS-MITIGATIONS.md` — there are 150 catalogued issues, 4 Critical and ~10 High. The migration is also an opportunity to retire those, regardless of target.

## Decision dimensions

Before evaluating routes, decide which of these matter most. The right route depends almost entirely on the answers.

| Dimension | Question |
|---|---|
| **UX paradigm** | Stay desktop kiosk on Windows, or shift to web/mobile/cross-platform? |
| **Team language** | What does the team already write? What can they hire for? |
| **Deployment target** | Single Windows machine per hotel, or central server with thin clients, or hosted SaaS? |
| **Reporting fidelity** | Do existing `.rpt` Crystal Reports need pixel-perfect replication, or can they be redesigned? |
| **Migration risk tolerance** | Big-bang cutover, parallel run, or strangler-fig phased replacement? |
| **Budget** | Commercial migration tooling (Mobilize VBUC etc. — five to six figures) on the table, or open-source/in-house only? |
| **Timeline pressure** | Crystal Reports 2020 mainstream support ends Dec 2026; CR 32-bit runtime ended Dec 2025. Is that a hard deadline? |

## What every route must replace

Independent of language target, these decisions need answers:

- **The database.** Microsoft Access `.mdb` (Jet 4.0) is being deprecated by Microsoft and is 32-bit only. Realistic targets: **SQLite** (single-file, drop-in), **SQL Server LocalDB / Express** (Microsoft-aligned, free for small deployments), or **PostgreSQL** (full RDBMS, needs a server). Migration of the 12 tables is mechanical; the schema in `Module/modDatabase.bas:CreateDB` is the source of truth.
- **The report engine.** Crystal Reports 8.5 was never ported to .NET Core / 8+ and is at end-of-life. The 7–9 `.rpt` files cannot be transparently re-targeted; they have to be rebuilt in a new tool. Options vary by route (see each section).
- **The auth/security model.** `GoldFishEncode` must be replaced with bcrypt/argon2id (see `BUGS-MITIGATIONS.md` §1). Existing hashes can't be transparently upgraded — users get re-hashed on next login.
- **Idle-watchdog and Void features.** Currently inert in the original (debug overrides). Either fix and forward-port, or design out.
- **Configuration and storage paths.** `App.Path`-relative `Config.txt`/`Error.txt` writes are broken under Program Files installs. New target should use `%APPDATA%`-equivalent.

These can be done in any order, but the database swap is usually first because every route depends on it.

---

## Route A — Incremental VB6 → VB.NET → C#/.NET 10

**The "bridge" path.** Use a commercial migration tool to translate VB6 to VB.NET (targeting .NET Framework 4.8 first, then forward to .NET 10 LTS via the .NET Upgrade Assistant). Optionally migrate VB.NET to C# afterwards. The path Microsoft and the tooling ecosystem have most explicitly built for.

### Pros

- **Smallest semantic gap.** VB6 forms map onto VB.NET WinForms with fairly direct correspondence — event handlers, controls, and the global-state pattern all survive. Mobilize.Net VBUC (now GAPVelocity AI Migrator) and VB Migration Partner claim 40–70% automation; in practice a small codebase like this could land closer to 80% with cleanup.
- **Most mature tooling.** VBUC has been a production tool since the mid-2000s, recently augmented with "hybrid AI" passes that produce more idiomatic .NET. VS 2026 supports `.slnx` solutions and target frameworks up to .NET 10. The .NET Upgrade Assistant handles VB.NET → .NET 8/9/10 forward migration automatically.
- **Crystal Reports has the most plausible bridge.** SAP Crystal Reports for Visual Studio still exists for .NET Framework 4.8; you can keep the existing `.rpt` files running while you replace them piece by piece. Drop-in replacements (Telerik Reporting, DevExpress XtraReports, ActiveReports, FastReport.NET) all have CR import or CR-style designers.
- **Microsoft ecosystem continuity.** ADO 2.8 → ADO.NET → Entity Framework Core is well-trodden. Win32 manifest, COM controls (MSCOMCTL), and Common Controls all have direct .NET equivalents.
- **Hiring pool.** Far more .NET developers than VB6 developers; far more .NET than any single non-Microsoft alternative.
- **Can be incremental.** With sufficient COM interop, parts of the app can run as VB6 while others run as .NET. Enables a strangler-fig style replacement instead of big-bang cutover.

### Cons

- **You inherit VB.NET, not VB6.** Migration tools produce VB.NET. Idiomatic C# requires a second translation pass. Many teams stop at VB.NET and live with it forever — including its quirks (`On Error GoTo`, default properties, late binding).
- **Migrated code is rarely idiomatic.** VBUC output is recognizably translated, not native. Globals stay global; event-handler-heavy forms stay event-handler-heavy. The 150-bug list survives the translation; you fix them after, not during.
- **Commercial tooling cost.** VBUC and VB Migration Partner are five-to-six-figure licenses for production migrations. Open-source alternatives (the bundled VS 2008 Upgrade Wizard) produce notoriously broken output and are not recommended.
- **WinForms is the realistic UI target.** Migration tools target WinForms, not WPF/MAUI/Avalonia. WinForms is fine but is itself a 2002-era framework — modern would be a third hop.
- **The architectural smells survive.** `gstrSQL` global mutation, hand-rolled SQL builder, hub-and-spoke form chaining — translation preserves all of it. A *clean* port would discard them; a *migrated* port keeps them.
- **Crystal Reports for VS is itself end-of-life-flavored.** It works but isn't being modernized. You eventually have to replace it anyway.

### Roadmap

| Phase | Work | Approx scale |
|---|---|---|
| 0 | Documentation & feature freeze on `ORIGINAL-CODE/` (already done) | — |
| 1 | Schema export → SQLite or SQL Server LocalDB; data migration script. Re-test original app against new DB via OLE DB driver if possible. | 1–2 weeks |
| 2 | Run VBUC / VB Migration Partner / GAPVelocity AI Migrator against the codebase. Target VB.NET on .NET Framework 4.8. Get it building. | 2–4 weeks |
| 3 | Manual cleanup: replace COM dependencies (ADO, MSCOMCTL) with .NET equivalents; fix what the translator couldn't. | 4–8 weeks |
| 4 | Replace Crystal Reports 8.5 with a .NET-Framework-compatible engine (CR for VS as bridge, or direct to Telerik/DevExpress). Re-author 7–9 reports. | 2–4 weeks |
| 5 | `.NET Upgrade Assistant`: VB.NET on .NET Framework 4.8 → .NET 8/10. Switch from `.frm` resources to modern WinForms designer. | 1–2 weeks |
| 6 | Address `BUGS-MITIGATIONS.md` Critical/High items (auth rewrite, parameterized queries, etc.) | 4–6 weeks |
| 7 (optional) | VB.NET → C# (manual or with tools like ICSharpCode.CodeConverter). | 2–4 weeks |
| 8 (optional) | WinForms → WPF or Avalonia for cross-platform. | 4–8 weeks |

**Total elapsed: 4–8 months for a single experienced developer through Phase 6; Phases 7–8 are extensions if the team wants them.**

### Tooling

- **Migrator:** GAPVelocity AI Migrator (formerly Mobilize.Net VBUC) **or** VB Migration Partner — pick one based on trial output quality.
- **IDE:** Visual Studio 2026 (handles VB6 import, .NET migration, and modern target frameworks in one tool).
- **Forward migration:** Microsoft `.NET Upgrade Assistant` CLI/extension.
- **Reporting (bridge):** SAP Crystal Reports for Visual Studio (.NET Framework only).
- **Reporting (final):** Telerik Reporting, DevExpress XtraReports, GrapeCity ActiveReports, or **FastReport.NET** (the cheapest commercial option) / **QuestPDF** (open-source code-first, no designer).
- **Database:** SQLite via `Microsoft.Data.Sqlite`, or SQL Server LocalDB via `Microsoft.Data.SqlClient`.
- **ORM (optional):** Entity Framework Core 8/10, or Dapper for thin mapping.
- **Auth:** `BCrypt.Net-Next` or `Konscious.Security.Cryptography.Argon2`.

---

## Route B — Greenfield C# / .NET 10 rewrite

**The "clean slate" .NET path.** Skip the auto-migration. Use the docs in `docs/forms/` and `docs/modules/` as a specification, and build a new C# / .NET 10 app from scratch — modern WinForms, WPF, or Avalonia. The original VB6 stays as a reference, never as input to a translator.

### Pros

- **No translated-code debt.** The output is idiomatic modern .NET from day one. Globals get replaced with DI; event-handler spaghetti gets replaced with MVVM or MVU; `gstrSQL` doesn't exist.
- **Same .NET ecosystem benefits as Route A** — best-in-class libraries, hiring pool, tooling, Microsoft support.
- **Free choice of UI framework.** WinForms (closest to original UX), WPF (modern Windows desktop, mature), Avalonia (cross-platform, MIT-licensed, growing fast), or .NET MAUI (cross-platform but more app-store-oriented). Not constrained by what migration tools target.
- **150 known bugs designed out, not fixed in.** The Critical/High issues from `BUGS-MITIGATIONS.md` simply never appear — you build with parameterized queries, KDF-based auth, FK constraints, and per-operation transactions from the start.
- **Cleaner Crystal Reports replacement.** No CR-for-VS bridge — go directly to a modern engine.

### Cons

- **Bigger upfront cost.** Phase 1 produces nothing usable; you're rebuilding everything. The auto-migration's "running app in 2 weeks" milestone doesn't exist here. Realistically 3–6 months before there's a usable replacement.
- **Specification quality determines fidelity.** The `docs/` analyses are excellent for a small system but won't catch every edge case. Expect "we never realized the form did X" bugs during the first month of UAT.
- **No incremental option.** Can't run new + old in parallel without significant adapter work, so it's a big-bang cutover at the end.
- **Higher discipline burden.** A fresh codebase invites scope creep. "While we're rewriting, let's also add multi-tenancy / mobile app / API" — those discussions are easier to have but harder to bound.
- **Still need to migrate user data.** A one-shot script at cutover, not an incremental concern, but the data does need to survive.

### Roadmap

| Phase | Work |
|---|---|
| 0 | Treat `docs/` as the spec. Identify gaps; back-fill via source reading. |
| 1 | Domain model + repository layer + SQLite migration of schema. CRUD for every entity; no UI yet. |
| 2 | Auth/login: bcrypt, parameterized queries, session abstraction. |
| 3 | Report engine integration — pick one (QuestPDF for code-first PDF; Telerik/DevExpress for designer-driven), build the 7–9 reports. |
| 4 | UI: build forms. Recommended order: Splash → Login → Dashboard → Booking → Find Customer → Reports → Admin tools. |
| 5 | Idle watchdog, audit trail, error logging — features that were broken or missing in original. |
| 6 | UAT against the original app on parallel data. |
| 7 | Data migration script + cutover. |

**Total elapsed: 5–9 months for a single experienced developer; 3–5 months for a 2-person team.**

### Tooling

- **Language/runtime:** C# / .NET 10 LTS (supported through Nov 2028).
- **UI framework recommendation:** **Avalonia 11+** if cross-platform matters; **WPF** if Windows-only and team has WPF experience; **WinForms** if you want maximum proximity to the original UX.
- **Reporting:** **QuestPDF** (open-source, MIT, code-first — best for invoices/receipts) **plus** Telerik or DevExpress for ad-hoc/dashboard reports if needed.
- **Database:** SQLite via EF Core 10 with `Microsoft.EntityFrameworkCore.Sqlite`.
- **Auth:** `BCrypt.Net-Next` (simplest) or `Konscious.Security.Cryptography.Argon2` (best).
- **Logging:** Serilog with rolling-file sink (replaces both `LogErrorDB` and `LogErrorText`).
- **Testing:** xUnit + FluentAssertions + a small set of integration tests against an ephemeral SQLite.

---

## Route C — Python port (PySide6 + SQLite)

**The "simple stack" path.** Rewrite as a Python desktop app using PySide6 (LGPL Qt bindings), SQLite for storage, and ReportLab or WeasyPrint for reports. Cross-platform out of the box.

### Pros

- **Cross-platform desktop for free.** PySide6 / Qt 6 runs on Windows, macOS, and Linux from the same codebase. The original was Windows-only; this opens deployment options (Linux thin client kiosks, Mac admin terminals).
- **PySide6 is LGPL.** Free for commercial / proprietary use with no royalties — unlike PyQt6, which is GPL-only without a commercial license. (Critical distinction; got commercial Python desktop projects in trouble for years.)
- **Smallest dependency footprint of the cross-platform options.** A bundled PySide6 app is in the 50–80MB range with PyInstaller — bigger than native but smaller than Electron.
- **Excellent rapid-prototyping.** Python's REPL + Qt Designer + signals/slots makes form-driven development fast. The hub-and-spoke pattern of the original maps cleanly onto Qt windows.
- **Mature reporting alternatives.** ReportLab (open-source, code-first PDF generation, very capable) and WeasyPrint (HTML/CSS → PDF) are both production-grade.
- **Solid ORMs.** SQLAlchemy 2.x (with type hints) is the gold-standard Python ORM; or stick with raw `sqlite3` for a small app.
- **The ecosystem is excellent for adjacent needs** — CSV/Excel import/export, charting, data-science integrations if the hotel ever wants analytics.

### Cons

- **No automated migration.** Every line is hand-rewritten. Slowest start of any route once you account for the lack of tooling.
- **Distribution friction on Windows.** PyInstaller / Briefcase / cx_Freeze produce single-folder or single-EXE bundles, but they're not as clean as a native .NET deployment. Antivirus false-positives on PyInstaller bundles are a recurring annoyance.
- **Smaller hiring pool for Python desktop specifically.** Python has plenty of developers, but most do web/data — Qt experience is rarer than .NET WinForms experience.
- **No bridge to Crystal Reports.** Existing `.rpt` files can't run; reports must be redesigned from scratch.
- **Performance ceiling lower than .NET/Java.** Not relevant for a hotel POS, but worth noting if requirements change.
- **Python packaging is a perpetual headache** — venvs, requirements pinning, Python version compatibility on customer machines. Mitigated by bundling, but mitigated, not eliminated.

### Roadmap

| Phase | Work |
|---|---|
| 0 | Docs as spec; pin Python 3.12 LTS; stand up `uv` + project skeleton. |
| 1 | SQLite schema migration (Alembic or raw); domain model with SQLAlchemy. |
| 2 | Auth + session, with `passlib` (bcrypt). |
| 3 | Report rebuild in ReportLab — receipts, daily/weekly/monthly bookings, customer history. |
| 4 | UI in PySide6 — forms in Qt Designer, save as `.ui`, load via `uic` or compile to `.py`. Build hub-and-spoke navigation. |
| 5 | Bug-list cleanup designed in (parameterized queries via SQLAlchemy; argon2/bcrypt auth; transactions per business operation). |
| 6 | Bundle via PyInstaller or Briefcase; smoke-test on clean Windows + Linux VMs. |
| 7 | Data migration; cutover. |

**Total elapsed: 6–10 months for a single developer.** Slower than Route B because there's no migration tool, but faster than Route D (Java) because of language tractability.

### Tooling

- **Language/runtime:** Python 3.12+.
- **UI:** **PySide6** (NOT PyQt6 — license trap). Use Qt Designer for layouts.
- **Database:** SQLite via SQLAlchemy 2.x or raw `sqlite3`.
- **Reporting:** **ReportLab** for code-first PDFs (receipts/invoices); **WeasyPrint** if you'd rather author in HTML/CSS.
- **Auth:** `passlib[argon2]` or `argon2-cffi`.
- **Bundling:** PyInstaller (most common) or Briefcase (more modern, BeeWare project).
- **Testing:** pytest + pytest-qt for UI tests.
- **Project mgmt:** `uv` (fast) or `poetry` (mature).

---

## Route D — Java port (JavaFX + JasperReports)

**The "enterprise-grade" path.** Rewrite as a Java desktop app using JavaFX for UI and JasperReports for reporting. Strong typing, robust ecosystem, predictable runtime.

### Pros

- **JasperReports is the strongest Crystal-Reports-style replacement.** Mature (since 2001), open-source, designer-driven (Jaspersoft Studio), and idiomatic for the kind of pixel-perfect reports the original uses. Direct conceptual port of the `.rpt` workflow.
- **Strong, statically-typed language.** Catches more bugs at compile time than Python or VB.NET. The 150-issue bug list from `BUGS-MITIGATIONS.md` includes several runtime-only failures that wouldn't compile in Java.
- **Cross-platform via the JVM.** Same bytecode runs on Windows, macOS, Linux. JavaFX is mature and bundled with JDKs since 17.
- **Excellent enterprise libraries.** Spring (overkill for this) or just plain JDBC + HikariCP. Hibernate or jOOQ for ORM. Mature transaction support, connection pooling, logging (SLF4J + Logback).
- **Long-term support story.** Java 21 LTS supported through 2031; Java 25 LTS through 2033. .NET LTS only goes 3 years.
- **JasperReports can keep the report definition style alive.** The `.jrxml` format is designer-driven (in Jaspersoft Studio) and conceptually similar to `.rpt`. This means the existing report knowledge ports relatively cleanly.

### Cons

- **JavaFX UX feels less native than .NET WinForms or PySide6.** Looks polished but slightly off-brand on Windows compared to a native app — fine for an internal hotel POS, less fine if branding matters.
- **No automated migration.** Like Python and TypeScript, fully hand-rewritten.
- **Heavier runtime.** A JavaFX app distributed with `jpackage` is 60–120MB. Not bad, but bigger than .NET self-contained or PySide6 bundled.
- **Java development feels dated to many teams.** Modern Java (records, pattern matching, virtual threads in 21+) is much better than 2010-era Java, but the perception lags.
- **Smaller talent pool for Java desktop specifically.** Most Java work today is server-side; JavaFX is a niche.
- **Build tooling has more moving parts** — Maven or Gradle, JLink/jpackage for distribution, module system if used. More to set up than Python or .NET.
- **JasperReports licensing nuance.** Library is LGPL; Jaspersoft Studio (the designer) is GPL/AGPL with a commercial track. Usable for proprietary apps if you stick to the LGPL pieces, but read carefully.

### Roadmap

| Phase | Work |
|---|---|
| 0 | Docs as spec; pin JDK 21 or 25 LTS; Maven or Gradle skeleton. |
| 1 | SQLite schema migration (Flyway or Liquibase); domain model. |
| 2 | Auth + session with `bcrypt-java` or `argon2-jvm`. |
| 3 | Report rebuild in **Jaspersoft Studio** as `.jrxml`. Same workflow as Crystal Reports designers — should feel familiar. |
| 4 | UI in JavaFX with FXML (designer-friendly) loaded via `FXMLLoader`. Hub-and-spoke navigation. |
| 5 | Bug-list designed-in, same as Route C. |
| 6 | Distribute via `jpackage` for Windows/macOS/Linux installers. |
| 7 | Data migration; cutover. |

**Total elapsed: 7–11 months for a single developer.** Slowest of the rewrites because of the verbosity tax and ecosystem scaffolding.

### Tooling

- **Language/runtime:** Java 21 or 25 LTS via Adoptium / Temurin.
- **UI:** **JavaFX 21+** with FXML and Scene Builder (designer).
- **Database:** SQLite via `sqlite-jdbc`; HikariCP for pooling; jOOQ or Hibernate for ORM (or plain JDBC).
- **Reporting:** **JasperReports** (the library) + **Jaspersoft Studio** (the designer) for the `.rpt` replacement.
- **Auth:** `bcrypt` (`org.mindrot:jbcrypt`) or `argon2-jvm`.
- **Logging:** SLF4J + Logback.
- **Build/distribute:** Gradle + `jpackage` (JDK-bundled) for native installers.
- **Testing:** JUnit 5 + TestFX for UI tests.

---

## Route E — TypeScript port (Tauri 2 + SQLite, or full web app)

**The "maximum modernization" path.** Rewrite the front-end in TypeScript (React/Vue/Svelte), back it with SQLite, and deliver as either a Tauri 2 desktop app or a hosted web app. Biggest paradigm shift; biggest range of future options.

### Pros

- **Best long-term flexibility.** A React + TS app delivered today as a Tauri desktop bundle can be redeployed as a hosted SaaS, a PWA, an Electron app, or a mobile app (Tauri 2 supports iOS/Android) with mostly UI tweaks. No other route gives you this.
- **Tauri 2 is production-ready in 2026 and dramatically better than Electron** — ~12MB bundles vs Electron's 180MB, ~85MB RAM vs 450MB, native WebView2 on Windows. The "Electron is bloated" objection no longer applies if you choose Tauri.
- **Largest hiring pool of any route.** TypeScript + React/Vue is the most common skill stack in the industry. Hiring is easier than for any other route.
- **Best library ecosystem for adjacent features.** Charts, file import/export, dashboards, real-time updates, multi-tenant — all out of the box. Hotels eventually want these even if they don't ask for them now.
- **Reporting via PDF generation has matured.** Server-side PDF (Puppeteer + headless Chrome, or Carbone, or PDFKit) gives pixel-perfect output from HTML/CSS templates — much faster to author than `.rpt` or `.jrxml`.
- **Web option opens up centralized deployment.** One server, multiple hotel front-desk terminals, browser-only — eliminates per-machine install/upgrade pain. Especially valuable if a chain operates multiple properties.
- **Type system catches bugs auto-translation can't.** TypeScript's strict mode + a real linter (Biome / ESLint) prevents whole categories of issues.

### Cons

- **Biggest paradigm shift of any route.** Forms-and-events VB6 doesn't map cleanly onto React's component/state model. The translation is conceptual, not syntactic. Higher risk of misunderstanding the original.
- **Two-codebase reality.** Even with Tauri, you have a TS frontend and a Rust backend (Tauri commands) — or alternatively a Node backend if you go full Electron. More moving parts than the others.
- **No automated migration tooling whatsoever.** Hand-rewritten everything, like Python and Java.
- **Reporting is the loosest fit.** Generating pixel-perfect PDFs from HTML/CSS is doable but requires more design work than a designer-driven tool like JasperReports or Telerik. Carbone (Word-template-based) and similar tools narrow this gap.
- **Offline/local-first story is more complex.** A web app needs PWA + service workers + local SQLite (sql.js or wa-sqlite) to work offline; a Tauri app handles this natively but locks you in to Rust for backend work.
- **Single-machine hotel POS is the original design point — moving to web is a real product decision** that affects training, hardware, network reliability, and support.
- **Browser security model.** Front-desk staff using a web app inherit all the considerations of web security; the original avoided most of these by being a local app.

### Roadmap

| Phase | Work |
|---|---|
| 0 | Docs as spec. Decide deployment model: Tauri 2 desktop app, or hosted web app. |
| 1 | Backend: Node + TypeScript + SQLite via `better-sqlite3` (or Rust + `rusqlite` if Tauri). Schema migration. Domain model. Repository layer. |
| 2 | Auth: argon2 (`@node-rs/argon2`) + JWT or session cookies. |
| 3 | Report rebuild — HTML/CSS templates rendered via Puppeteer for PDF, or Carbone for Word-template-based reports. |
| 4 | Frontend: pick stack (recommended: **React 19 + Vite + TanStack Router + TanStack Query + shadcn/ui**, or Svelte 5 if team prefers). Build forms component-by-component. |
| 5 | Tauri 2 wrapping (if desktop) — `tauri.conf.json`, command registration, native menus. |
| 6 | Bug-list designed-in. |
| 7 | Distribution: Tauri builds MSI/DMG/AppImage; or deploy web app behind reverse proxy. |
| 8 | Data migration; cutover. |

**Total elapsed: 7–12 months for a single developer; 4–6 months for a 2-person team.** Highest variance because the design decisions (Tauri vs web, React vs Svelte, ORM vs raw SQL, JWT vs session) are all open.

### Tooling

- **Language/runtime:** TypeScript 5.5+; Node 22 LTS (if web/Electron) or Rust stable (if Tauri).
- **Desktop wrapper:** **Tauri 2** (default to this over Electron in 2026 unless ecosystem reasons demand Electron).
- **Frontend:** React 19 + Vite + TanStack Router + TanStack Query + shadcn/ui (Radix + Tailwind). Or Svelte 5 + SvelteKit if SSR/web matters.
- **Database:** SQLite via `better-sqlite3` (Node) or `rusqlite` (Rust/Tauri). Drizzle ORM or Kysely for typed queries.
- **Reporting:** **Carbone** (Word templates → PDF) or **Puppeteer + headless Chrome** for HTML→PDF, or **PDFKit** for code-first.
- **Auth:** `@node-rs/argon2` or `bcrypt`.
- **Testing:** Vitest + Playwright for end-to-end.

---

## Side-by-side comparison

| Dimension | Route A (Migrate to .NET) | Route B (Greenfield .NET) | Route C (Python) | Route D (Java) | Route E (TypeScript) |
|---|---|---|---|---|---|
| **Automation potential** | High (40–80%) | None | None | None | None |
| **Total elapsed (1 dev)** | 4–8 months | 5–9 months | 6–10 months | 7–11 months | 7–12 months |
| **Code idiomaticity at finish** | Low–Medium | High | High | High | High |
| **Crystal Reports bridge** | Yes (CR-for-VS) | No | No | No (but Jasper is a clean conceptual port) | No |
| **Cross-platform** | No (WinForms) / Yes (Avalonia) | Free choice | Yes | Yes | Yes |
| **Mobile reach** | No | No | Limited | No | Yes (Tauri 2 / web) |
| **Hiring pool** | Large | Large | Large (web) / Small (desktop) | Medium | Largest |
| **Tool licensing cost** | $$$ (commercial migrator) | Free | Free | Free | Free |
| **Designer-driven reports** | Yes | Yes (commercial only) | No (code-first) | Yes (Jasper) | Limited |
| **Bridges from existing data** | OLE DB | One-shot script | One-shot script | One-shot script | One-shot script |
| **Risk of "we missed an edge case"** | Low (mechanical) | Medium | Medium | Medium | Medium–High |
| **Future-flexibility ceiling** | Medium (desktop-bound) | Medium | Medium | Medium | High (web/mobile/desktop) |

---

## Hybrid options (briefly)

These don't replace a primary route but can layer on top of one:

- **Strangler fig.** Keep the VB6 app running. Build a new API layer (any of the above stacks) against the same `.mdb`. Migrate one form at a time to the new system. Hardest path operationally but lowest cutover risk. Realistic only if the new stack is .NET (Route A or B) — COM interop with VB6 is otherwise painful.
- **Reports-first.** Replace just the reporting engine first (any modern tool) while leaving the rest of the VB6 app untouched. Buys time on the Crystal Reports EOL clock and de-risks the biggest single component of any future port.
- **Database-first.** Migrate `.mdb` to SQLite or SQL Server LocalDB first, point the existing VB6 app at the new DB via OLE DB. Same effect — buys time, removes one EOL risk, doesn't change anything user-facing.
- **Bridge mode.** Run the VB6 app inside a Windows VM or container indefinitely while the rewrite proceeds in parallel. Acceptable for hotels that can stay on a known-good build for 12 months.

The reports-first or database-first hybrids are both **highly recommended preludes** to any primary route, because they de-risk the hardest pieces independently.

---

## Recommendation

**Pick Route B (greenfield C# / .NET 10 with Avalonia + QuestPDF + SQLite).** Here's why:

1. **The codebase is small enough that auto-migration isn't worth the licensing cost.** ~13K lines of behavioral code is roughly 4–8 weeks of clean-rewrite work for a competent .NET developer using the existing `docs/` as a spec. Route A's auto-migration only saves time on codebases ~10× this size.

2. **The 150 catalogued bugs are easier to *avoid* than to *fix*.** Route A preserves the architectural smells (`gstrSQL` global, naive sanitization, no transactions, weak hash). Routes C/D/E avoid them too, but Route B does so with the lowest paradigm shift and best tooling fit.

3. **.NET 10 LTS is the longest realistic support window.** Supported through Nov 2028. By then the next LTS exists. Java LTS is longer (Java 21 → 2031) but the desktop story is weaker.

4. **Avalonia gives cross-platform without locking you in.** Same C# code targets Windows / macOS / Linux. If the hotel ever wants a Mac admin workstation or a Linux thin client, no rewrite needed. WinForms or WPF can replace it if the team prefers Microsoft-native — the choice can come a few weeks in.

5. **QuestPDF is a code-first PDF library that fits a small report set well.** The original has only 7–9 reports. They're not so complex that they need a designer-driven tool. Code-first means the report definitions live in the same Git repo as the rest of the app, are reviewable in PRs, and don't need a separate tool license.

6. **Hiring and contracting are easy.** .NET developers are plentiful; C# specifically is what most modernization shops default to.

**The strongest alternative is Route E (TypeScript + Tauri 2)** if the long-term plan involves multiple hotels / multi-tenant / mobile. The flexibility ceiling is materially higher, at the cost of a meaningfully bigger paradigm shift up front.

**Avoid Route A** unless the codebase grows substantially before migration starts (e.g., other VB6 apps from the same shop are added to the scope), at which point the auto-migration savings start to matter.

**Avoid Routes C and D** unless the team's existing skillset strongly favors them. Both produce great results, but neither has the ecosystem advantages over Route B that would justify choosing them on technical grounds for a Windows-first hotel POS.

## First steps regardless of route

These four pieces of work are valuable under every route. Doing them now de-risks the route choice itself:

1. **Export the schema and sample data to SQLite.** Read `Module/modDatabase.bas:CreateDB` and `CreateSampleData`; write a one-shot script that recreates everything in SQLite. This costs ~1 day and makes routes B/C/D/E all start with a working DB.

2. **Pick the new report engine and rebuild *one* report end-to-end.** Receipt is the smallest. Doing this once tells you whether the chosen stack handles your most frequent report. Costs ~3–5 days.

3. **Triage `BUGS-MITIGATIONS.md`.** Mark which findings are blocking-for-rewrite (Critical/High that change the data model or auth), nice-to-fix (Medium), and out-of-scope (Low cosmetic). This becomes the rewrite's acceptance criteria.

4. **Decide UX paradigm: stay desktop, or go web.** Everything else flows from this. Don't start writing code before this is decided.

---

## Roadmap visualization

```
          ┌── Route A: VB6 → VB.NET (.NET Fx 4.8) → .NET 10 → C# → Avalonia
          │
          ├── Route B: ★ Greenfield C# / .NET 10 + Avalonia + QuestPDF + SQLite  [RECOMMENDED]
          │
VB6 ──────┼── Route C: Python 3.12 + PySide6 + ReportLab + SQLite
(today)   │
          ├── Route D: Java 21+ + JavaFX + JasperReports + SQLite
          │
          └── Route E: TypeScript + React + Tauri 2 + Carbone + SQLite
                       (or hosted web variant)

Common preludes (any route):
  ├── Schema → SQLite / SQL Server LocalDB
  ├── Reports-first or DB-first strangler
  └── Triage BUGS-MITIGATIONS.md to define acceptance criteria
```

## Sources

The 2026-state research informing this doc:

- [VB6 to .NET in 2026 — GAPVelocity AI](https://www.gapvelocity.ai/blog/vb6-to-dotnet)
- [VB Migration Partner — feature comparison](https://www.vbmigration.com/whitepapers/featurecomparison.aspx)
- [Convert VB6 to .NET: Migration Strategy for 2026 — Legacy Leap](https://www.legacyleap.ai/blog/gen-ai-powered-vb6-to-dot-net-modernization/)
- [Crystal Reports Alternative for .NET Core — Dotnet Report (2026)](https://dotnetreport.com/blog/crystal-reports-alternative-dotnet-core)
- [5 Best Crystal Reports Alternatives for 2026 (End of Life Guide) — BlazeSQL](https://www.blazesql.com/blog/crystal-reports-alternatives)
- [SAP Crystal Reports for Visual Studio (.NET) — SAP Community](https://pages.community.sap.com/topics/crystal-reports/visual-studio)
- [Tauri vs Electron in 2026 — Rustify](https://rustify.rs/articles/rust-tauri-vs-electron-2026)
- [Best Desktop App Frameworks in 2026 — pkgpulse](https://www.pkgpulse.com/blog/best-desktop-app-frameworks-2026)
- [Which Python GUI library should you use in 2026? — Python GUIs](https://www.pythonguis.com/faq/which-python-gui-library/)
- [PyQt6 vs PySide6 licensing — Python GUIs](https://www.pythonguis.com/faq/pyqt6-vs-pyside6/)
