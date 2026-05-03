# vb6-rework-reverse-forward

> Reverse-engineering documentation and forward-port migration plan for **StarHotel**, a Visual Basic 6 hotel reservation system originally released 2014–2021.

**Current phase:** documentation and migration planning. The forward-port implementation has not started; this repository currently contains the upstream VB6 source as a frozen reference plus comprehensive analysis docs that will drive the rewrite.

---

## Contents

- [What this is](#what-this-is)
- [The original system](#the-original-system)
- [Project phases](#project-phases)
- [Documentation index](#documentation-index)
- [Repository layout](#repository-layout)
- [Migration approach](#migration-approach)
- [Reading order for newcomers](#reading-order-for-newcomers)
- [Development workflow](#development-workflow)
- [Status](#status)
- [License](#license)
- [Acknowledgments](#acknowledgments)

---

## What this is

A two-phase modernization effort:

1. **Reverse** — Read the original VB6 source line by line and produce long-form analysis documents covering every form, module, and project artifact. Catalog every bug, security flaw, half-implemented feature, and architectural smell. Done. Lives in [`docs/`](docs/).
2. **Forward** — Use those analyses as the specification for a clean-slate rewrite on a modern stack. Plan complete; implementation not yet started. Roadmap in [`docs/MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md).

The point of doing reverse first is that the original codebase has no design docs, no tests, and no maintainers — and 150+ catalogued issues including critical security flaws. A direct auto-translation would carry all of that forward. The reverse pass builds the understanding necessary to design those problems out.

## The original system

`StarHotel` is a single-machine, single-user-at-a-time Windows desktop hotel booking system covering room reservations, check-in/out, payments, customer history, reporting, and access control for a property of up to 55 rooms across four floors. Built by Aeric Poon (Computerise System Solutions), MIT-licensed, last touched in 2021. Upstream: [pyhoon/star-hotel-vb6](https://github.com/pyhoon/star-hotel-vb6).

Technology profile of the original:

- Visual Basic 6, compiled to a single Windows executable
- Microsoft Access 2000 (`.mdb`, Jet 4.0) for data, password-protected
- ADO 2.8 for data access; ADOX for schema creation
- Crystal Reports 8.5 for reporting (released 2001, end-of-life ~2005)
- MSCOMCTL.OCX / MSCOMCT2.OCX for ListView, Toolbar, DateTimePicker
- ~13,000 lines of behavioral code across 7 modules and 16 forms

The frozen source is preserved under [`ORIGINAL-CODE/star-hotel-vb6-main/`](ORIGINAL-CODE/star-hotel-vb6-main/) including the pre-built `StarHotel.exe`, all redistributables, and the demo Access database. **Do not modify files under `ORIGINAL-CODE/`** — it is the immutable reference for everything else in the repo.

## Project phases

| Phase | Status | Artifacts |
|---|---|---|
| Reverse-engineering | **Complete** | 16 form analyses, 7 module analyses, 4 project-artifact analyses, 2 architecture diagrams |
| Bug & security cataloguing | **Complete** | [`BUGS-MITIGATIONS.md`](docs/BUGS-MITIGATIONS.md) — 150 findings with severity and mitigations |
| Migration route selection | **Complete** | [`MIGRATION-OPTIONS.md`](docs/MIGRATION-OPTIONS.md) — 5 routes evaluated, recommendation made |
| Implementation playbook | **Complete** | [`MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md) — 14-arc roadmap with slice-loop protocol |
| Forward-port: Arc 01 (Foundation) | Not started | — |
| Forward-port: Arcs 02–14 | Not started | — |
| v1.0.0 release | Not started | — |

## Documentation index

All docs live under [`docs/`](docs/). Each is self-contained but references peers.

### Reverse-engineering analyses

| Document | Subject |
|---|---|
| [`docs/forms/README.md`](docs/forms/README.md) | System overview, form catalog, hub-and-spoke architecture |
| [`docs/forms/UI-FLOW.md`](docs/forms/UI-FLOW.md) | Lifecycle phases and navigation flow across the 16 forms |
| [`docs/forms/`](docs/forms/) | 16 per-form analyses (`ADMIN-FORM.md`, `BOOKING-FORM.md`, `DASHBOARD-FORM.md`, `DATABASE-FORM.md`, `DIALOG-FORM.md`, `FIND-CUSTOMER-FORM.md`, `MODULE-ACCESS-FORM.md`, `PRINT-FORM.md`, `REPORT-FORM.md`, `REPORT-MAINTAIN-FORM.md`, `ROOM-MAINTAIN-FORM.md`, `ROOM-TYPE-MAINTAIN-FORM.md`, `SPLASH-FORM.md`, `USER-CHANGE-PASSWORD-FORM.md`, `USER-LOGIN-FORM.md`, `USER-MAINTAIN-FORM.md`) |
| [`docs/modules/`](docs/modules/) | 7 per-module analyses for `modCommon`, `modDatabase`, `modEncryption`, `modFunction`, `modGlobalVariable`, `modMain`, `modTextFile` |
| [`docs/modules/DATA-FLOW-DIAGRAM.md`](docs/modules/DATA-FLOW-DIAGRAM.md) | Layer-cake architecture: how a click becomes a query becomes a row |
| [`docs/project/`](docs/project/) | Build/packaging artifact analyses (`PROJECT-VBP.md`, `RES-FILE.md`, `VBW-FILE.md`, `MANIFEST-FILE.md`) |

### Cross-cutting and forward-port

| Document | Subject |
|---|---|
| [`docs/BUGS-MITIGATIONS.md`](docs/BUGS-MITIGATIONS.md) | 150 catalogued bugs, security flaws, dead code, and smells across 10 categories. Each has location, severity, description, and a concrete mitigation. The triage list for the rewrite. |
| [`docs/MIGRATION-OPTIONS.md`](docs/MIGRATION-OPTIONS.md) | Five candidate migration routes (incremental .NET via VBUC, greenfield C#/.NET 10, Python/PySide6, Java/JavaFX, TypeScript/Tauri 2 or web) with pros/cons, roadmaps, tooling, and effort estimates. Side-by-side comparison and opinionated recommendation. |
| [`docs/MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md) | Implementation playbook for the chosen route. Target architecture (SOLID + layered), tech stack, coding standards, test strategy, Git workflow, slice-loop protocol, 14-arc roadmap, templates. |
| [`docs/DEV-LOG.md`](docs/DEV-LOG.md) | Append-only chronological journal, one entry per slice. Empty until Arc 01 begins. |
| [`docs/arcs/`](docs/arcs/) | Per-arc planning files. Each arc gets a `ARC-NN-NAME.md` written before its slices begin. Currently empty. |
| [`CLAUDE.md`](CLAUDE.md) | Operating instructions for Claude Code sessions working in this repo. Architecture summary and pointers to all of the above. |

## Repository layout

```
vb6-rework-reverse-forward/
├── ORIGINAL-CODE/
│   └── star-hotel-vb6-main/        Frozen upstream VB6 source. Do not edit.
│       ├── Form/                   16 .frm/.frx pairs
│       ├── Module/                 7 .bas modules
│       ├── Data/DemoData.mdb       Password-protected Access database
│       ├── Report/                 9 Crystal Reports .rpt files
│       ├── publish/                Pre-built StarHotel.exe + all redistributables
│       └── StarHotel.vbp           VB6 project manifest
├── docs/
│   ├── forms/                      Per-form reverse-engineering analyses
│   ├── modules/                    Per-module reverse-engineering analyses
│   ├── project/                    Build/packaging artifact analyses
│   ├── arcs/                       Per-arc planning files (empty until Arc 01)
│   ├── BUGS-MITIGATIONS.md
│   ├── DEV-LOG.md
│   ├── MIGRATION-OPTIONS.md
│   └── MIGRATION-DOTNET.md
├── CLAUDE.md
├── LICENSE                         (TBD — see License section)
└── README.md
```

When the rewrite begins, `src/`, `tests/`, and `tools/` will appear at the repo root per the project structure defined in [`MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md).

## Migration approach

The full evaluation of options is in [`docs/MIGRATION-OPTIONS.md`](docs/MIGRATION-OPTIONS.md). Summary:

- **Five routes were evaluated** — incremental VB6→.NET via commercial migrator, greenfield C#/.NET, Python, Java, TypeScript.
- **Recommended:** greenfield C#/.NET 10 + Avalonia 11 + EF Core + SQLite + QuestPDF + Argon2id. Rationale: the codebase is small enough that auto-migration's licensing cost outweighs its time savings, the 150 catalogued bugs are easier to design out than fix, and the .NET ecosystem fits a Windows-first hotel POS without locking out cross-platform deployment.
- **Total estimated effort:** 5–9 months single-developer through to v1.0.0.

The implementation is structured as **14 arcs** (epics), each broken into stories and then into commit-sized slices:

| Arc | Focus |
|---|---|
| 01 | Foundation — solution skeleton, DI, test infrastructure |
| 02 | Data Layer — SQLite + EF Core, schema, repositories |
| 03 | Auth & Sessions — Argon2id, login flow, session lifecycle |
| 04 | Domain Core — Booking, Room, User aggregates with audit events |
| 05 | Reports — QuestPDF integration, rebuild 9 reports |
| 06 | Presentation Foundation — Avalonia shell, navigation, chrome |
| 07 | Login & Dashboard UI |
| 08 | Booking Lifecycle UI |
| 09 | Find Customer & Reports UI |
| 10 | Admin Tools UI |
| 11 | Cross-cutting Features — idle watchdog, audit tables, logging |
| 12 | Migration Tooling — `.mdb` → SQLite importer |
| 13 | Polish & Hardening — manifest, accessibility, performance |
| 14 | Release — installer, signing, cutover |

Each arc gets a planning file in [`docs/arcs/`](docs/arcs/) before its slices begin.

## Reading order for newcomers

If you've just landed on this repo and want to understand it quickly:

1. **This README** for orientation (you are here).
2. [`docs/forms/README.md`](docs/forms/README.md) for what the original system *does* and how it's organized — system overview and form catalog.
3. [`docs/modules/DATA-FLOW-DIAGRAM.md`](docs/modules/DATA-FLOW-DIAGRAM.md) for how a click becomes a database query — the architectural layer model.
4. [`docs/BUGS-MITIGATIONS.md`](docs/BUGS-MITIGATIONS.md) for what's wrong with the original — the 150-finding catalog with severities.
5. [`docs/MIGRATION-OPTIONS.md`](docs/MIGRATION-OPTIONS.md) for how the modernization could be approached — five evaluated routes plus recommendation.
6. [`docs/MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md) for the implementation playbook — architecture, workflow, roadmap.
7. Individual form/module analyses in [`docs/forms/`](docs/forms/) and [`docs/modules/`](docs/modules/) on demand, when context is needed for a specific area.

About 90 minutes of reading covers items 1–6 end to end. The per-form/module analyses are reference material best read when working on a specific area.

## Development workflow

Once Arc 01 begins, the project follows the **slice loop** protocol defined in [`docs/MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md). Summary:

- Work is broken into Arcs → Stories → Slices. Each slice is 1–3 days, ≤ 500 lines of meaningful diff.
- Implementer (LLM agent or developer) creates a `slice/<arc>-<story>-<slice>-<short-name>` branch, writes code + tests + doc updates, runs the full test gate (unit + integration + smoke), and presents the diff.
- Reviewer (the user) reviews, stages, commits, and squash-merges to `main`. **The implementer never commits to `main` directly.**
- Each slice updates [`DEV-LOG.md`](docs/DEV-LOG.md) (one entry, append-only) and the active arc file (slice marked complete). When a slice resolves a finding from [`BUGS-MITIGATIONS.md`](docs/BUGS-MITIGATIONS.md), that finding is struck through with a slice reference.
- `main` is protected: linear history, squash-merge only, all CI gates green.
- Each arc completes with a tag `arc-NN-complete` and a retro entry in `DEV-LOG.md`.

Conventional Commits style is used for all commits. See [`MIGRATION-DOTNET.md`](docs/MIGRATION-DOTNET.md) § Git workflow for the exact conventions.

## Status

| Component | Status |
|---|---|
| Reverse-engineering documentation | Complete (29 docs, ~290KB) |
| Bug catalog | Complete (150 findings) |
| Migration option analysis | Complete (5 routes evaluated) |
| Implementation playbook | Complete (14 arcs scoped) |
| Source code (rewrite) | Not started |
| Tests | Not started |
| Installer / release | Not started |

There is no build system in this repository yet. The original VB6 source under `ORIGINAL-CODE/` is buildable only inside the VB6 IDE on Windows with Crystal Reports 8.5 and ADO 2.8 installed; a pre-built `StarHotel.exe` plus all redistributables sits in `ORIGINAL-CODE/star-hotel-vb6-main/publish/` for those who want to run the original.

## License

This repository contains two distinct bodies of work with different licensing considerations:

- **Original VB6 source** under `ORIGINAL-CODE/star-hotel-vb6-main/` is MIT-licensed by Poon Yip Hoon (2021). See [`ORIGINAL-CODE/star-hotel-vb6-main/LICENSE`](ORIGINAL-CODE/star-hotel-vb6-main/LICENSE).
- **Reverse-engineering documentation, migration plans, and any forward-port code** authored in this repository — license to be confirmed before public release. Anticipated to be MIT to match the upstream.

A repo-root `LICENSE` file will be added when the licensing decision is finalized for the rewrite.

## Acknowledgments

- **Aeric Poon (Poon Yip Hoon)** — original author of StarHotel and `modDatabase`/`modEncryption`/`modCommon`/`modTextFile`. The reverse-engineering documentation in this repo is only possible because the upstream code was open-sourced under a permissive licence and is well-structured enough to be readable line by line.
- **Computerise System Solutions** — original publisher.

The reverse-engineering analyses in `docs/forms/`, `docs/modules/`, and `docs/project/` are observational documentation of the original code's behavior. They do not include or redistribute the original source itself; that lives only under `ORIGINAL-CODE/` in unmodified form.
