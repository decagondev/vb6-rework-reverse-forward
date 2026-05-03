# Migration: Greenfield .NET 10 / C# / Avalonia / SQLite

The implementation playbook for the recommended migration path from `MIGRATION-OPTIONS.md` (Route B). This doc is the operating manual for the rewrite — it defines the target architecture, the per-slice development loop, the documentation artifacts that must be kept current, and the full Epic → Story → Slice roadmap.

It is written so that a future implementer (whether human or LLM-assisted) can pick up at any slice and continue without losing context.

## Contents

1. [Purpose & scope](#purpose--scope)
2. [Target architecture](#target-architecture)
3. [Solution & project structure](#solution--project-structure)
4. [Tech stack & rationale](#tech-stack--rationale)
5. [Coding standards](#coding-standards)
6. [Test strategy](#test-strategy)
7. [Required documentation artifacts](#required-documentation-artifacts)
8. [Git workflow](#git-workflow)
9. [The slice loop](#the-slice-loop)
10. [Epic / Story / Slice hierarchy](#epic--story--slice-hierarchy)
11. [Roadmap: 14 arcs](#roadmap-14-arcs)
12. [Templates](#templates)
13. [Definition of Done](#definition-of-done)

---

## Purpose & scope

This document is the single source of truth for **how** the rewrite happens. The **what** lives in three companion docs that must be kept current alongside this one:

- `BUGS-MITIGATIONS.md` — the catalog of 150 issues to fix or design out, by severity.
- `DEV-LOG.md` — the chronological development journal, one entry per slice.
- `docs/arcs/ARC-NN-NAME.md` — one file per arc, holding stories, slices, acceptance criteria, and per-arc references to bugs/mitigations.

Anything in this file describes the **process and target architecture**. Anything specific to a particular slice or feature lives in the relevant arc file. If a process rule changes mid-rewrite, edit this file and reference the change in `DEV-LOG.md`.

Out of scope: business-process decisions (e.g., "should the new app support multi-tenancy"), deployment infrastructure beyond a single-machine MSI, and per-customer customization.

---

## Target architecture

### SOLID principles, applied concretely

The original violates almost every SOLID principle (global `gstrSQL`, fat modules, hardcoded form lists, deep coupling between forms). The rewrite designs each one out.

| Principle | What we do | What we don't do |
|---|---|---|
| **Single Responsibility** | One ViewModel per view; one Repository per aggregate root; one UseCase per business operation. | Don't put validation, persistence, and UI logic in the same class. No "manager" / "helper" / "utility" classes that grow without bound. |
| **Open/Closed** | Strategy pattern for report generation (`IReportRenderer` with one implementation per report). New reports added by adding a class, not modifying existing ones. Same pattern for `IPasswordHasher`. | Don't switch on enum types in business logic — add a new type/strategy. |
| **Liskov Substitution** | Repository interfaces have concrete impls (`SqliteBookingRepository : IBookingRepository`). Tests run against in-memory or fake impls without changing consumer code. | Don't have method-specific behavior buried in concrete types that fakes can't replicate. |
| **Interface Segregation** | Small focused interfaces: `IPasswordHasher`, `IClock`, `ISessionService`, `IBookingRepository`. | No fat `IDataService` with 50 methods. No `IUtility`. |
| **Dependency Inversion** | Application layer depends on abstractions defined in Application. Infrastructure provides implementations. The composition root is the only place concretes are wired up. | No `static` singletons. No `new SqliteConnection(...)` outside Infrastructure. No global state. |

### Layer model

Strict one-way dependency arrows. Anything below depends only on things above it (or pure types).

```
┌─────────────────────────────────────────────────────────┐
│  StarHotel.Desktop (Avalonia entry point)               │  ← depends on all below
├─────────────────────────────────────────────────────────┤
│  StarHotel.Presentation (ViewModels, Views, Navigation) │  ← depends on Application, Domain
├─────────────────────────────────────────────────────────┤
│  StarHotel.Infrastructure (SQLite, Argon2, QuestPDF)    │  ← depends on Application, Domain
├─────────────────────────────────────────────────────────┤
│  StarHotel.Application (Use cases, abstractions)        │  ← depends on Domain only
├─────────────────────────────────────────────────────────┤
│  StarHotel.Domain (Entities, value objects, events)     │  ← depends on nothing
└─────────────────────────────────────────────────────────┘
```

- **Domain** has no external dependencies. Pure C#. No EF Core attributes, no JSON attributes, no Avalonia.
- **Application** defines `interface IBookingRepository` etc. and use cases (`CreateBookingUseCase`). It does **not** know about SQLite or Avalonia.
- **Infrastructure** implements Application's interfaces. SQLite, EF Core, QuestPDF, Argon2 all live here. Forbidden from referencing Presentation.
- **Presentation** consumes use cases via constructor injection. ViewModels never touch repositories or DbContext directly.
- **Desktop** is the composition root — wires everything up via `Microsoft.Extensions.DependencyInjection`, picks the SQLite path, and starts Avalonia.

### Modular design within a layer

Each domain area gets a self-contained folder across all relevant layers:

```
StarHotel.Domain/Booking/        Booking.cs, BookingState.cs, BookingId.cs
StarHotel.Application/Booking/   IBookingRepository.cs, CreateBookingUseCase.cs
StarHotel.Infrastructure/Booking/ SqliteBookingRepository.cs, BookingMapper.cs
StarHotel.Presentation/Booking/  BookingViewModel.cs, BookingView.axaml
```

Adding a feature touches one folder per layer, not 15 files scattered across the solution.

---

## Solution & project structure

```
vb6-rework-reverse-forward/
├── ORIGINAL-CODE/                  (frozen reference; never edit)
├── docs/
│   ├── arcs/                       (per-arc plans)
│   │   ├── ARC-01-FOUNDATION.md
│   │   ├── ARC-02-DATA-LAYER.md
│   │   └── ...
│   ├── BUGS-MITIGATIONS.md
│   ├── DEV-LOG.md
│   ├── MIGRATION-OPTIONS.md
│   ├── MIGRATION-DOTNET.md         (this file)
│   ├── forms/  modules/  project/
│   └── ...
├── src/
│   ├── StarHotel.Domain/
│   ├── StarHotel.Application/
│   ├── StarHotel.Infrastructure/
│   ├── StarHotel.Presentation/
│   └── StarHotel.Desktop/
├── tests/
│   ├── StarHotel.Domain.Tests/
│   ├── StarHotel.Application.Tests/
│   ├── StarHotel.Infrastructure.Tests/
│   ├── StarHotel.Presentation.Tests/
│   └── StarHotel.Smoke.Tests/
├── tools/
│   └── StarHotel.MdbImport/         (one-shot .mdb → SQLite migrator; Arc 12)
├── StarHotel.slnx                   (VS 2026 solution format)
├── Directory.Build.props            (shared MSBuild config: nullable, lang version, analyzers)
├── .editorconfig
├── .gitignore
└── README.md
```

The `Directory.Build.props` enforces one set of compiler/analyzer settings everywhere. No per-csproj drift.

---

## Tech stack & rationale

| Concern | Choice | Why |
|---|---|---|
| Runtime | **.NET 10 LTS** | Released Nov 2025, supported through Nov 2028. Latest stable LTS. |
| Language | **C# 14** with nullable reference types enabled | Fewer NREs, more expressive. |
| UI | **Avalonia 11.x** | Cross-platform (Windows/macOS/Linux), MIT licence, mature MVVM. |
| MVVM | **CommunityToolkit.Mvvm** | `[ObservableProperty]`, `[RelayCommand]` source generators — terse, fast. |
| DI | **Microsoft.Extensions.DependencyInjection** | Standard. |
| ORM | **EF Core 10** with SQLite provider | Migrations, LINQ, change tracking. |
| Database | **SQLite** | Single file, zero admin, fast enough. |
| Reporting | **QuestPDF** | Code-first, MIT, no designer needed for ~9 reports. |
| Auth | **Konscious.Security.Cryptography.Argon2** | Argon2id is the modern default. |
| Logging | **Serilog** + rolling file sink | Replaces both `LogErrorDB` and `LogErrorText`. |
| Testing | **xUnit** + **FluentAssertions** + **NSubstitute** | Standard, low ceremony. |
| UI testing | **Avalonia.Headless** + **Avalonia.Headless.XUnit** | Run UI tests without an X server. |
| Time | **`TimeProvider`** (BCL, .NET 8+) injected as `IClock` | Testable, replaces `Now()`. |

No commercial tooling. No paid licenses. Total runtime dependency is ~50MB self-contained.

---

## Coding standards

Defined in `.editorconfig` and `Directory.Build.props`; the highlights:

- `Nullable` enabled, treated as error.
- `TreatWarningsAsErrors` true; build red on any warning.
- `LangVersion` latest; preview features off unless flagged in an arc plan.
- File-scoped namespaces.
- `var` only when the type is obvious from the right-hand side.
- No `using` directives outside their namespace block; sorted with `System.*` first.
- Public APIs documented with XML doc comments only when the *why* is non-obvious. Don't restate the method name.
- One public type per file; matching filename.
- Internal-by-default; `public` only when consumed across project boundaries.
- No `static` mutable state. No singletons except via DI scopes.
- No regions.
- Async methods end in `Async` and take a `CancellationToken` parameter (last, defaulted).

If a slice introduces an exception to any of these, document it in the arc file and `DEV-LOG.md`.

---

## Test strategy

Every slice ends with three test gates green before handoff to the user:

### 1. Unit tests
- Pure logic in Domain and Application layers.
- Run on every build. Fast (< 1s for the suite at start; budget < 30s at completion).
- Use NSubstitute for interface fakes; do not mock concrete types.

### 2. Integration tests
- Infrastructure layer against a real ephemeral SQLite (`:memory:` or temp file).
- Repository round-trips, EF migrations apply cleanly, transaction boundaries behave.

### 3. Smoke tests
- End-to-end happy paths through the composition root. **One smoke test per major user flow.** Examples:
  - Login → Dashboard → Open booking → Save → Logout
  - Login → Find Customer → Print Customer History
  - Login → Reports → Daily Booking → Preview
- Run via `Avalonia.Headless` so they execute on CI without a display server.
- Slow but small in number. ~10 smoke tests at completion.

### 4. Regression tests
- Defined as: **the full unit + integration + smoke suite**. "Regression" doesn't mean a separate suite; it means we don't ship a slice that breaks any existing test.
- Every slice runs the full suite before handoff. CI enforces it on push.

### Test gates per slice

```
Slice ready for commit when:
  ✓ dotnet build → 0 warnings, 0 errors
  ✓ dotnet test (unit + integration) → all green
  ✓ dotnet test (smoke) → all green
  ✓ Code coverage on new Domain/Application code ≥ 80%
  ✓ DEV-LOG.md entry added
  ✓ Arc file updated (slice marked complete)
  ✓ BUGS-MITIGATIONS.md updated if any finding resolved (strikethrough)
```

The implementer (LLM or human) runs these locally before handing off.

---

## Required documentation artifacts

Three files must stay synchronized with the codebase. Failure to update any of them blocks the slice from being marked done.

### `docs/DEV-LOG.md` — append-only development journal

One entry per slice, newest-first. Captures what was done, what was decided, what's still open. This is the artifact that makes resumption-after-context-loss possible.

Skeleton (also in [Templates](#templates)):

```markdown
## YYYY-MM-DD — Arc NN / Story NN-NN / Slice NN-NN-NN — <short title>

**Branch:** `slice/NN-NN-NN-short-name`
**Status:** ready-for-commit | in-progress | blocked

### Worked on
- ...

### Decisions
- <decision>: <one-line rationale>

### Tests
- Unit: 42 / 42 passing
- Integration: 12 / 12 passing
- Smoke: 3 / 3 passing
- Coverage on touched files: 87%

### BUGS-MITIGATIONS.md updates
- §1.1 (GoldFishEncode hash) — resolved by introducing `IPasswordHasher` + Argon2id impl

### Open questions / follow-ups
- ...
```

### `docs/arcs/ARC-NN-NAME.md` — per-arc plan

One file per arc, written before the arc starts. Lists stories, slices, acceptance criteria, and which `BUGS-MITIGATIONS.md` items the arc resolves. Slices get checked off as they complete.

### `docs/BUGS-MITIGATIONS.md` — kept current

When a slice resolves a finding, **strike through** the title in the catalog (per the existing convention: `~~Title~~`) and append a one-line note: `*(Resolved in slice NN-NN-NN, YYYY-MM-DD.)*`. Do not delete — the audit trail matters.

When a new bug is discovered during the rewrite (and is a property of the original or a constraint on the rewrite, not a transient bug in the new code), append it to the appropriate category and bump the summary count.

---

## Git workflow

### Branching model

- `main` is the protected trunk. Every commit on `main` builds, passes tests, and is shippable.
- `slice/<arc>-<story>-<slice>-<short-name>` for each slice. Example: `slice/01-02-03-add-domain-booking-entity`.
- No long-lived feature branches. A slice lives 1–3 days, then merges.
- Releases are tagged on `main`: `v0.1.0`, `v0.2.0`, etc. (Tags are an arc-completion artifact.)

### Commit conventions

Conventional Commits style:

```
<type>(<scope>): <subject>

<body — what and why, not how>
```

- `feat`, `fix`, `refactor`, `test`, `docs`, `chore`, `perf`, `build`, `ci`.
- `<scope>` is the touched module name in lowercase (e.g., `booking`, `auth`, `infra`).
- Subject ≤ 72 chars, imperative mood ("add", not "added").
- Body explains the *why* and references the arc/slice.

Example:

```
feat(auth): add IPasswordHasher with Argon2id implementation

Replaces GoldFishEncode obfuscation (BUGS-MITIGATIONS §1.1, Critical).
Argon2id with t=3, m=64MB, p=1 — calibrated to ~100ms on baseline hardware.

Slice 03-01-01.
```

### Branch protection

- `main` requires a passing CI check before merge.
- Linear history (rebase, no merge commits) — keeps `git log` legible.
- Squash-merge slices into `main` so each slice is one commit on the trunk.

### Tags

- Each completed arc → an annotated tag `arc-NN-complete`.
- Each shippable build → a semver tag `vX.Y.Z`.

---

## The slice loop

This is the per-slice operational protocol. The implementer and the user have well-defined roles. The user owns Git mutations to `main`; the implementer never pushes, never merges, never modifies `main` directly.

### Roles

- **Implementer** (LLM agent or developer): creates the slice branch, writes code and tests, runs the test gates, updates docs, presents the diff for review.
- **User** (human reviewer): reviews the diff, stages, commits, merges into `main`.

### Step-by-step

1. **Slice selected.** User says "start slice X" (or implementer picks the next unticked slice from the current arc file). Implementer reads `MIGRATION-DOTNET.md` (this file), the active arc file, the last 3–5 entries of `DEV-LOG.md`, and any referenced bugs.

2. **Branch created.** Implementer runs:
   ```bash
   git checkout main
   git pull --ff-only
   git checkout -b slice/<arc>-<story>-<slice>-<short-name>
   ```

3. **Work performed.** Implementer writes code, tests, and the relevant doc updates. Commits along the way are *fine* on the slice branch — they will be squashed at merge time.

4. **Test gates run.**
   ```bash
   dotnet build /warnaserror
   dotnet test
   dotnet test --filter Category=Smoke
   ```
   All must be green.

5. **Docs updated.** `DEV-LOG.md` gets a new entry (status: `ready-for-commit`). Arc file's slice checkbox flipped. `BUGS-MITIGATIONS.md` strikethrough applied if any finding resolved.

6. **Handoff to user.** Implementer reports:
   > "Slice `<id>` ready for review. Branch: `slice/...`. Tests: green (X unit, Y integration, Z smoke). Files changed: <list>. Run `git diff main...HEAD` to review."

7. **User reviews.** User runs `git diff main...HEAD`, opens changed files, sanity-checks. If changes are needed, user says so; implementer iterates on the same branch.

8. **User stages and commits.** When satisfied, user squash-merges the slice into `main`:
   ```bash
   git checkout main
   git merge --squash slice/<arc>-<story>-<slice>-<short-name>
   git commit -m "<conventional commit message>"
   ```
   (Or uses a GUI / tooling that does the same.)

9. **Cleanup.** User deletes the slice branch:
   ```bash
   git branch -D slice/<arc>-<story>-<slice>-<short-name>
   ```

10. **Next slice.** User says "next slice" — back to step 1.

### Implementer constraints

- Implementer **never** runs `git commit`, `git merge`, or `git push` against `main`.
- Implementer **never** rewrites history that the user has already committed.
- Implementer **does** create the slice branch, **does** commit on the slice branch (these get squashed at merge), **does** modify `DEV-LOG.md` / arc files / `BUGS-MITIGATIONS.md`.
- If a slice grows beyond 1–3 days of work or 500 lines of meaningful diff, **stop and split** into smaller slices. Update the arc file accordingly and report back.

### When tests fail

- If a test gate is red at handoff, the slice is **not** ready for commit. Implementer reports failure, suggests fix or asks the user.
- Never disable a failing test to make a slice green. If a test is wrong, fix it as part of the slice. If a test is brittle, log it in `DEV-LOG.md` follow-ups.

---

## Epic / Story / Slice hierarchy

| Level | Granularity | Lives in | Definition |
|---|---|---|---|
| **Arc / Epic** | Multi-week feature area | `docs/arcs/ARC-NN.md` | A capability the system either has or doesn't (e.g., "Auth & Sessions"). 14 arcs span the rewrite. |
| **Story** | A user-visible capability | Section within an arc file | A single observable behavior (e.g., "User can log in with username and password"). 3–8 per arc. |
| **Slice** | An atomically-shippable change | Subsection within a story | 1–3 days of work, builds + tests green at the end (e.g., "Add `IPasswordHasher` and Argon2id impl with unit tests"). 2–5 per story. |

### Sizing rules

- **Arc:** 4–8 weeks single-developer. If bigger, split it.
- **Story:** 1–2 weeks. If bigger, split it.
- **Slice:** 1–3 days, ≤ 500 lines of meaningful diff (excluding generated files, tests). If bigger, split it.
- **Slice numbering:** `<arc>-<story>-<slice>`, two digits each. `03-02-04` = Arc 3, Story 2, Slice 4.

### Example

> **Arc 03 — Auth & Sessions**
> └ **Story 03-01 — User can log in**
>   ├ Slice 03-01-01: Add `IPasswordHasher` + Argon2id impl + tests
>   ├ Slice 03-01-02: Add `LoginUseCase` with rate limiting + tests
>   ├ Slice 03-01-03: Add `LoginViewModel` + view + smoke test
>   └ Slice 03-01-04: Wire DI in composition root + happy-path E2E

---

## Roadmap: 14 arcs

The full rewrite, broken into arcs. Each arc gets its own file in `docs/arcs/` written before the arc starts. The list below is the index; details live in the arc files.

### Arc 01 — Foundation
Solution scaffolding, projects, DI, test infrastructure, CI skeleton, `Directory.Build.props`, `.editorconfig`, lint rules, first smoke test that just opens an empty Avalonia window.
- Resolves: structural prerequisites, no `BUGS-MITIGATIONS.md` items directly.
- Stories: Solution skeleton; Test infrastructure; CI baseline; Empty Avalonia shell.

### Arc 02 — Data Layer
SQLite + EF Core 10. Entity types in Domain. DbContext in Infrastructure. Migrations applied to a fresh DB. Repository interfaces in Application; concrete impls in Infrastructure. **No business logic yet** — just CRUD plumbing for all 12 tables.
- Resolves: `BUGS-MITIGATIONS.md` §5 (no FK constraints, denormalization documented), §2 (no parameterized queries — solved by EF Core).
- Stories: Schema + entities; Repositories; Migrations + seeding; Transaction-aware UnitOfWork.

### Arc 03 — Auth & Sessions
`IPasswordHasher` (Argon2id), `IClock`, `ISessionService`. `LoginUseCase` with rate limiting. Replaces `GoldFishEncode`, default `admin/admin`, copyright auto-fill, freeze-bug, `gintUserIdle = 0` override.
- Resolves: `BUGS-MITIGATIONS.md` §1.1, §1.2, §1.3, §1.4, §1.5, §1.6, §3.2, §7.1.
- Stories: Password hashing; Login flow; Session lifecycle; Forced password change.

### Arc 04 — Domain Core
Pure-domain entities: `Booking`, `Room`, `RoomType`, `User`, `ModuleAccess`, `Report`. Value objects (`BookingId`, `Money`, `DateRange`). Booking state machine. Audit-trail capture on Booking and Room writes.
- Resolves: `BUGS-MITIGATIONS.md` §3.7 (audit gaps, partially), §5.5, §5.7.
- Stories: Booking aggregate; Room aggregate; User aggregate; Audit events.

### Arc 05 — Reports
QuestPDF integration. `IReportRenderer` interface and one implementation per report. Rebuild all 7 reports (Daily/Weekly/Monthly Booking, Weekly Booking Graph, Shift by Staff, Shift All Staff, Customer Transaction History) plus Receipt and Temporary Receipt.
- Resolves: `BUGS-MITIGATIONS.md` §9.1 (Crystal Reports EOL), §5.6 (`UpdateWeekDayTable` race).
- Stories: Report infrastructure; Per-report implementations (one slice per report).

### Arc 06 — Presentation Foundation
Avalonia shell, navigation service, base ViewModel, theme resources, the chrome (header band with logo, toolbar, clock, user-id label) extracted as a reusable user control. Translates the original's "every form has the same chrome" pattern.
- Resolves: `BUGS-MITIGATIONS.md` §8.13 (color constants), §8.14 (idle-watchdog duplication — done once centrally).
- Stories: Shell + navigation; Chrome user control; Theme resources; Base ViewModel.

### Arc 07 — Login & Dashboard UI
First user-visible features. Login view + ViewModel + smoke test. Dashboard with the 55-button room grid, color states, F-key shortcuts. Right-click context menu.
- Resolves: nothing new (auth already done in Arc 03; this arc wires UI).
- Stories: Login view; Dashboard grid; Dashboard interactions; Keyboard shortcuts.

### Arc 08 — Booking Lifecycle UI
The biggest UI arc. Booking form: state machine (New → Booked → Occupied → Housekeeping), payments, deposits, refunds, business rules. Receipts wired to QuestPDF reports.
- Resolves: `BUGS-MITIGATIONS.md` §5.10 (dead 2pm branch), §4.7 (Void decision — either implement or remove), §5.2 (temp booking races).
- Stories: Booking form scaffold; State transitions; Payments + refunds; Receipts; Void (decision-pending).

### Arc 09 — Find Customer & Reports UI
Customer search with corrected operator-precedence in WHERE clauses. Report catalog UI with date filters. Print pipeline through QuestPDF.
- Resolves: `BUGS-MITIGATIONS.md` §2.5 (operator precedence), §2.6 (PrintHistory `gstrSQL` corruption), §2.7 (`$UserID$` substitution).
- Stories: Find Customer search; Customer detail panel; Report catalog; Report preview/print.

### Arc 10 — Admin Tools UI
Room/RoomType/User/ModuleAccess maintain forms. ReportMaintain with the developer-password gate **removed** (replaced with `MOD_REPORT_EDIT_EXPERT` permission only).
- Resolves: `BUGS-MITIGATIONS.md` §3.1 (expert password), §3.4 (no last-admin guard), §4.3 (Confirm field decorative), §4.4 (empty-password INSERT), §4.5 (no strength check), §3.9 (immediate-write ModuleAccess).
- Stories: Room maintain; Room type maintain; User maintain; Module access; Report maintain.

### Arc 11 — Cross-cutting Features
Idle watchdog (centralized — one timer in the shell, not per-form). Audit-trail tables for User, ModuleAccess, Report. Centralized error handling (Serilog). Session-globals-cleared-on-logout.
- Resolves: `BUGS-MITIGATIONS.md` §4.1 (idle override), §3.7 (audit gaps fully), §6.* (logging), §7.5 (session globals on logout).
- Stories: Idle watchdog; Audit-trail tables; Logging; Logout cleanup.

### Arc 12 — Migration Tooling
A separate console project (`tools/StarHotel.MdbImport/`) that reads an existing `.mdb` (via OLE DB or ODBC) and writes data into the new SQLite. Re-hashes passwords on first login post-migration. One-shot tool, not part of the runtime.
- Resolves: nothing new; enables cutover.
- Stories: Read .mdb; Map to new schema; Write to SQLite; Verify migration.

### Arc 13 — Polish & Hardening
DPI awareness manifest. Long-path manifest. `%APPDATA%`-based config and logs. Accessibility pass (keyboard navigation, screen reader labels). Localization scaffolding (English-only initially, but the schema supports it). Performance pass.
- Resolves: `BUGS-MITIGATIONS.md` §9.6, §9.7, §9.8 (manifest), §5.13 (config path).
- Stories: Manifest; Config paths; Accessibility; Performance.

### Arc 14 — Release
MSI/MSIX installer. Code signing. Auto-update channel (optional). Cutover runbook. Rollback plan. Tag `v1.0.0`.
- Resolves: deployment.
- Stories: Installer; Signing; Cutover runbook; v1.0.0 release.

---

## Templates

Copy these when creating new arcs / log entries / slices.

### Arc file template (`docs/arcs/ARC-NN-NAME.md`)

```markdown
# Arc NN — <Name>

**Status:** planned | in-progress | complete
**Started:** YYYY-MM-DD
**Completed:** YYYY-MM-DD (when done)
**Tag on completion:** `arc-NN-complete`

## Goal
One paragraph: what this arc delivers and why it's a meaningful unit.

## Out of scope
What this arc deliberately does *not* address (and which arc handles it instead).

## BUGS-MITIGATIONS.md items resolved
- §X.Y — <title>
- §X.Z — <title>

## Stories

### Story NN-01 — <title>
**Acceptance:** <observable behavior that proves the story is done>

#### Slice NN-01-01 — <title>
- [ ] Implement
- [ ] Tests (unit, integration, smoke as applicable)
- [ ] Doc updates (DEV-LOG, this file, BUGS-MITIGATIONS if applicable)
- [ ] Ready for commit

**Touches:** <list of folders / projects>
**Notes:** <any constraints or decisions specific to this slice>

#### Slice NN-01-02 — <title>
- [ ] Implement
- ...

### Story NN-02 — <title>
...

## Definition of Done for this arc
- [ ] All slices complete and merged
- [ ] All listed BUGS-MITIGATIONS items struck through
- [ ] Smoke tests added for any new user flow
- [ ] DEV-LOG.md has an arc-completion entry
- [ ] Tag `arc-NN-complete` applied to `main`
```

### DEV-LOG entry template

```markdown
## YYYY-MM-DD — Arc NN / Story NN-NN / Slice NN-NN-NN — <short title>

**Branch:** `slice/NN-NN-NN-short-name`
**Status:** ready-for-commit | in-progress | blocked

### Worked on
- ...

### Decisions
- <decision>: <one-line rationale>

### Tests
- Build: 0 warnings, 0 errors
- Unit: 42 / 42 passing
- Integration: 12 / 12 passing
- Smoke: 3 / 3 passing
- Coverage on touched files: 87%

### BUGS-MITIGATIONS.md updates
- §X.Y (<title>) — resolved by <how>

### Open questions / follow-ups
- ...
```

### Arc completion DEV-LOG entry

```markdown
## YYYY-MM-DD — Arc NN complete: <name>

**Tag:** `arc-NN-complete`
**Duration:** YYYY-MM-DD → YYYY-MM-DD (X weeks)
**Slices merged:** N
**BUGS-MITIGATIONS items resolved:** X (§A.B, §A.C, ...)

### Retrospective
- What worked: ...
- What didn't: ...
- Carried forward to Arc NN+1: ...
```

### Slice plan (lightweight, used during slice work)

When picking up a slice, the implementer writes a 5–10 line plan at the top of the slice's working scratchpad (not committed; just used during the work):

```
Slice NN-NN-NN — <title>
Goal: <one sentence>
Touches: <projects/folders>
Steps:
  1. ...
  2. ...
  3. Tests
  4. Doc updates
Test gates: unit / integration / smoke
Bugs resolved: §X.Y, §X.Z
```

---

## Definition of Done

### Slice DoD
- [ ] Code compiles with 0 warnings, 0 errors.
- [ ] All test gates green (unit, integration, smoke).
- [ ] Coverage on new Domain/Application code ≥ 80%.
- [ ] DEV-LOG.md entry written, status `ready-for-commit`.
- [ ] Arc file slice checkbox flipped.
- [ ] BUGS-MITIGATIONS.md updated if any finding resolved.
- [ ] No new `TODO` comments introduced unless tracked in the arc file as follow-up.
- [ ] Branch present and up-to-date with `main`.

### Story DoD
- [ ] All slices in the story complete and merged.
- [ ] Acceptance criterion observably met (smoke test or manual demo).
- [ ] Story marked complete in arc file.

### Arc DoD
- [ ] All stories in the arc complete.
- [ ] All listed BUGS-MITIGATIONS items struck through with slice references.
- [ ] Arc-completion DEV-LOG entry written.
- [ ] Tag `arc-NN-complete` on `main`.
- [ ] Retro captured in arc file.

### Project DoD (v1.0.0)
- [ ] All 14 arcs complete.
- [ ] All Critical and High items in BUGS-MITIGATIONS.md struck through (or explicitly out-of-scope with rationale).
- [ ] Cutover runbook validated against a staging copy of real customer data.
- [ ] Installer signed, smoke-tested on clean Windows VMs.
- [ ] `v1.0.0` tag applied to `main`.

---

## How to use this document

- **First-time pickup of the project.** Read this file end-to-end. Then read `MIGRATION-OPTIONS.md`, `BUGS-MITIGATIONS.md`, and the most recent arc file. Then read the last 5 entries of `DEV-LOG.md`. That's enough context to start the next slice.
- **Starting a new arc.** Create `docs/arcs/ARC-NN-NAME.md` from the template before writing any code. List stories and slices upfront — adjust as you learn, but start from a plan.
- **Starting a new slice.** Read the active arc file's slice description, the BUGS-MITIGATIONS items it references, and any code the slice will modify. Then create the branch and start.
- **Process changes mid-rewrite.** Edit this file. Note the change in `DEV-LOG.md`. Don't rewrite history of completed arcs — process changes apply forward.

This file is the operating manual. Keep it current. When the rewrite ships, this file becomes the post-mortem reference for how it was built.
