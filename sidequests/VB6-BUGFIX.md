# SQ1 — Bugfix the original VB6 code

**Status:** Idea
**Owner:** unassigned
**Last updated:** 2026-05-04

## Concept

Back-port a tight, surgical set of fixes to the original VB6 source under `ORIGINAL-CODE/star-hotel-vb6-main/` so any deployed StarHotel sites still running the legacy build can survive the period between *now* and *whenever the .NET rewrite reaches feature parity*. The audit work for this is already done — `docs/BUGS-MITIGATIONS.md` enumerates ~150 issues with severities and mitigations. The side quest is the *patching*, not the *finding*.

## Why it's a side quest

The repo's stated convention is that `ORIGINAL-CODE/` is a **frozen reference copy**, and the project's `CLAUDE.md` is explicit: *"Do not modify it. New work goes under `docs/` (or a future `src/` for the rewrite)."* So this side quest deliberately violates that convention — it would require either lifting the freeze for a tagged maintenance branch, or forking `ORIGINAL-CODE/` into a sibling directory like `LEGACY-MAINT/` so the unmodified reference stays intact.

It's off the critical path because the main quest's bet is "rewrite, don't repair." If the .NET port lands on schedule, no one needs the VB6 patches. This side quest is insurance for the case where the port slips and a real customer is sitting on a `Critical`-severity issue *today*.

## Scope

### In scope (candidate fix list, drawn from `docs/BUGS-MITIGATIONS.md`)

A *minimum viable patch set* — the smallest set of changes that materially reduces the worst risks without rewriting the architecture. Roughly the Critical-severity items only:

| Area              | Fix                                                                                                              | Why patch in VB6                                                                                       |
| ----------------- | ---------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| Auth              | Remove the auto-fill of `admin`/`admin` when the copyright label on `frmUserLogin` is clicked.                   | One-line change. Eliminates a backdoor with zero behavioral risk to legitimate users.                  |
| Auth              | Remove the hardcoded `expert` developer password gating `frmReportMaintain`.                                     | One-line change. Replace with a real `ModuleAccess` check against the existing permissions matrix.     |
| Auth              | Remove the `gintUserIdle = 0` debug override that disables idle-logout.                                          | One-line change. Restores the watchdog the team already wrote.                                         |
| SQL injection     | Parameterize the highest-traffic queries in `modCommon`'s `SQL_WHERE_*` family (customer search, booking search).| ADO already supports parameters via `Command.Parameters`; the global `gstrSQL` builder can be bypassed for the specific call sites with the highest exposure. |
| Crash recovery    | Wrap `Booking` save and `Room` rate-change in a single ADO transaction in `modDatabase.QuerySQL` callers.        | Eliminates the partial-save risk for the two operations the user notices most when they corrupt data. |
| Password storage  | Migrate `GoldFishEncode` hashes to PBKDF2 on first successful login (lazy migration), keeping the column shape.  | Possible to do entirely in `modEncryption` + a single login-path change. No schema migration needed.   |

The principle is: **change one or two lines per fix wherever possible**. Anything that requires architectural reshape (the SQL builder, the global session state, the report engine) is explicitly out — that's what the rewrite is for.

### Out of scope

- Replacing Crystal Reports 8.5.
- Replacing the Access `.mdb` with anything else.
- Adding new features.
- Refactoring `modCommon`'s SQL builder pattern.
- Touching `modGlobalVariable` semantics.
- Rebuilding the permissions model. (Existing four-group `ModuleAccess` matrix stays.)
- Anything that would invalidate the existing `LogBooking` / `LogRoom` history rows or change the schema.

## Dependencies on the main quest

Effectively none — this can run entirely in parallel. It does not depend on any arc of `docs/MIGRATION-DOTNET.md`. In fact it competes with the main quest for attention; that's the reason it's a side quest rather than an arc.

The one weak coupling: if this side quest fixes the SQL injection items in `modCommon`, the rewrite's Application-layer port of those queries should adopt the same call-site shape (parameter names, ordering) so behavior parity is verifiable against either the patched VB6 or the new .NET code.

## Sketch of approach

1. **Lift the freeze with a fork.** Copy `ORIGINAL-CODE/star-hotel-vb6-main/` to `LEGACY-MAINT/star-hotel-vb6-main/` and tag the original commit so the reference copy remains pristine. All edits land under `LEGACY-MAINT/`.
2. **Stand up a VB6 build environment.** Document the exact COM dependency stack already listed in the project `CLAUDE.md` (ADO 2.8, ADOX, Crystal Reports 8.5 ActiveX runtime, MSCOMCTL/MSCOMCT2). Capture install order and registration steps in `LEGACY-MAINT/BUILD.md`. Without this, no one else can pick up the side quest.
3. **Pick a single Critical fix and ship it end-to-end** before queuing the rest, to flush out the build/test/release loop. Recommended first: removing the `admin`/`admin` auto-fill — it's a one-line change with no behavioral surprise and forces the whole pipeline (build, smoke test, sign, package) into existence.
4. **Smoke-test protocol.** No automated tests exist and we won't add any in VB6. Define a manual smoke-test checklist (login as admin, create booking, take payment, check in, check out, run a couple of reports) in `LEGACY-MAINT/SMOKE.md`. Run it end-to-end before each release.
5. **Versioning.** Bump the `Version` field in each modified module's header (the docs already note these are load-bearing — `1.2.22` etc.) and tag releases as `vb6-maint-1.x.y`.
6. **Distribution.** Mirror the existing `publish/` pattern (pre-built `StarHotel.exe` plus all redistributables) so deployment to a customer site is the same drop-in copy it always was.

## Effort estimate

Order of magnitude: **2–4 weeks of focused work** for the Critical-only patch set, *if* a working VB6 build environment already exists. Add 3–5 days if the build environment has to be reconstituted from scratch (Windows VM, COM registration, Crystal Reports 8.5 install, IDE quirks). The patches themselves are small; the slow part is the absence of any automated test gate, so every fix needs manual smoke-testing through the full booking lifecycle.

If the scope creeps to include High-severity items (the rest of the SQL injection list, the Void-feature reactivation, the missing transactions on multi-step business operations), double the estimate.

## Risks & open questions

- **Who owns "deployed sites"?** This side quest only makes sense if there's at least one customer running the legacy build that would actually receive a patch. If the answer is "nobody", abandon this side quest and put the time into the main quest. Worth checking with whoever knows the customer roster before committing.
- **Crystal Reports 8.5 is end-of-life.** Even if we patch the VB6, the report runtime is unsupported by SAP and won't work cleanly on newer Windows. Any patch release inherits this risk.
- **Regression cost is high.** Without an automated test suite, every patch ships on smoke-testing alone. Two changes in flight at once will be hard to bisect when something breaks. Recommend strict one-fix-per-release discipline.
- **The `LEGACY-MAINT/` fork dilutes the "frozen reference" narrative.** Need to be loud in the side quest README and the top-level `CLAUDE.md` about which directory is the canonical reference and which is the maintenance branch, or future contributors will diff against the wrong tree.
- **Open question:** is it cheaper to ship a *minimum-viable* .NET port (login + booking only) faster than to maintain VB6 in parallel? If yes, this side quest is dominated by the main quest and should be abandoned.
