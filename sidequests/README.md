# Side Quests

This directory catalogs **optional, off-critical-path workstreams** that are tangentially related to the main reverse-then-forward rework of StarHotel but aren't required to ship the .NET 10 / Avalonia port described in `docs/MIGRATION-DOTNET.md`.

The main migration is the **main quest**: read the original VB6, document it, then reimplement it slice-by-slice. Everything in this folder is something a future contributor (or a future version of us) might pick up *in addition to* — or *instead of* — that main path, when there's appetite, budget, or curiosity for it.

## What counts as a side quest?

A side quest meets at least one of these criteria:

- **Doesn't block the main port.** The .NET rewrite can ship without it.
- **Lives outside the rewrite's scope.** Touches the original VB6 codebase, an entirely new product surface, integrations with third-party systems, or research spikes.
- **Is speculative.** Worth scoping but not yet committed. Some side quests will graduate into main-quest arcs (and move into `docs/arcs/`); others will stay here as long-lived "would be nice" docs; some will be abandoned with a postmortem of *why*.

If something is required for the rewrite to reach feature parity with the original, it belongs in `docs/arcs/`, not here. If it's a bug or smell discovered during reverse-engineering, it belongs in `docs/BUGS-MITIGATIONS.md`. Side quests are the **parallel ambitions** — not the homework.

## Catalog

| #   | Title                                                          | Status | One-line                                                                                          |
| --- | -------------------------------------------------------------- | ------ | ------------------------------------------------------------------------------------------------- |
| SQ1 | [Bugfix the original VB6 code](VB6-BUGFIX.md)                  | Idea   | Back-port a tight, surgical set of fixes to the frozen VB6 source so any sites still on it stop bleeding. |
| SQ2 | [LLM / chat integration in the .NET app](LLM-CHAT-INTEGRATION.md) | Idea   | Use `Microsoft.Extensions.AI` to plug Groq, OpenRouter, OpenAI, etc. into the rewritten app for data and guest insights. |
| SQ3 | [Guest self-service portal](GUEST-SELF-SERVICE-PORTAL.md)      | Idea   | A web-facing booking/check-in portal that talks to the same SQLite/EF Core core as the desktop app. |

### Status legend

- **Idea** — captured here, not yet scoped in detail. The doc is mostly *why* and *what*, light on *how*.
- **Scoped** — the doc has a concrete plan, slice list, and effort estimate. Could be picked up.
- **In Progress** — someone is actively working on it. The doc tracks current state.
- **Done** — shipped. The doc is preserved as a record (and may link to the eventual home of the work).
- **Abandoned** — dropped, with a postmortem section explaining why. We keep abandoned side quests around so the next person doesn't reinvent the same dead end.

## Adding a new side quest

1. Pick a short, kebab-case slug for the filename (`MOBILE-HOUSEKEEPING.md`, `OTA-INTEGRATION.md`).
2. Copy the structure from one of the existing side quest files. Each one should have:
   - **Concept** — the elevator pitch.
   - **Why it's a side quest** — what makes this off-critical-path.
   - **Scope** — what's in, what's explicitly out.
   - **Dependencies on the main quest** — what arcs of the main port this needs (or doesn't).
   - **Sketch of approach** — enough that someone else could pick up where the doc leaves off.
   - **Effort estimate** — order-of-magnitude (days / weeks / months) with the assumptions behind it.
   - **Risks & open questions** — the hairy bits.
3. Add a row to the catalog table above with status `Idea`.
4. If a side quest graduates into a main-quest arc, leave a stub here pointing to the new arc file in `docs/arcs/`, and update its status to `Done` (or `Scoped → moved`).

Match the project doc voice: long-form prose, named code/file citations in backticks, tables for catalogs, explicit call-outs of risks. These docs are written for the next person to pick up the thread, not as marketing copy.
