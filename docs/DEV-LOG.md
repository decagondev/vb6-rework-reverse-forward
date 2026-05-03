# Development Log

Append-only chronological journal of the rewrite. One entry per slice, **newest first**. Format and protocol defined in `MIGRATION-DOTNET.md` § Required documentation artifacts.

This file exists so that any future session — human or LLM — can resume work after context loss by reading the most recent 5–10 entries plus the active arc file.

**Conventions:**
- Newest entry at the top, directly under this header.
- Entries are immutable once written. Corrections go in a new entry that references the corrected one.
- Status values: `ready-for-commit` | `in-progress` | `blocked` | `merged`.
- An arc-completion entry is added at the end of each arc, before the first entry of the next arc.

**Templates** are in `MIGRATION-DOTNET.md` § Templates.

---

<!-- New entries go below this line, newest first. -->

_No entries yet. The rewrite has not started — work begins with Arc 01 (Foundation), planned in `docs/arcs/ARC-01-FOUNDATION.md`._
