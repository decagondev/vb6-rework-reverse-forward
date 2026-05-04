# SQ3 — Guest self-service portal

**Status:** Idea
**Owner:** unassigned
**Last updated:** 2026-05-04

## Concept

A web-facing portal where **guests themselves** — not hotel staff — can:

- Search availability and book a room.
- View / modify / cancel an existing reservation.
- Upload ID documents and complete pre-arrival check-in paperwork.
- Get an emailed receipt or invoice.
- (Optional v2) Request late checkout, housekeeping, or a specific room preference.

The portal is a separate front-end that talks to the **same** EF Core / SQLite (or SQL Server, depending on deployment) data store as the rewritten desktop app. No data duplication, no sync layer.

## Why it's a side quest

The original VB6 system has no guest-facing surface at all — every interaction is mediated by a clerk in front of `frmDashboard`. The main quest is a faithful port of *that*, which means the rewrite also has no guest-facing surface unless we explicitly add one. That's a product expansion, not a port, and so it sits here as a side quest until the base system is shipping.

It also has a meaningfully different operational profile from the desktop app: public internet exposure, anonymous traffic, payment processing, GDPR-style consent flows, email/SMS infrastructure, captchas. None of that overlaps with what the desktop port has to solve. It's better treated as a sibling product than as an arc of the main rewrite.

## Scope

### In scope

- A web app (likely **Blazor Server** or **ASP.NET Core + minimal API + a SPA frontend**, TBD) that consumes the same `Application` and `Domain` layers as the desktop app. Adding it should not require any change to the Domain layer; the whole point of the SOLID-layered architecture in `MIGRATION-DOTNET.md` is that the Application layer is reusable across hosts.
- Guest authentication: passwordless email magic links for first-time bookings, optional password for returning guests.
- A booking flow that respects the same `Room` / `RoomType` / `Booking` invariants as the desktop, including the snapshot-at-booking-time semantic that the docs flag as load-bearing in `BUGS-MITIGATIONS.md` and `BOOKING-FORM.md`.
- Payment integration via **Stripe** (or equivalent — abstracted so swap is possible). Auth/capture pattern; refunds for cancellations within policy.
- Email transactional messages (booking confirmation, reminder 24h before arrival, post-stay receipt) via SendGrid / Postmark / SES — abstracted.
- Soft-launch behind a feature flag so a property can opt in property-by-property.
- A staff-side "incoming web booking" inbox in the desktop app that lets a clerk approve, reject, or auto-approve based on configurable rules.

### Out of scope

- Replacing the desktop app. The desktop app remains the source of truth for staff workflows.
- Channel manager / OTA syndication (Booking.com, Expedia, etc.). That's its own monstrous side quest.
- Loyalty programs / points.
- Real-time chat with hotel staff (the LLM chat in SQ2 is for staff, not guests).
- Native mobile apps. Mobile-responsive web is enough for v1.

## Dependencies on the main quest

Heavier coupling than SQ2 — this side quest **cannot start meaningfully** until significant chunks of the main quest are done:

| Required from main quest                                      | Why                                                                |
| ------------------------------------------------------------- | ------------------------------------------------------------------ |
| Domain layer with `Booking`, `Room`, `Customer`, `RoomType`   | These are exactly what the portal manipulates.                     |
| Application layer with booking creation/modification commands | The portal is a thin caller into these.                            |
| Infrastructure layer with a DB that supports concurrent connections | SQLite is tight for this. Likely needs SQL Server or PostgreSQL deployment option.  |
| Auth model that distinguishes **staff** users from **guest** users | Today's `User` table is staff-only. The portal needs a parallel `Guest` identity table or a discriminator column. |
| Conflict / overbooking strategy                               | Two browsers booking the same room at the same time must be safe. The desktop app's lack of transactions for multi-step writes (flagged in `BUGS-MITIGATIONS.md`) cannot survive into the portal. |

Earliest sensible start: **after Arc 8 (Booking) lands**. Before that, there's no booking domain to call.

## Sketch of approach

### Tech choices (initial bias, not decided)

- **Blazor Server** for the portal frontend. Pros: maximum code reuse with the desktop app's C# domain — no DTO duplication, no separate API surface. Cons: WebSocket-heavy, weaker for offline or third-party embedding.
- **ASP.NET Core + minimal API + Astro/Next/React** is the alternate path. Better SEO, better embed story for hotels that want to drop the booking widget on their existing site. More code.
- **PostgreSQL** as the supported DB for portal-deployed properties. SQLite is fine for desktop-only single-tenant; the portal's concurrency profile pushes us toward something with a real connection pool. EF Core abstracts this.

### Architecture

```
                              ┌────────────────────┐
                              │  Hotel website /   │
                              │  embed widget      │
                              └─────────┬──────────┘
                                        │
                                        v
┌──────────────────────┐    ┌────────────────────┐
│ Desktop app (Avalonia)│    │ Guest portal (web) │
│  - staff users        │    │  - guest users     │
└──────────┬───────────┘    └─────────┬──────────┘
           │                          │
           v                          v
   ┌───────────────────────────────────────┐
   │   Application + Domain (shared)       │
   │   - same BookingService, etc.         │
   └───────────────────┬───────────────────┘
                       v
   ┌───────────────────────────────────────┐
   │   Infrastructure (EF Core)            │
   │   - PostgreSQL (multi-host)           │
   │   - or SQLite (desktop-only)          │
   └───────────────────────────────────────┘
```

### Concurrency / overbooking

This is the hard problem. The desktop port can get away with the original VB6's optimistic posture because there's typically one clerk at a time. The portal cannot. Options:

1. **Database-level pessimistic lock** on the room+date range for the duration of the booking transaction. Simple, but ugly in a web context where users tab away mid-flow.
2. **Optimistic concurrency** with row versions on `Booking`, retry on conflict, present a friendly "someone just booked that room" message on the unlucky path. Standard EF Core pattern.
3. **Inventory hold** — when the user picks a room, hold it for 10 minutes via a row in a `BookingHold` table; expire if the user doesn't complete checkout. Closer to how OTAs work.

Recommendation: option 3 in v1, because it matches user expectations from every other booking site they've used.

### Payments

Stripe Checkout (hosted page) for v1 — minimizes PCI scope. The portal never sees a card number. Stripe webhooks update the `Booking` status from `PendingPayment` → `Confirmed`. The desktop app's existing payment-recording paths (the ones in `frmBooking` and adjacent) get a new "external/Stripe" payment-method enum value.

### Auth

- **First-time guest:** enters email + booking dates, gets a magic link, completes booking. After confirmation, can set a password if they want a permanent account.
- **Returning guest:** logs in with email/password or magic link.
- The `Guest` identity is **separate** from the `Customer` row in the Domain. A `Guest` may be linked to one or more `Customer` rows after a clerk reconciliation step (different addresses, family members, etc.). Don't try to unify them automatically — the desktop app's existing duplicate-customer problem (flagged in `CUSTOMER-FORM.md` notes) means we can't trust auto-match.

### Staff inbox

In the desktop app: a new sub-form under `frmDashboard` showing incoming web bookings, with approve/reject actions. Configurable auto-approve rules (e.g., "auto-approve if guest has a prior completed stay and total < $X") so staff aren't drowning in clicks for routine bookings.

## Effort estimate

Order of magnitude: **2–3 months** of focused work after the main quest's Arc 8 lands. This is the largest of the three side quests in this folder. Roughly:

- Weeks 1–2: Architecture + tech-stack decision, project scaffold, deploy pipeline.
- Weeks 3–4: Auth (guest identity, magic links).
- Weeks 5–7: Booking flow (search, hold, checkout, confirm).
- Weeks 8–9: Payments (Stripe integration, webhook handling).
- Week 10: Staff inbox in desktop app.
- Weeks 11–12: Email transactional templates, GDPR consent UX, polish, soft-launch checklist.

## Risks & open questions

- **Public internet exposure.** Everything in `BUGS-MITIGATIONS.md` about SQL injection, weak password storage, and missing rate limiting that's *uncomfortable* in a desktop LAN context becomes *catastrophic* on the open internet. The portal cannot inherit any of those weaknesses — it has to be hardened from day one. Concretely: the password storage rewrite (PBKDF2 / Argon2id) must be done before any guest auth is built.
- **PCI scope.** Even with Stripe Checkout offloading the card number, the moment the portal stores a Stripe `customer_id` and links it to a guest, the deployment is in PCI's "merchant" tier. Document the scope clearly so customers know what they're signing up for.
- **GDPR / privacy.** Storing IDs uploaded by guests adds a sensitive data store. Need explicit retention policy, encryption at rest, and a guest-facing "delete my data" flow.
- **The desktop app's data model isn't multi-host-ready.** Several places assume single-writer semantics (the global `gstrSQL` builder is the most obvious; the lack of transactions on multi-step writes is the most dangerous). The rewrite is supposed to fix these, but this side quest will discover any places the rewrite *didn't* fix them.
- **Open question:** is this a feature of the StarHotel rewrite, or a separate product line? If separate, it lives in its own repo and the shared layers (`Domain`, `Application`) get extracted into a NuGet package. That's an architectural choice that affects the main quest's repo layout.
- **Open question:** do we charge extra for the portal? Affects whether it's behind a feature flag (yes) or a license check (also yes, but a more annoying yes).
