# SQ2 — LLM / chat integration in the .NET app

**Status:** Idea
**Owner:** unassigned
**Last updated:** 2026-05-04

## Concept

Add a chat panel to the rewritten Avalonia desktop app that lets front-desk and management users ask natural-language questions about their data — bookings, occupancy, guests, revenue — and get back grounded, citation-backed answers. Pluggable LLM provider: bring-your-own-key for **Groq, OpenRouter, OpenAI, Anthropic, Azure OpenAI, Ollama (local)**, etc., via the **`Microsoft.Extensions.AI`** abstraction so the application code stays provider-agnostic.

Examples of the kind of question the chat should answer well:

- "Which corporate guests stayed more than 5 nights last quarter and haven't returned?"
- "Show me revenue by room type for April vs March, and tell me what changed."
- "Draft a follow-up email to the guest in Room 412 — they're checking out tomorrow and had a complaint about the AC last night."
- "Why is occupancy down this week compared to last year?"
- "Find duplicate customer records — same phone number, different names."

## Why it's a side quest

The main quest's job is feature parity with the original VB6 system. The original has no AI surface, no chat panel, and no natural-language query interface — adding one is a *new* product capability, not a port. It's also a research-flavored workstream: provider quality, latency, and cost characteristics shift quickly, the right system prompt is unknown, and the safety story (especially around guest PII) needs design work. None of that is on the critical path to shipping the rewrite.

That said, this is the side quest most likely to graduate into a main-quest arc once the base port is stable and there's a real desktop app to attach a chat panel to.

## Scope

### In scope

- A chat panel in the Avalonia Presentation layer that lives alongside `frmDashboard`'s replacement (likely a docked side panel or a hotkeyed overlay).
- Provider plumbing using `Microsoft.Extensions.AI`'s `IChatClient` abstraction, so swapping Groq for OpenAI is a config-file change, not a code change.
- A **read-only** function-calling / tool-use surface over the Application layer: the model can call typed methods like `GetOccupancyByDateRange`, `FindCustomersByName`, `GetBookingHistoryForCustomer` — never raw SQL, never write operations. This sidesteps the entire "LLM hallucinates an `UPDATE` statement" failure mode.
- **Citation/grounding.** Every answer includes the rows/queries it was based on, rendered inline so the user can audit before acting.
- **Per-user opt-in.** The chat panel is off by default and a per-user setting in the new `User` table, gated by the same `ModuleAccess` permission matrix the rest of the app uses.
- **Local model option (Ollama).** Single-tenant deployments where guest PII can't leave the machine should be able to run a local model and get a degraded-but-functional experience.

### Out of scope

- Customer-facing chatbots (a guest-facing concierge would be its own side quest — see `GUEST-SELF-SERVICE-PORTAL.md`).
- Voice / speech input.
- Multi-modal (image) input.
- Training or fine-tuning a custom model.
- Any **write** path through the LLM. The model proposes; the user clicks a real button to commit. This is a hard rule, not a phase-1 limitation.
- LLM-driven report generation. (Tempting, but Crystal Reports' replacement is its own arc.)

## Dependencies on the main quest

This side quest needs the rewrite to have reached a specific point before it can start meaningfully:

| Required from main quest                                                              | Why                                                                          |
| -------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Application layer with typed query methods (Arcs 4–6 in `MIGRATION-DOTNET.md`)         | These are the tools the LLM is allowed to call. Without them, no grounding.  |
| Avalonia Presentation shell with a docking surface (Arc 7 onward)                      | Need somewhere to put the chat panel.                                        |
| `User` table + permission matrix port (Arc 3, Auth)                                    | Need somewhere to attach the per-user opt-in flag and audit who used the LLM. |
| Serilog wired up                                                                       | Every prompt and response goes to a structured audit log. Non-negotiable for PII compliance. |

Earliest sensible start: **after Arc 7 (Login/Dashboard)** lands. Before that, there's no shell to attach a chat to.

## Sketch of approach

### Architecture

```
+---------------------------------------+
|   Avalonia chat panel (Presentation)  |
|   - prompt input                      |
|   - streaming response renderer       |
|   - citation chips → Application      |
+-----------------+---------------------+
                  |
                  v
+---------------------------------------+
|   Application.AI                      |
|   - IChatClient (Microsoft.Extensions.AI)
|   - ToolRegistry (typed read-only fns)|
|   - SystemPromptBuilder               |
|   - PiiRedactor                       |
+-----------------+---------------------+
                  |
                  v
+---------------------------------------+
|   Application (existing query layer)  |
|   GetOccupancyByDateRange, etc.       |
+---------------------------------------+
```

The new project is `Application.AI` — a thin layer that depends on `Application` for tools, on `Microsoft.Extensions.AI` for the model abstraction, and on Serilog for audit. No reach-around from `Application.AI` into Domain or Infrastructure.

### Provider plumbing

`Microsoft.Extensions.AI` ships an `IChatClient` interface and provider adapters land as separate NuGet packages — `OpenAI`-compatible adapters cover OpenAI, Azure OpenAI, Groq (OpenAI-compatible endpoint), and OpenRouter (OpenAI-compatible endpoint with a different base URL and routing header) with the same client. Anthropic and Ollama have their own adapters. Concretely:

```csharp
// pseudo-code, exact API surface to be confirmed against the current
// Microsoft.Extensions.AI release at implementation time
IChatClient client = config.Provider switch
{
    "openai"     => new OpenAIClient(config.ApiKey).AsChatClient(config.Model),
    "groq"       => new OpenAIClient(config.ApiKey, new() { Endpoint = "https://api.groq.com/openai/v1" }).AsChatClient(config.Model),
    "openrouter" => new OpenAIClient(config.ApiKey, new() { Endpoint = "https://openrouter.ai/api/v1" }).AsChatClient(config.Model),
    "azure"      => new AzureOpenAIClient(...).AsChatClient(config.Deployment),
    "anthropic"  => new AnthropicClient(config.ApiKey).AsChatClient(config.Model),
    "ollama"     => new OllamaChatClient(config.Endpoint, config.Model),
    _            => throw new ConfigException(...)
};

// then wrap with cross-cutting concerns
client = new ChatClientBuilder(client)
    .UseLogging(loggerFactory)            // Serilog audit
    .UseFunctionInvocation()              // tool calling
    .UseDistributedCache(memoryCache)     // optional, for repeated queries
    .Build();
```

The configuration lives in user-scoped settings (so the customer can BYO key without rebuilding) and is encrypted at rest using the same DPAPI-based pattern as the main app's other secrets.

### Tool surface

The Application layer exposes a `ToolRegistry` of typed methods marked with `[AIFunction]` (or the equivalent attribute in the current `Microsoft.Extensions.AI` release). Each tool:

- Has a strongly-typed parameter record (for schema generation).
- Returns a strongly-typed result record (so the Presentation layer can render citations from real objects, not strings).
- Is **read-only**. No tool mutates state.
- Logs to Serilog with the user, prompt ID, parameters, and row count returned.

Initial tool list (deliberately small — start with the questions actually asked):

- `GetOccupancyByDateRange(DateOnly start, DateOnly end, RoomType? type)`
- `GetRevenueByDateRange(DateOnly start, DateOnly end, GroupBy granularity)`
- `FindCustomersByName(string query, int limit)`
- `FindCustomersByPhoneOrEmail(string query, int limit)`
- `GetBookingHistoryForCustomer(CustomerId id)`
- `GetCurrentInHouseGuests()`
- `GetUpcomingArrivals(DateOnly date)`
- `GetRoomStatusSnapshot()`

Tools are added as questions surface in real use, not pre-emptively.

### System prompt + grounding

The system prompt establishes:

- The model is a hotel operations assistant with access to *only* the listed tools.
- It must call a tool before making any factual claim.
- It must return a citation block (which tool was called, with what parameters) inline with the answer.
- It must refuse to answer questions outside the hotel domain or that require write access.
- Guest PII (full names, phone numbers, IDs) is included in tool results but the model is instructed not to volunteer PII unless directly asked, and the UI redacts it on display by default with a click-to-reveal pattern.

### Safety / PII

- **Prompt + completion logging.** Every interaction goes to the `LogAIQuery` table (new) with user, timestamp, prompt, tools invoked, completion. Same audit posture as the existing `LogBooking` / `LogRoom`.
- **Provider allowlist.** The customer admin picks one provider per deployment. Mixing providers (e.g., OpenAI for chat + Groq for fast follow-ups) is a v2 problem.
- **Data residency disclosure.** The settings panel shows the current provider's data-handling policy (or a link to it) and a clear "your data is sent to {provider}" line. Ollama is the no-egress option.
- **No tool may return raw `User` rows.** Even read-only — the password hash and salt columns aren't going anywhere near a model.

## Effort estimate

Order of magnitude: **3–6 weeks** for a useful first version, assuming the main quest has already delivered the Application layer, Avalonia shell, and auth.

- Week 1: Project scaffold (`Application.AI`), provider plumbing, two tools, end-to-end test that asks "How many guests checked in yesterday?" and gets a grounded answer.
- Week 2: Avalonia chat panel, streaming response rendering, citation chips.
- Week 3: Tool surface fleshed out (~8 tools), system prompt iteration.
- Week 4: PII redaction, audit logging, settings UI, provider switching tested across at least Groq + OpenAI + Ollama.
- Week 5–6: Slack-test with real users, fix the questions it gets wrong, write the user-facing docs.

## Risks & open questions

- **`Microsoft.Extensions.AI` is a moving target.** The API surface has been evolving across .NET 9 / 10 releases. Pin a version and budget for a re-rev when an LTS lands. The abstraction is the right bet — but it is not yet stable in the way `IServiceCollection` is.
- **Function-calling fidelity varies wildly by provider.** Frontier models (recent OpenAI, Anthropic, Gemini) handle multi-turn tool use cleanly. Smaller open-weight models hosted on Groq/OpenRouter are inconsistent. The system needs to gracefully degrade — if the chosen model can't do tool use, the chat panel should visibly say so rather than hallucinate.
- **Latency vs cost vs quality is a per-deployment tradeoff.** A small hotel running Groq's `llama-3.x-8b` cheaply might be perfectly happy. A boutique chain auditing high-value guests probably wants Claude Opus on Anthropic. Don't pick a default; let the customer pick.
- **PII leaving the building is a real legal question.** This may need a DPA / data-processor agreement with the chosen provider depending on jurisdiction. The Ollama path exists specifically so customers who can't sign one have a way to use the feature at all.
- **Hallucinated columns / hallucinated rooms.** Even with grounding, models will sometimes invent. The UI must always render the citation, and answers without citations should be visually flagged as "ungrounded — verify before acting."
- **Open question:** does this side quest also become the *replacement* for some of the report-builder UX in the original? Several reports in the `Report` table are basically saved queries with date filters — a natural-language query interface could subsume them. Worth deciding before sinking effort into modernizing all 30+ reports as Crystal-Reports replacements.
- **Open question:** should the chat panel also see the user's **own** activity (last booking they edited, last customer they searched for) as ambient context, or stay strictly on user-typed prompts? Ambient context makes the assistant feel more useful but expands the PII-egress surface.
