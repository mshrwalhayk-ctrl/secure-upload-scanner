# Smile AI — an Enterprise Knowledge Operating System

> Not a chatbot. Not an ERP. Not an "AI assistant."
> A cognitive, multi-layer platform that turns a company's scattered files into **verified, explainable knowledge** — with evidence, audit, and human governance at its core.

## The thesis: *System beats Model*

The real intelligence isn't in the model — it's in the **system** that grounds it. Hand a frontier model a spreadsheet and it's brilliant, but it *doesn't know your company*. This project makes the company's own data a governed knowledge graph the system reads, verifies, and reasons over — so the answer is not just fluent, it's **grounded, cited, and auditable**.

**Distributed Intelligence.** The models (and a deterministic Python engine) are *replaceable thinking tools* behind a single gateway. Swap the model next year — the knowledge, the entities, the evidence, and the audit trail all stay. *The value lives in the system, not the model.*

## Principles

- **Evidence First** — nothing is trusted, not even the AI, until it is verified.
- **Contract First** — a stable Internal Core API; consumers never touch internals.
- **Python computes · AI reasons · Evidence verifies · a human approves.**
- **Single Source of Truth** — one company knowledge base. Every update (from an ingested file *or* from chat) passes the *same* governance gate: `Proposal → Approval → Knowledge → Audit → Version`. There is no other path.
- **Never Build Twice · No Vendor Lock-in · No Silent Failure.**
- **Risk-First** — every phase begins with a Spike → Thin Slice → Gate before it is allowed to scale.
- **Reality Ceiling** — a promise it keeps 100%: a system that never fails silently, documents everything, and asks when unsure. It is honestly **not** a replacement for a human expert.

## Architecture (in layers)

- **Frozen Core** — Gateway · Router · Memory · Evidence Engine · Python Engine · File Pipeline · Security. Adversarially audited and frozen; changed only through an evidence-gated process.
- **Internal Core API** — a small set of neutral, domain-free primitives (`ingest · compute · facts · events · verify · gate · model_call · register_capability`) behind a **Default-Deny** boundary. Everything above talks only to this contract.
- **Consumers** — built entirely on the API, never touching the core: a Knowledge Layer, a Human-Entity model, an Excel operating system, and domain experts (finance, HR, legal, …). New capabilities *extract* core logic through the API — they never duplicate it.

## Engineering stance

Test-driven throughout: atomic writes, verification-after-write, append-only audit, and executed guards that fail the suite on architectural drift. The foundation was frozen only after a multi-agent adversarial review found **no blocker**.

---

*An independent exploration of how enterprise knowledge — not raw model IQ — becomes durable, verifiable, and trustworthy.*
