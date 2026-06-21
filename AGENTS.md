# AGENTS.md

Guidance for AI agents (Kiro, Copilot, Claude, etc.) working in this repository.

## Project Overview

This repository (`AZ-900-Review-FINAL`) is a study and review workspace for the
**Microsoft Azure Fundamentals (AZ-900)** certification exam. Its purpose is to
collect, organize, and refine cloud-computing notes, summaries, practice
questions, and quick-reference material that help prepare for the exam.

The AZ-900 exam covers the following domains:

1. **Cloud Concepts** — benefits of cloud computing, IaaS/PaaS/SaaS service
   models, public/private/hybrid deployment models, CapEx vs. OpEx, the shared
   responsibility model, high availability, scalability, elasticity, and
   disaster recovery.
2. **Azure Architecture & Services** — core architectural components (regions,
   availability zones, resource groups, subscriptions, management groups) and
   core Azure compute, networking, storage, and database services.
3. **Azure Management & Governance** — cost management, governance and
   compliance tools (Azure Policy, role-based access control, locks, tags),
   and monitoring tools.

## Repository Conventions

- All study material is written in **Markdown** (`.md`).
- Keep explanations concise and exam-focused; prefer bullet points and tables
  over long prose.
- When adding a new topic, create or extend a clearly named file (for example
  `cloud-concepts.md`, `compute-services.md`, `governance.md`).
- Use Azure's official terminology and spelling (for example "availability
  zone", "resource group", "Microsoft Entra ID").
- Cite the official Microsoft Learn source when introducing factual claims.

## Memory & Chat-History Protocol (IMPORTANT)

This repository maintains a persistent log of work in **`MEMORY.md`**. Agents
must treat it as the project's long-term memory of conversations and changes.

At the **start** of every session:

1. Read `MEMORY.md` to recover context from previous conversations.
2. Use that context to stay consistent with earlier decisions and material.

At the **end** of every session (or whenever meaningful work is completed):

1. Append a new dated entry to the **Session Log** in `MEMORY.md`.
2. Summarize what the user asked, what was decided, and what files changed.
3. Record any open questions or follow-up tasks under "Next Steps".
4. Never delete or rewrite past entries — `MEMORY.md` is append-only history.
   Only the "Current State" and "Open Items" summary sections may be edited.

## Working Style

- Confirm intent before large or destructive changes.
- Prefer small, reviewable commits with descriptive messages.
- Do not add automated tests or build tooling unless explicitly requested —
  this is a study/notes repository, not an application.
- Push work to a feature branch and open a pull request for review rather than
  committing directly to `main`.
