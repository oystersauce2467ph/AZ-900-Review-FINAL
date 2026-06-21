# MEMORY.md

> **Persistent memory & chat-history log for the AZ-900-Review-FINAL repository.**
>
> This file is the project's long-term memory. AI agents read it at the start of
> every session to recover context, and append a new entry to the **Session Log**
> at the end of every session. See `AGENTS.md` → "Memory & Chat-History Protocol"
> for the rules.
>
> **Append-only:** never delete or rewrite past Session Log entries. Only the
> "Current State" and "Open Items" sections below may be edited in place.

---

## Project Snapshot

- **Repository:** `AZ-900-Review-FINAL`
- **Purpose:** Study and review notes for the Microsoft Azure Fundamentals
  (AZ-900) certification — a cloud-computing fundamentals exam.
- **Format:** Markdown study material.

---

## Current State

_A living summary of where the project stands. Edit this section in place as things change._

- Repository initialized with a README and project documentation.
- `AGENTS.md` and `MEMORY.md` added to define agent behavior and persistent memory.
- No study content (topic notes / practice questions) added yet.

---

## Open Items

_Outstanding questions and follow-up tasks. Edit in place; remove items once done._

- [ ] Add core study notes by AZ-900 domain (Cloud Concepts, Architecture &
      Services, Management & Governance).
- [ ] Decide on a folder/file structure for topics and practice questions.

---

## Session Log

_Append-only. Add a new entry below for each session. Newest entries go at the bottom._

### 2026-06-21 — Initial setup of agent memory

- **User request:** In this cloud-computing GitHub repository, generate an
  appropriate `AGENTS.md` and `MEMORY.md` that continuously remembers and
  updates our chat history in the repo.
- **Context found:** Repository contained only `README.md` with the project
  title. It is an AZ-900 (Azure Fundamentals) review repo.
- **Actions taken:**
  - Created `AGENTS.md` describing the project, conventions, working style, and
    a Memory & Chat-History Protocol that requires agents to read and update
    this file each session.
  - Created `MEMORY.md` (this file) with a project snapshot, a living
    "Current State" summary, an "Open Items" list, and this append-only
    Session Log.
- **Decisions:** `MEMORY.md` is append-only for the Session Log; only the
  "Current State" and "Open Items" sections may be edited in place.
- **Next steps:** Begin adding AZ-900 study notes; agents should append a new
  Session Log entry after each future conversation.

<!--
TEMPLATE — copy for each new session:

### YYYY-MM-DD — <short title>

- **User request:** ...
- **Context found:** ...
- **Actions taken:** ...
- **Decisions:** ...
- **Next steps:** ...
-->
