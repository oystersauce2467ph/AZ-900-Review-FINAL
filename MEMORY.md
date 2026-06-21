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
- Three source reviewer PDFs uploaded: `AZ-900 Fundamentals Reviewer.pdf` (P1),
  `AZ900_Practice_Exam_With_Answers.pdf` (P2), and
  `RADOVAN_AZ-900_ Microsoft Azure Fundamentals Reviewer.pdf` (P3).
- `AZ-900_Kahoot_Import.xlsx` generated from the 3 PDFs: a Kahoot-ready question
  bank (219 items) classified into Multiple Choice, True/False, and Type-Answer.
- `build_kahoot_xlsx.py` kept as the reproducible generator for the workbook.

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

### 2026-06-21 — Built Kahoot question bank from the 3 reviewer PDFs

- **User request:** Analyze the 3 PDF reviewers, extract all data in extreme
  detail, and put it into an .xlsx for importing into Kahoot. Classify items
  into multiple choice (4 lettered options), True/False, and type-answer use
  cases (user types the answers). Flag redundant questions. Do NOT hallucinate
  or invent questions — everything must come strictly from the PDFs.
- **Context found:** Three PDFs were pushed to `main` (P1 fundamentals notes,
  P2 practice exam Q1–Q36 with answers, P3 RADOVAN domain + practice questions).
  Extracted text from all three with PyMuPDF (no OCR needed).
- **Actions taken:**
  - Wrote `build_kahoot_xlsx.py` and generated `AZ-900_Kahoot_Import.xlsx`.
  - Workbook has 6 sheets: READ ME, Multiple Choice (5), True or False (36),
    Use Cases / Type Answer (178), Redundant Questions (25), Source Map.
  - 219 total quiz items, all cited back to a specific PDF; no invented
    distractors or facts. Kahoot char limits respected (0 violations).
- **Decisions:** Multiple choice uses ONLY options that exist in the PDFs (so
  it is intentionally small — 5 items); all Yes/No statements from P2 became
  True/False; scenarios/definitions/matching became Type-Answer with the
  PDF-sourced answer provided plus a suggested <=20-char short answer.
- **Next steps:** User to import tabs into Kahoot (MC + T/F via spreadsheet
  importer; Type-Answer added manually). Could optionally split the workbook
  per Kahoot template if the importer requires the exact official layout.

<!--
TEMPLATE — copy for each new session:

### YYYY-MM-DD — <short title>

- **User request:** ...
- **Context found:** ...
- **Actions taken:** ...
- **Decisions:** ...
- **Next steps:** ...
-->
