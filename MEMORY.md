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
- `AZ-900_Kahoot_MultipleChoice.xlsx` (183 q) and `AZ-900_Kahoot_TrueFalse.xlsx`
  (36 q): single-sheet, Kahoot-template files that import cleanly on the free plan.
  Replaced the earlier `AZ-900_Kahoot_Import.xlsx` whose Type-Answer sheet Kahoot
  could not import (free plan + spreadsheet importer only accept quiz questions).
- `AZ-900_Reference.xlsx`: full PDF answers, sources, original use-case answers,
  and the redundancy report (for study and verification).
- `build_kahoot_xlsx.py` kept as the reproducible generator for all three files.

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

### 2026-06-21 — Fix: Type-Answer questions would not import into Kahoot

- **User request:** When importing the .xlsx into Kahoot (free plan), none of the
  Use Cases / Type-Answer questions were recognized or added. Fix so all
  questions and answers display correctly in Kahoot.
- **Root cause (verified via Kahoot support docs):** Kahoot's spreadsheet
  importer accepts ONLY multiple-choice (quiz) questions, and "Type Answer" is a
  paid feature — so all 178 type-answer rows were ignored.
- **Actions taken:**
  - Rewrote `build_kahoot_xlsx.py` to convert every use case into a
    multiple-choice question. Correct option = exact PDF answer; distractors =
    OTHER real Azure terms from the same PDFs (no invented facts).
  - Produced two single-sheet, Kahoot-template import files:
    `AZ-900_Kahoot_MultipleChoice.xlsx` (183 q) and
    `AZ-900_Kahoot_TrueFalse.xlsx` (36 q) — 219 importable items total.
  - Added `AZ-900_Reference.xlsx` (full answers, sources, original use-case
    answers, redundancy report). Removed the old broken `AZ-900_Kahoot_Import.xlsx`.
  - Enforced importer limits: question ≤95 chars, answer ≤60 chars, time in
    {5,10,20,30,60,120}; correct-answer position is rotated so it is not always A.
- **Decisions:** Split into two files so each kahoot stays under the 200-question
  limit; distractors are real PDF terms, never fabricated facts.
- **Next steps:** User imports each file as its own kahoot via
  "Import questions from spreadsheet".

### 2026-06-21 — Split multiple-choice into 5 files of 40 (free-plan time limit)

- **User request:** On the Kahoot free plan a live session must finish in under
  ~15 minutes, so split all Multiple Choice content into 5 quiz files of 40
  questions each, as a continuation until all are covered.
- **Actions taken:**
  - Extended `build_kahoot_xlsx.py` with `build_mc_split_files(40)`.
  - Generated `AZ-900_Kahoot_MC_Part1_of_5.xlsx` ... `Part5_of_5.xlsx`
    (40/40/40/40/23 = 183 questions). Same Kahoot template + option rotation
    as the full file.
  - Verified the 5 splits equal the full 183-question set exactly (order
    preserved, no duplicates/gaps, 0 validation issues).
- **Decisions:** Parts 1-4 have 40 questions; Part 5 has the remaining 23.
  Kept the combined `AZ-900_Kahoot_MultipleChoice.xlsx` as the master copy.
- **Next steps:** User imports each Part file as a separate kahoot, played in
  sequence.

### 2026-06-21 — Per-question images for the 5 MC files

- **User request:** Add an appropriate photo per question slide (relating to the
  answer) for the 5 Multiple Choice Kahoot files.
- **Key finding (verified via Kahoot docs):** Kahoot's spreadsheet importer
  CANNOT carry images — there is no image column and embedded images are ignored.
  Images can only be added in the editor (upload or built-in image library).
- **Actions taken:**
  - Added `generate_kahoot_images.py`, which generates one PNG per MC question
    (183 total) themed to the question's SUBJECT AREA (Storage, Networking,
    Security, etc.) using drawn icons — generated, so no licensing issues, and
    it doesn't spoil the answer on screen.
  - Output organized as `kahoot_images/Part1..5/PartX_Qnn.png`, matching the
    import file order 1:1.
  - Added `AZ-900_Kahoot_Image_Guide.xlsx` (READ ME + 5 part sheets) mapping each
    question to its image file AND a Kahoot image-library search keyword (so the
    user can alternatively pick a real stock photo in the editor).
- **Decisions:** Images reflect subject area, not the literal answer, to avoid
  spoilers; user adds them via drag-drop (Option A) or Kahoot's library (Option B).
- **Next steps:** User drags each PartX_Qnn.png into the matching question, or
  uses the keyword to pick a library image.

### 2026-06-21 — Convert True/False category to Yes/No + add its images

- **User request:** For the True/False statements, choose Yes if true else No.
  Replace True->Yes and False->No (rename the category to Yes/No). Also add a
  photo per question slide for this category.
- **Actions taken:**
  - Updated `build_kahoot_xlsx.py` so the statements file uses answer options
    "Yes"/"No" (Yes = statement true = option 1, No = false = option 2) and
    saves as `AZ-900_Kahoot_YesNo.xlsx`. Removed the old
    `AZ-900_Kahoot_TrueFalse.xlsx`. Updated the Reference workbook (READ ME +
    "Yes-No (with source)" sheet).
  - Extended `generate_kahoot_images.py` to also create 36 subject-area-themed
    images at `kahoot_images/YesNo/YesNo_Qnn.png` and added a "Yes-No" sheet to
    `AZ-900_Kahoot_Image_Guide.xlsx`.
  - Verified: 36 Yes/No rows, options Yes/No, correct mapping 0 issues; 36 images.
- **Decisions:** Statement wording is unchanged (only the answer options changed
  from True/False to Yes/No); images follow the same subject-area approach as MC.
- **Next steps:** User imports `AZ-900_Kahoot_YesNo.xlsx` as a kahoot and drags
  in the YesNo_Qnn.png images (or uses the guide's search keywords).

### 2026-06-21 — Zip the kahoot images for one-click download

- **User request:** Compress the kahoot_images folder into a single zip to avoid
  downloading each image/folder individually.
- **Actions taken:** Created `AZ-900_Kahoot_Images.zip` (4.6 MB) containing all
  219 images (Part1-5 = 183 MC images + YesNo = 36), preserving the folder
  structure. Committed to the repo so it can be downloaded as one file.
- **Next steps:** Download the zip from the repo, extract, and drag images into
  Kahoot. Re-run `zip -r AZ-900_Kahoot_Images.zip kahoot_images` if images change.

<!--
TEMPLATE — copy for each new session:

### YYYY-MM-DD — <short title>

- **User request:** ...
- **Context found:** ...
- **Actions taken:** ...
- **Decisions:** ...
- **Next steps:** ...
-->
