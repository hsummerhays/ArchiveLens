# The Knowledge Ecosystem Vision 🌐

This document outlines the long-term vision of how **DailyNotes**, **ArchiveLens**, and **RootNode** operate together as a cohesive personal knowledge management ecosystem.

```
┌────────────────────────┐
│ Capture: DailyNotes    │ <── Journal entries, active thoughts, and immediate logs
└───────────┬────────────┘
            │
            ▼
┌────────────────────────┐
│ Ingest: ArchiveLens    │ <── Normalizes files, emails, code, logs, and CSVs into a SQLite DB
└───────────┬────────────┘
            │
            ▼
┌────────────────────────┐
│ Understand: RootNode   │ <── Semantic search, topic nodes, summaries, and timeline relation
└────────────────────────┘
```

---

## Architectural Breakdown

### 1. Ingest: ArchiveLens 🔍
* **Role**: Ingestion and normalization engine.
* **Sources**: 
  - Outlook PST / OST archives
  - Local folders (PDFs, Word docs, code files, CSVs, txt, markdown)
  - *Future sources*: Gmail, OneDrive, Git commits history, Slack/Teams exports, Notion, Jira.
* **Output**: A unified, structured SQLite database (`archive.db`). 
* **Design Philosophy**: Loose coupling. ArchiveLens handles the complexity of parsing raw, heterogeneous data formats so that downstream applications do not have to.

### 2. Understand: RootNode 🌲
* **Role**: Synthesis, relationship mapping, and semantic querying.
* **Capabilities**:
  - **Hierarchical Summaries**: Dynamically roll up search contexts.
  - **Topic Nodes**: Reconstruct topics (e.g., *QuickBooks*, *SellMedia*, *Billing*) into a knowledge graph.
  - **Timeline Construction**: Cross-reference dates from emails, CSVs, commit logs, and notes to show how a project evolved over years.
  - **Long-term Memory**: Maintain historical query contexts for deep semantic search.

### 3. Capture: DailyNotes ✍️
* **Role**: The capture mechanism and entry point for fresh day-to-day knowledge.
* **Workflow**:
  1. Record notes, timings, and logs during the workday.
  2. ArchiveLens ingests these Daily Notes.
  3. RootNode synthesizes the new entries, relates them to existing topic nodes, and updates active project timelines.

---

## The Ultimate Showcase

By decoupling the ingestion (`ArchiveLens`) from the cognitive engine (`RootNode`), we build an architecture capable of answering high-level temporal and semantic questions, such as:

> *"How did the SellMedia billing logic evolve between 2006 and 2012?"*

Instead of returning a flat list of keywords, the ecosystem builds a structured timeline from:
- **2006 MAPI emails** (extracting decisions).
- **SQL scripts** (documenting database schema updates).
- **Daily Notes** (reflecting active debugging thoughts).
- **Source Code commits** (tracking implementation details).
