# The Knowledge Ecosystem Vision 🌐

This document outlines the long-term vision of how **DailyNotes**, **ArchiveLens**, and **RootNode** operate together as a cohesive personal knowledge management ecosystem.

## The Verb Flow

Rather than treating these as separate products, the ecosystem is built around a progressive pipeline of information processing:

```
 Capture ──> Ingest ──> Understand ──> Recall ──> Create
```

1. **Capture (DailyNotes)**: Actively logging daily journal entries, engineering tasks, meeting outcomes, and fresh ideas.
2. **Ingest (ArchiveLens)**: Reading, extracting, and normalizing heterogeneous history (emails, source code, design specs, notes, spreadsheets) into a unified local SQLite datastore.
3. **Understand (RootNode)**: Parsing, summarizing, relating, and mapping information into topic nodes and communication graphs.
4. **Recall (Query & Search)**: Retrieving matching records via structured SQL search or semantic vector similarity.
5. **Create (Synthesis & Generation)**: Generating architecture specs, project timelines, onboarding materials, and onboarding summaries directly from your historical engineering context.

---

## Decoupled Architecture: Ingest as a Service

A key architectural boundary is that **ArchiveLens is a universal ingestion pipeline** and is not coupled to or owned by any single consumer. This allows multiple specialized tools to consume the same normalized SQLite index:

```
                            Outlook PST / OST ────┐
                            Local Directories ────┼──> ArchiveLens Ingestion
                            Git Commits / Logs ───┘           │
                                                              ▼
                                                        [archive.db]
                                                              │
                    ┌─────────────────┬───────────────────────┼───────────────────────┐
                    ▼                 ▼                       ▼                       ▼
               RootNode          AI Assistant          Timeline Viewer            Search UIs
         (Understand & Create)   (Interactive Q&A)   (Chronological Graphs)  (Keywords & Vectors)
```

### Downstream Consumers
- **RootNode**: Focuses purely on summarizing, mapping relationships, constructing semantic knowledge graphs, and creating fresh documents.
- **AI Assistant**: Performs conversational RAG over the archive to answer quick inquiries.
- **Timeline Viewer**: Generates chronological visual tracks of what was happening on specific projects, companies, or files over three decades.
- **DailyNotes**: References historical contexts directly when logging new notes, closing the loop between *Capture* and *Recall*.

---

## The Ultimate Showcase

By separating the ingestion (`ArchiveLens`) from the semantic analysis (`RootNode`), we build an architecture capable of answering high-level temporal and semantic questions, such as:

> *"How did the SellMedia billing logic evolve between 2006 and 2012?"*

Instead of returning a flat keyword index, the ecosystem synthesizes:
- **2006 MAPI emails** (extracting initial design decisions).
- **SQL scripts** (documenting database schema updates).
- **Daily Notes** (reflecting active debugging thoughts).
- **Source Code commits** (tracking implementation details).

