# ArchiveLens 🔍

ArchiveLens is a local-first AI-powered email intelligence platform that transforms Outlook PST archives into searchable project history, technical documentation, relationship graphs, and organizational knowledge.

Instead of uploading multi-gigabyte PST files to cloud services, ArchiveLens uses a local SQLite index and a two-stage Retrieval-Augmented Generation (RAG) pipeline to query and analyze your emails locally and securely.


It is designed to sit alongside other practical engineering utilities like **JobTrackerSync**.

## Features

- **Local PST Parsing**: Traverses your PST folders using Windows Outlook COM integration (no heavy C/C++ compilation required).
- **SQLite Database Index**: Fast incremental indexing that skips already-parsed items and deduplicates emails using `Message-ID` (with fallback to `subject` + `sender` + `received_time` for internal or draft emails). Supports lightning-fast queries.
- **Rich CLI UI**: Features beautiful formatting, tables, highlight matching, and progress bars powered by `rich`.
- **RAG Questions (Claude Integration)**: Ask complex, natural language questions (e.g., summarizing work, extraction, list compiling) using Claude. It intelligently converts your questions to SQL queries, pulls matching emails, and synthesizes clean markdown reports.

---

## Installation & Setup

### 1. Prerequisites
- **OS**: Windows (with Microsoft Outlook installed).
- **Python**: Python 3.9 or higher.

### 2. Install Dependencies
Install the required packages using pip:
```bash
pip install -r requirements.txt
```

### 3. API Key Setup
If you want to use the natural language questioning (`ask`) feature, you need to set the `ANTHROPIC_API_KEY` environment variable:
```cmd
set ANTHROPIC_API_KEY=your_claude_api_key_here
```

---

## Usage Guide

### 1. Indexing a PST file
To index a PST file (e.g., `archive.pst`), run:
```bash
python archive_lens.py index "C:\path\to\your\archive.pst"
```
*Note: This will open a MAPI session in Outlook, attach the store, read the folders, insert/update the database (`archive.db`), and clean up by removing the store.*

### 2. Indexing a document folder
To index a folder containing non-PST files (e.g., `.txt`, `.md`, `.pdf`, `.docx`, `.py`, `.js`, `.json`, `.csv`, `.cs`, etc.) recursively, run:
```bash
python archive_lens.py index-folder "C:\path\to\your\documents"
```
*Note: This parses text files directly, handles PDFs (via `pypdf` if installed), and Word documents (via `docx` if installed) using an incremental index that skips unchanged files based on modified timestamp and size.*

### 3. Display Statistics
To inspect database size, indexed email folders, document count, and file types:
```bash
python archive_lens.py stats
```

### 4. List Email Folders
To print a table of all indexed Outlook folders and their email count:
```bash
python archive_lens.py folders
```

### 5. Basic Keyword Search
To run a fast, case-insensitive keyword search across both emails and documents:
```bash
python archive_lens.py search "Conde Nast"
```
To filter by an Outlook email folder or limit results:
```bash
python archive_lens.py search "Power BI" --limit 10 --folder "Sent Items"
```

### 6. Natural Language Questions (RAG)
You can ask complex questions directly. ArchiveLens will draft a SQLite query targeting emails, documents, or both, retrieve matching records, and synthesize a polished response via Claude:
```bash
python archive_lens.py ask "Show me every email related to Conde Nast."
python archive_lens.py ask "Find all discussions about Power BI."
python archive_lens.py ask "Summarize everything I did for Reuters."
python archive_lens.py ask "List every compliment I received from management."
python archive_lens.py ask "Extract every architecture discussion involving MediaSphere."
```


---

## Configuration

You can customize ArchiveLens behavior in `config.json`:
- `database_path`: Location of the SQLite database.
- `ignored_folders`: Outlook folder names to exclude during indexing (e.g., Junk Email, Deleted Items).
- `max_body_char_limit`: Maximum characters to index per email body to keep the index size optimal.
- `default_claude_model`: The model used for answering questions (`claude-3-5-sonnet-20241022` by default).

---

## Ingestion Pipeline & Roadmap

ArchiveLens establishes a robust, local ingestion and analysis workflow:

```
Outlook PST ──> MAPI ──> Traverse Folders ──> Extract Messages ──> Normalize ──> SQLite Index ──> Future Search / AI
```

### Immediate Wins
- **Full-Text Search**: Query nearly three decades of email locally and quickly.
- **Resilient Search**: Find conversations even when Outlook's desktop search fails.
- **Multi-PST Support**: Index and query multiple PST archives in a single consolidated database.
- **Decoupled Architecture**: Build custom views, metrics, and automation tools independent of the Outlook desktop interface.

### Future Capabilities
- **Conversation Reconstruction**: Trace threads and reconstruct dialogue trees using `Message-ID` and reply headers.
- **People Explorer**: Query communication history graphs (e.g., *"Show me everyone I emailed in 2007"*).
- **Timeline View**: Chronologically group and filter historical context (*"What was happening in March 2014?"*).
- **Cross-PST Deduplication**: Advanced duplicate detection across separate historical PST archives.
- **Deep Knowledge RAG**: Complex local synthesis (e.g., *"What do I know about Company X?"*).

