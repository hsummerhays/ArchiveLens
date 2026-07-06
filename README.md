# ArchiveLens 🔍

ArchiveLens is a local-first AI-powered email intelligence platform that transforms Outlook PST archives into searchable project history, technical documentation, relationship graphs, and organizational knowledge.

Instead of uploading multi-gigabyte PST files to cloud services, ArchiveLens uses a local SQLite index and a two-stage Retrieval-Augmented Generation (RAG) pipeline to query and analyze your emails locally and securely.


It is designed to sit alongside other practical engineering utilities like **JobTrackerSync**.

## Features

- **Local PST Parsing**: Traverses your PST folders using Windows Outlook COM integration (no heavy C/C++ compilation required).
- **SQLite Database Index**: Fast incremental indexing that skips already-parsed emails and supports lightning-fast queries.
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

### 2. Display Statistics
To inspect database size, indexed folders, and dates:
```bash
python archive_lens.py stats
```

### 3. List Folders
To print a table of all indexed folders and their email count:
```bash
python archive_lens.py folders
```

### 4. Basic Keyword Search
To run a fast, case-insensitive keyword search:
```bash
python archive_lens.py search "Conde Nast"
```
To filter by a folder or limit results:
```bash
python archive_lens.py search "Power BI" --limit 10 --folder "Sent Items"
```

### 5. Natural Language Questions (RAG)
You can ask complex questions directly. ArchiveLens will draft a SQLite query, retrieve matching emails, and synthesize a polished response via Claude:
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
