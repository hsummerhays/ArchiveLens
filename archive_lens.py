import os
import sys
import re
import json
import sqlite3
import argparse
import urllib.request
import urllib.error
from datetime import datetime

# Initialize Rich Console and styling
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, MofNCompleteColumn
from rich.panel import Panel
from rich.markdown import Markdown

# Ensure console supports UTF-8 on Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

console = Console()

def load_env():

    """Loads environment variables from a local .env file if it exists."""
    dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(dotenv_path):
        try:
            with open(dotenv_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" in line:
                        key, val = line.split("=", 1)
                        os.environ[key.strip()] = val.strip().strip("'\"")
        except Exception as e:
            console.print(f"[yellow]Warning: Could not read .env file: {str(e)}[/yellow]")

load_env()

# Load configuration
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

try:
    with open(CONFIG_PATH, "r") as f:
        config = json.load(f)
except Exception:
    config = {
        "database_path": "archive.db",
        "ignored_folders": ["Junk Email", "Deleted Items", "Conversation History", "Sync Issues", "Conflicts", "Local Failures", "Server Failures", "Drafts", "Outbox"],
        "max_body_char_limit": 100000,
        "default_claude_model": "claude-3-5-sonnet-20241022",
        "max_search_results_for_llm": 100
    }

DB_PATH = config.get("database_path", "archive.db")


def get_db_connection():
    """Establishes connection to the SQLite database."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initializes the database schema."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            entry_id TEXT PRIMARY KEY,
            message_id TEXT,
            subject TEXT,
            sender_name TEXT,
            sender_email TEXT,
            to_recipients TEXT,
            cc_recipients TEXT,
            bcc_recipients TEXT,
            received_time TEXT,
            body TEXT,
            folder_path TEXT,
            importance INTEGER,
            has_attachments INTEGER
        )
    """)
    # Add message_id if missing in legacy databases
    try:
        cursor.execute("ALTER TABLE emails ADD COLUMN message_id TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            file_path TEXT PRIMARY KEY,
            file_name TEXT,
            file_type TEXT,
            file_size INTEGER,
            modified_time TEXT,
            body TEXT
        )
    """)
    conn.commit()
    conn.close()



def call_claude(system_prompt, user_prompt):
    """Invokes Anthropic Claude using standard urllib to avoid external dependencies."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        console.print("[red]Error: ANTHROPIC_API_KEY environment variable is not set.[/red]")
        console.print("Please set it in your environment or local .env file.")
        sys.exit(1)

    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    
    data = {
        "model": config.get("default_claude_model", "claude-3-5-sonnet-20241022"),
        "max_tokens": 4000,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt}
        ]
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(data).encode("utf-8"),
        headers=headers,
        method="POST"
    )
    
    try:
        with urllib.request.urlopen(req) as response:
            res = json.loads(response.read().decode("utf-8"))
            return res["content"][0]["text"]
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        try:
            error_json = json.loads(error_body)
            error_msg = error_json.get("error", {}).get("message", error_body)
        except Exception:
            error_msg = error_body
        console.print(f"[red]Anthropic API Error ({e.code}): {error_msg}[/red]")
        sys.exit(1)
    except Exception as e:
        console.print(f"[red]Connection Error: {str(e)}[/red]")
        sys.exit(1)


def find_pst_folder(outlook_namespace, pst_filepath):
    """Finds the root folder of the added PST store."""
    abs_filepath = os.path.abspath(pst_filepath).lower()
    for store in outlook_namespace.Stores:
        try:
            if store.IsDataFileStore and store.FilePath:
                if os.path.abspath(store.FilePath).lower() == abs_filepath:
                    return store.GetRootFolder()
        except Exception:
            pass
            
    # Fallback to display name matching
    pst_name = os.path.splitext(os.path.basename(pst_filepath))[0].lower()
    for folder in outlook_namespace.Folders:
        try:
            if folder.Name.lower() == pst_name or pst_name in folder.Name.lower():
                return folder
        except Exception:
            pass
            
    return None


def clean_body(body_text):
    """Cleans up email body text."""
    if not body_text:
        return ""
    # Truncate if it exceeds limit
    limit = config.get("max_body_char_limit", 100000)
    if len(body_text) > limit:
        body_text = body_text[:limit] + "\n\n[TRUNCATED]"
    return body_text


def parse_recipients(recipients_obj):
    """Extracts recipients from Outlook recipients list."""
    recipients = []
    try:
        for r in recipients_obj:
            try:
                name = r.Name or ""
                address = r.Address or ""
                if name and address:
                    recipients.append(f"{name} <{address}>")
                elif name:
                    recipients.append(name)
                elif address:
                    recipients.append(address)
            except Exception:
                pass
    except Exception:
        pass
    return ", ".join(recipients)


def extract_text_from_file(filepath):
    """Gracefully extracts text from various document formats."""
    ext = os.path.splitext(filepath)[1].lower()
    text_extensions = {
        '.txt', '.md', '.markdown', '.py', '.js', '.json', '.csv', '.cs', '.html', '.css', 
        '.ini', '.cfg', '.xml', '.yml', '.yaml', '.sql', '.sh', '.bat', '.ps1'
    }
    if ext in text_extensions:
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        except Exception:
            return None
    elif ext == '.pdf':
        try:
            import pypdf
            reader = pypdf.PdfReader(filepath)
            text = []
            for page in reader.pages:
                text.append(page.extract_text() or "")
            return "\n".join(text)
        except ImportError:
            return None
        except Exception:
            return None
    elif ext == '.docx':
        try:
            import docx
            doc = docx.Document(filepath)
            return "\n".join([p.text for p in doc.paragraphs])
        except ImportError:
            return None
        except Exception:
            return None
    return None


def index_folder(folder_path):
    """Indexes non-PST files in the specified folder recursively."""
    if not os.path.exists(folder_path):
        console.print(f"[red]Error: Folder not found at '{folder_path}'[/red]")
        sys.exit(1)
        
    init_db()
    
    ignored_dir_names = {".git", "__pycache__", ".venv", "venv", ".idea", ".vscode", "node_modules"}
    allowed_extensions = {
        # Text/code formats
        '.txt', '.md', '.markdown', '.py', '.js', '.json', '.csv', '.cs', '.html', '.css', 
        '.ini', '.cfg', '.xml', '.yml', '.yaml', '.sql', '.sh', '.bat', '.ps1',
        # Rich formats
        '.pdf', '.docx'
    }
    
    files_to_index = []
    for root, dirs, files in os.walk(folder_path):
        dirs[:] = [d for d in dirs if d not in ignored_dir_names]
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext == '.pst':
                continue
            if ext in allowed_extensions:
                files_to_index.append(os.path.join(root, file))
                
    if not files_to_index:
        console.print("[yellow]No indexable documents found in the specified folder.[/yellow]")
        return
        
    console.print(f"[cyan]Found {len(files_to_index)} candidate files to index.[/cyan]")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    docs_indexed = 0
    docs_skipped = 0
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        console=console
    ) as progress:
        task = progress.add_task("Indexing folder...", total=len(files_to_index))
        for filepath in files_to_index:
            progress.update(task, advance=1, description=f"Scanning {os.path.basename(filepath)}")
            try:
                stat = os.stat(filepath)
                file_size = stat.st_size
                mtime_str = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                
                # Check if unmodified
                cursor.execute(
                    "SELECT modified_time, file_size FROM documents WHERE file_path = ?", 
                    (filepath,)
                )
                row = cursor.fetchone()
                if row and row["modified_time"] == mtime_str and row["file_size"] == file_size:
                    docs_skipped += 1
                    continue
                    
                text = extract_text_from_file(filepath)
                if text is None:
                    docs_skipped += 1
                    continue
                
                text = clean_body(text)
                file_name = os.path.basename(filepath)
                file_type = os.path.splitext(file_name)[1].lower()
                
                cursor.execute("""
                    INSERT OR REPLACE INTO documents (
                        file_path, file_name, file_type, file_size, modified_time, body
                    ) VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    filepath, file_name, file_type, file_size, mtime_str, text
                ))
                
                docs_indexed += 1
                if docs_indexed % 50 == 0:
                    conn.commit()
            except Exception:
                pass
                
    conn.commit()
    conn.close()
    
    console.print(Panel(
        f"[green]Document indexing completed successfully![/green]\n\n"
        f"• Files newly indexed/updated: [bold]{docs_indexed}[/bold]\n"
        f"• Files skipped (unchanged/unreadable): [bold]{docs_skipped}[/bold]\n"
        f"• Database file: [bold]{DB_PATH}[/bold]",
        title="Success"
    ))


def index_pst(pst_path):
    """Indexes emails from a PST file into the SQLite database."""
    if not os.path.exists(pst_path):
        console.print(f"[red]Error: PST file not found at '{pst_path}'[/red]")
        sys.exit(1)
        
    init_db()
    
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        console.print("[red]Error: pywin32 library is not installed. Run 'pip install pywin32'[/red]")
        sys.exit(1)
        
    pythoncom.CoInitialize()
    
    console.print(f"[cyan]Connecting to Outlook MAPI session...[/cyan]")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        console.print(f"[red]Failed to connect to Outlook: {str(e)}[/red]")
        console.print("[yellow]Please make sure Microsoft Outlook is installed and configured on your system.[/yellow]")
        sys.exit(1)

    abs_pst_path = os.path.abspath(pst_path)
    console.print(f"[cyan]Adding PST store to Outlook: [bold]{abs_pst_path}[/bold][/cyan]")
    
    try:
        outlook.AddStore(abs_pst_path)
    except Exception as e:
        console.print(f"[red]Failed to add PST store: {str(e)}[/red]")
        sys.exit(1)
        
    pst_root = None
    try:
        pst_root = find_pst_folder(outlook, abs_pst_path)
        if not pst_root:
            console.print("[red]Error: Added PST store could not be found in Outlook stores.[/red]")
            sys.exit(1)
            
        console.print(f"[green]Successfully loaded store: [bold]{pst_root.Name}[/bold][/green]")
        
        folders_to_process = []
        ignored_folders = set(config.get("ignored_folders", []))
        
        def collect_folders(folder, current_path=""):
            folder_name = folder.Name
            path = f"{current_path}/{folder_name}" if current_path else folder_name
            
            if folder_name in ignored_folders:
                return
                
            folders_to_process.append((folder, path))
            for subfolder in folder.Folders:
                collect_folders(subfolder, path)
                
        collect_folders(pst_root)
        console.print(f"[cyan]Found {len(folders_to_process)} folders to scan.[/cyan]")
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        emails_indexed = 0
        emails_skipped = 0
        
        for folder_obj, folder_path in folders_to_process:
            try:
                items = folder_obj.Items
                total_items = items.Count
                if total_items == 0:
                    continue
                    
                with Progress(
                    SpinnerColumn(),
                    TextColumn("[progress.description]{task.description}"),
                    BarColumn(),
                    MofNCompleteColumn(),
                    console=console
                ) as progress:
                    task = progress.add_task(f"Scanning {folder_path}...", total=total_items)
                    
                    for i in range(1, total_items + 1):
                        progress.update(task, advance=1)
                        try:
                            item = items.Item(i)
                            if not hasattr(item, 'EntryID'):
                                continue
                                
                            entry_id = item.EntryID
                            
                            # Check if already indexed by Entry ID
                            cursor.execute("SELECT 1 FROM emails WHERE entry_id = ?", (entry_id,))
                            if cursor.fetchone():
                                emails_skipped += 1
                                continue
                                
                            subject = getattr(item, "Subject", "")
                            sender_name = getattr(item, "SenderName", "")
                            
                            sender_email = ""
                            try:
                                sender_email = getattr(item, "SenderEmailAddress", "")
                            except Exception:
                                pass
                                
                            to_recipients = ""
                            try:
                                to_recipients = parse_recipients(item.Recipients)
                            except Exception:
                                pass
                                
                            received_time = ""
                            try:
                                dt = item.ReceivedTime
                                if dt:
                                    received_time = dt.strftime("%Y-%m-%d %H:%M:%S")
                            except Exception:
                                try:
                                    dt = item.SentOn
                                    if dt:
                                        received_time = dt.strftime("%Y-%m-%d %H:%M:%S")
                                except Exception:
                                    pass
                                    
                            # Extract Message-ID for duplicate tracking
                            message_id = ""
                            try:
                                message_id = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                            except Exception:
                                pass
                                
                            # Deduplication checks:
                            if message_id:
                                # Primary check: Message-ID
                                cursor.execute("SELECT 1 FROM emails WHERE message_id = ?", (message_id,))
                                if cursor.fetchone():
                                    emails_skipped += 1
                                    continue
                            else:
                                # Fallback check: Subject + Sender + Received Time (to catch drafts, internal, or calendar items)
                                if subject and received_time:
                                    cursor.execute(
                                        "SELECT 1 FROM emails WHERE subject = ? AND sender_name = ? AND received_time = ?",
                                        (subject, sender_name, received_time)
                                    )
                                    if cursor.fetchone():
                                        emails_skipped += 1
                                        continue
                                        
                            body = clean_body(getattr(item, "Body", ""))
                            importance = getattr(item, "Importance", 1)
                            
                            has_attachments = 0
                            try:
                                if item.Attachments and item.Attachments.Count > 0:
                                    has_attachments = 1
                            except Exception:
                                pass
                                
                            cursor.execute("""
                                INSERT OR REPLACE INTO emails (
                                    entry_id, message_id, subject, sender_name, sender_email,
                                    to_recipients, cc_recipients, bcc_recipients,
                                    received_time, body, folder_path, importance, has_attachments
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                entry_id, message_id, subject, sender_name, sender_email,
                                to_recipients, "", "",
                                received_time, body, folder_path, importance, has_attachments
                            ))
                            
                            emails_indexed += 1
                            if emails_indexed % 100 == 0:
                                conn.commit()
                                
                        except Exception:
                            pass
                            
                conn.commit()
            except Exception as e:
                console.print(f"[yellow]Skipping folder {folder_path}: {str(e)}[/yellow]")
                
        conn.commit()
        conn.close()
        
        console.print(Panel(
            f"[green]Indexing completed successfully![/green]\n\n"
            f"• Emails newly indexed: [bold]{emails_indexed}[/bold]\n"
            f"• Emails skipped: [bold]{emails_skipped}[/bold]\n"
            f"• Database file: [bold]{DB_PATH}[/bold]",
            title="Success"
        ))
        
    finally:
        if pst_root:
            console.print(f"[cyan]Removing PST store from Outlook session...[/cyan]")
            try:
                outlook.RemoveStore(pst_root)
                console.print("[green]Store removed successfully.[/green]")
            except Exception as e:
                console.print(f"[yellow]Warning: Could not remove store cleanly: {str(e)}[/yellow]")


def show_stats():
    """Displays indexing statistics."""
    if not os.path.exists(DB_PATH):
        console.print(f"[red]Database not found at '{DB_PATH}'. Run 'index' or 'index-folder' first.[/red]")
        return
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Check if tables exist
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='emails'")
    has_emails = cursor.fetchone() is not None
    
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='documents'")
    has_documents = cursor.fetchone() is not None
    
    total_emails = 0
    total_folders = 0
    min_date, max_date = "N/A", "N/A"
    
    if has_emails:
        cursor.execute("SELECT COUNT(*) FROM emails")
        total_emails = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(DISTINCT folder_path) FROM emails")
        total_folders = cursor.fetchone()[0]
        cursor.execute("SELECT MIN(received_time), MAX(received_time) FROM emails WHERE received_time IS NOT NULL AND received_time != ''")
        min_date, max_date = cursor.fetchone()
        
    total_docs = 0
    doc_types = []
    if has_documents:
        cursor.execute("SELECT COUNT(*) FROM documents")
        total_docs = cursor.fetchone()[0]
        cursor.execute("SELECT file_type, COUNT(*) as cnt FROM documents GROUP BY file_type ORDER BY cnt DESC")
        doc_types = cursor.fetchall()
        
    conn.close()
    
    email_stats_text = (
        f"• Total Indexed Emails: [bold cyan]{total_emails}[/bold cyan]\n"
        f"• Email Folders: [bold cyan]{total_folders}[/bold cyan]\n"
        f"• Email Date Range: [bold]{min_date or 'N/A'} to {max_date or 'N/A'}[/bold]"
    )
    
    doc_stats_text = (
        f"• Total Indexed Documents: [bold cyan]{total_docs}[/bold cyan]\n"
        f"• Document Types: [bold]{', '.join([f'{r[0]} ({r[1]})' for r in doc_types]) or 'None'}[/bold]"
    )
    
    console.print(Panel(
        f"[bold underline]Email Archive Stats[/bold underline]\n{email_stats_text}\n\n"
        f"[bold underline]Document Folder Stats[/bold underline]\n{doc_stats_text}\n\n"
        f"• Database Path: [bold]{DB_PATH}[/bold]",
        title="ArchiveLens Database Stats"
    ))


def run_search(query, limit=20, folder=None):
    """Searches the database for a keyword query across emails and documents."""
    if not os.path.exists(DB_PATH):
        console.print(f"[red]Database not found at '{DB_PATH}'. Run 'index' first.[/red]")
        return
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Check what tables are populated
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='emails'")
    has_emails = cursor.fetchone() is not None
    
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='documents'")
    has_documents = cursor.fetchone() is not None
    
    email_rows = []
    if has_emails:
        email_sql = "SELECT subject, sender_name, received_time, folder_path, has_attachments FROM emails WHERE 1=1"
        email_params = []
        if query:
            email_sql += " AND (subject LIKE ? OR body LIKE ? OR sender_name LIKE ? OR sender_email LIKE ?)"
            term = f"%{query}%"
            email_params.extend([term, term, term, term])
        if folder:
            email_sql += " AND folder_path LIKE ?"
            email_params.append(f"%{folder}%")
        email_sql += " ORDER BY received_time DESC LIMIT ?"
        email_params.append(limit)
        cursor.execute(email_sql, email_params)
        email_rows = cursor.fetchall()
        
    doc_rows = []
    if has_documents:
        doc_sql = "SELECT file_name, file_path, file_type, file_size, modified_time FROM documents WHERE 1=1"
        doc_params = []
        if query:
            doc_sql += " AND (file_name LIKE ? OR body LIKE ?)"
            term = f"%{query}%"
            doc_params.extend([term, term])
        doc_sql += " ORDER BY modified_time DESC LIMIT ?"
        doc_params.append(limit)
        cursor.execute(doc_sql, doc_params)
        doc_rows = cursor.fetchall()
        
    conn.close()
    
    if not email_rows and not doc_rows:
        console.print(f"[yellow]No emails or documents found matching '{query}'.[/yellow]")
        return
        
    if email_rows:
        table = Table(title=f"Matching Emails for '{query}' (showing top {len(email_rows)})")
        table.add_column("Date", style="magenta", width=19)
        table.add_column("Sender", style="cyan", max_width=25)
        table.add_column("Subject", style="green", max_width=60)
        table.add_column("Folder", style="blue")
        table.add_column("Att.", justify="center", style="yellow")
        
        for row in email_rows:
            att = "📎" if row["has_attachments"] else ""
            subj = row["subject"] or "(No Subject)"
            if query:
                subj = re.sub(f"({re.escape(query)})", r"[bold yellow]\1[/bold yellow]", subj, flags=re.IGNORECASE)
            table.add_row(row["received_time"] or "N/A", row["sender_name"] or "N/A", subj, row["folder_path"], att)
        console.print(table)
        console.print()
        
    if doc_rows:
        table = Table(title=f"Matching Documents for '{query}' (showing top {len(doc_rows)})")
        table.add_column("Modified Time", style="magenta", width=19)
        table.add_column("File Name", style="green", max_width=40)
        table.add_column("Type", style="yellow")
        table.add_column("Size (KB)", justify="right", style="cyan")
        table.add_column("Path", style="blue", max_width=60)
        
        for row in doc_rows:
            fname = row["file_name"]
            if query:
                fname = re.sub(f"({re.escape(query)})", r"[bold yellow]\1[/bold yellow]", fname, flags=re.IGNORECASE)
            size_kb = round(row["file_size"] / 1024, 1) if row["file_size"] else 0
            table.add_row(row["modified_time"] or "N/A", fname, row["file_type"] or "N/A", str(size_kb), row["file_path"])
        console.print(table)


def list_folders():
    """Lists all indexed folders."""
    if not os.path.exists(DB_PATH):
        console.print(f"[red]Database not found at '{DB_PATH}'. Run 'index' first.[/red]")
        return
        
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT folder_path, COUNT(*) as cnt FROM emails GROUP BY folder_path ORDER BY folder_path")
    rows = cursor.fetchall()
    conn.close()
    
    table = Table(title="Indexed Folders")
    table.add_column("Folder Path", style="cyan")
    table.add_column("Email Count", justify="right", style="green")
    for r in rows:
        table.add_row(r["folder_path"], str(r["cnt"]))
    console.print(table)


def run_ask(question):
    """Executes a natural language question using a RAG pipeline with Claude."""
    if not os.path.exists(DB_PATH):
        console.print(f"[red]Database not found at '{DB_PATH}'. Index something first.[/red]")
        return

    # Step 1: Request Claude to generate the SQL query to fetch relevant data
    sql_generation_system_prompt = """You are an expert SQL assistant for ArchiveLens, an intelligence tool.
Your job is to generate a read-only SQLite SELECT query that finds the data needed to answer the user's natural language question.
Do NOT output any explanations or conversational text. Output ONLY a valid JSON object matching this schema:
{
  "reasoning": "A short sentence explaining what we are searching for.",
  "sql_query": "The sqlite SELECT query."
}

Database Schema:
1. Table `emails`:
   - entry_id (TEXT): Primary key
   - subject (TEXT): Subject of the email
   - sender_name (TEXT): Display name of the sender
   - sender_email (TEXT): Email address of the sender
   - to_recipients (TEXT): Recipients in 'To' field
   - received_time (TEXT): Timestamp formatted as 'YYYY-MM-DD HH:MM:SS'
   - body (TEXT): Cleaned body text of the email
   - folder_path (TEXT): Folder path in the PST structure
   - importance (INTEGER): Importance flag
   - has_attachments (INTEGER): 1 if attachments are present, 0 otherwise

2. Table `documents`:
   - file_path (TEXT): Absolute file path (primary key)
   - file_name (TEXT): File name
   - file_type (TEXT): Extension of the file (e.g. '.pdf', '.txt')
   - file_size (INTEGER): File size in bytes
   - modified_time (TEXT): Timestamp formatted as 'YYYY-MM-DD HH:MM:SS'
   - body (TEXT): Extracted text content of the document

Rules:
1. ONLY return select queries. Do not select `entry_id` or other identifiers unless necessary.
2. Use CASE-INSENSITIVE filters like `LIKE '%query%'` for text matches.
3. Order search results by chronological order (e.g., `received_time DESC` or `modified_time DESC`).
4. Limit the query results to a maximum of 50 records.
5. Generate a SQLite compatible query. You can query from `emails`, `documents`, or union results if appropriate.
"""

    console.print(f"[cyan]Analyzing question and generating database search query...[/cyan]")
    sql_generation_user_prompt = f"User Question: \"{question}\""
    
    response_text = call_claude(sql_generation_system_prompt, sql_generation_user_prompt)
    
    try:
        clean_response = response_text.strip()
        if clean_response.startswith("```json"):
            clean_response = clean_response[7:]
        if clean_response.endswith("```"):
            clean_response = clean_response[:-3]
        clean_response = clean_response.strip()
        
        parsed = json.loads(clean_response)
        reasoning = parsed.get("reasoning", "")
        sql_query = parsed.get("sql_query", "")
    except Exception as e:
        console.print(f"[red]Failed to parse SQL generation response as JSON: {str(e)}[/red]")
        console.print(f"Raw response: {response_text}")
        return

    # Basic safety validation
    sql_query_upper = sql_query.upper()
    forbidden = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "CREATE", "REPLACE", "TRUNCATE"]
    if any(word in sql_query_upper for word in forbidden) or not sql_query_upper.strip().startswith("SELECT"):
        console.print(f"[red]Safety violation: Generated query is invalid or unsafe.[/red]")
        console.print(f"Query: {sql_query}")
        return

    console.print(f"[cyan]Reasoning: {reasoning}[/cyan]")
    console.print(f"[cyan]Running SQL query: [bold]{sql_query}[/bold][/cyan]")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(sql_query)
        rows = cursor.fetchall()
    except Exception as e:
        console.print(f"[red]SQL Execution Error: {str(e)}[/red]")
        conn.close()
        return
    conn.close()
    
    if not rows:
        console.print("[yellow]No relevant data found in the database.[/yellow]")
        return
        
    console.print(f"[green]Retrieved {len(rows)} results. Synthesizing answer...[/green]")
    
    # Generic key-value formatting of returned rows
    context_items = []
    for idx, row in enumerate(rows):
        row_str = f"Result #{idx+1}:\n"
        for key in row.keys():
            val = row[key]
            if key == "body" and val and len(val) > 2000:
                val = val[:2000] + "... [TRUNCATED]"
            row_str += f"{key}: {val}\n"
        row_str += "----------------------------------------"
        context_items.append(row_str)
        
    context_str = "\n".join(context_items)
    
    answer_system_prompt = """You are an advanced search and intelligence assistant.
Your goal is to answer the user's natural language question based strictly and thoroughly on the provided database results.
Synthesize the information, reference dates, files, paths, and details accurately.
Formatting: Use beautiful GitHub-flavored markdown with clean lists, headers, and bullet points. Bold important words.
"""

    answer_user_prompt = f"""Below are the search results from the local intelligence database.
Use these records to answer the user's question.

User Question: "{question}"

Search Results Context:
----------------------------------------
{context_str}
----------------------------------------

Please provide a structured, detailed answer to the user's question based on the records above.
"""

    with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
        task = progress.add_task("Summarizing results...", total=None)
        answer = call_claude(answer_system_prompt, answer_user_prompt)
        
    console.print()
    console.print(Panel(Markdown(answer), title="ArchiveLens Answer", border_style="green"))


def main():
    parser = argparse.ArgumentParser(
        description="ArchiveLens: Local-First Email & Document Intelligence Platform",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # Ingestion command (PST)
    parser_index = subparsers.add_parser("index", help="Index emails from a PST file")
    parser_index.add_argument("pst_path", help="Path to the Outlook PST file")
    
    # Ingestion command (Folder of documents)
    parser_index_folder = subparsers.add_parser("index-folder", help="Index all non-pst documents recursively from a folder")
    parser_index_folder.add_argument("folder_path", help="Path to the folder containing documents")
    
    # Stats command
    subparsers.add_parser("stats", help="Show database size and indexing stats")
    
    # Folders command
    subparsers.add_parser("folders", help="List all indexed email folders")
    
    # Search command
    parser_search = subparsers.add_parser("search", help="Perform keyword search on indexed emails and documents")
    parser_search.add_argument("query", help="Keyword or phrase to search for")
    parser_search.add_argument("--limit", type=int, default=20, help="Max results to display")
    parser_search.add_argument("--folder", help="Filter by email folder path pattern")
    
    # Ask command (LLM)
    parser_ask = subparsers.add_parser("ask", help="Ask a natural language question about emails and documents (requires ANTHROPIC_API_KEY)")
    parser_ask.add_argument("question", help="Question to ask (e.g. 'Summarize what I did for Conde Nast.')")
    
    args = parser.parse_args()
    
    if args.command == "index":
        index_pst(args.pst_path)
    elif args.command == "index-folder":
        index_folder(args.folder_path)
    elif args.command == "stats":
        show_stats()
    elif args.command == "folders":
        list_folders()
    elif args.command == "search":
        run_search(args.query, args.limit, args.folder)
    elif args.command == "ask":
        run_ask(args.question)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
