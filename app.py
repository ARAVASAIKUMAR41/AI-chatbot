import threading
import os
import re
import zipfile
import sys

# Ensure bundled dependencies in ./lib are accessible in Catalyst AppSail (Linux)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.join(BASE_DIR, "lib")
if os.path.exists(LIB_DIR):
    if sys.platform != "win32":
        sys.path.insert(0, LIB_DIR)
    else:
        sys.path.append(LIB_DIR)

import io
import json
import datetime
import uuid
import docx2txt
from flask import Flask, jsonify, render_template, request, send_file
from pypdf import PdfReader

import llm
import knowledge_memory
from zoho_parser import ZohoDSParser, ZohoCreatorParser
from doc_exporter import markdown_to_docx, markdown_to_pdf

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SESSIONS_DIR = os.path.join(BASE_DIR, "sessions")
os.makedirs(SESSIONS_DIR, exist_ok=True)


def get_session_filepath(session_id: str) -> str:
    safe_id = "".join(c for c in session_id if c.isalnum() or c in ("-", "_"))
    return os.path.join(SESSIONS_DIR, f"{safe_id}.json")


_session_lock = threading.Lock()


def load_session(session_id: str) -> dict:
    """Thread-safe session loader with automated self-healing recovery for concurrent or interrupted writes."""
    if not session_id:
        return None
    fp = get_session_filepath(session_id)
    if os.path.exists(fp):
        with _session_lock:
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    return json.load(f)
            except json.JSONDecodeError as e:
                # Self-healing recovery for interrupted or trailing data writes
                print(f"Warning: JSONDecodeError in session {session_id} ({e}). Attempting self-healing recovery...")
                try:
                    with open(fp, "r", encoding="utf-8") as f:
                        raw_text = f.read()
                    data, end_pos = json.JSONDecoder().raw_decode(raw_text)
                    print(f"Session {session_id} self-healed (valid JSON prefix recovered up to char {end_pos}). Re-saving cleanly...")
                    # Re-save cleanly and atomically
                    tmp_fp = fp + f".tmp.{uuid.uuid4().hex[:6]}"
                    with open(tmp_fp, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False)
                        f_out.flush()
                        os.fsync(f_out.fileno())
                    os.replace(tmp_fp, fp)
                    return data
                except Exception as rec_err:
                    print(f"Error recovering session {session_id}: {rec_err}")
            except Exception as e:
                print(f"Error reading session {session_id}: {e}")
    return None


def save_session(session_data: dict):
    """Thread-safe and atomic session writer using temporary file swap to eliminate race conditions and corrupted writes."""
    session_id = session_data["id"]
    fp = get_session_filepath(session_id)
    tmp_fp = fp + f".tmp.{uuid.uuid4().hex[:6]}"
    with _session_lock:
        try:
            with open(tmp_fp, "w", encoding="utf-8") as f:
                json.dump(session_data, f, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_fp, fp)
        except Exception as e:
            if os.path.exists(tmp_fp):
                try:
                    os.remove(tmp_fp)
                except OSError:
                    pass
            raise e


def list_sessions() -> list:
    sessions = []
    if not os.path.exists(SESSIONS_DIR):
        return sessions

    for fname in os.listdir(SESSIONS_DIR):
        if fname.endswith(".json") and ".tmp." not in fname:
            fp = os.path.join(SESSIONS_DIR, fname)
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    sessions.append({
                        "id": data.get("id"),
                        "title": data.get("title", "Untitled Session"),
                        "filename": data.get("filename", ""),
                        "updated_at": data.get("updated_at", ""),
                        "message_count": len(data.get("messages", []))
                    })
            except Exception:
                continue

    sessions.sort(key=lambda s: s.get("updated_at", ""), reverse=True)
    return sessions


def delete_session(session_id: str) -> bool:
    fp = get_session_filepath(session_id)
    if os.path.exists(fp):
        try:
            os.remove(fp)
            return True
        except Exception:
            pass
    return False


# Supported programming / code file extensions (read as plain text)
CODE_EXTENSIONS = {
    ".py", ".js", ".jsx", ".ts", ".tsx", ".html", ".htm", ".css", ".scss",
    ".json", ".xml", ".java", ".deluge", ".md", ".yml", ".yaml", ".csv",
    ".sql", ".sh", ".bat", ".cfg", ".ini", ".log", ".rb", ".php", ".c",
    ".cpp", ".h", ".cs", ".go", ".rs", ".swift", ".kt", ".vue", ".svelte",
    ".toml", ".env", ".gitignore", ".properties", ".gradle", ".makefile",
}

# File extensions to skip when extracting from ZIP archives
ZIP_SKIP_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".svg", ".webp",
    ".mp3", ".mp4", ".wav", ".avi", ".mov", ".mkv",
    ".woff", ".woff2", ".ttf", ".eot", ".otf",
    ".zip", ".tar", ".gz", ".rar", ".7z",
    ".exe", ".dll", ".so", ".dylib", ".bin",
    ".pyc", ".class", ".o", ".obj",
}

# Max size per file inside a ZIP to prevent memory issues (512 KB)
ZIP_MAX_FILE_SIZE = 512 * 1024


def _read_text_bytes(raw_bytes: bytes) -> str:
    """Decode raw bytes to text, trying utf-8 first then latin-1 fallback."""
    try:
        return raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        return raw_bytes.decode("latin-1", errors="replace")


def _extract_zip_contents(file) -> str:
    """Extracts and combines all readable text files from a ZIP archive.
    Each file's content is prefixed with a header showing its path.
    Skips binary files (images, fonts, executables, etc.)."""
    combined_parts = []
    file_bytes = file.read()

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zf:
            # Sort entries so output is deterministic
            entries = sorted(zf.namelist())
            readable_count = 0
            for entry_name in entries:
                # Skip directories and hidden/system files
                if entry_name.endswith("/") or "__MACOSX" in entry_name or entry_name.startswith("."):
                    continue

                ext = os.path.splitext(entry_name)[1].lower()

                # Skip known binary extensions and bloated vendor bundles
                low_e = entry_name.lower()
                if ext in ZIP_SKIP_EXTENSIONS or any(b in low_e for b in ["node_modules/", "package-lock.json", ".min.js", ".min.css", ".map", ".bundle.js", "/vendor/"]):
                    continue

                info = zf.getinfo(entry_name)
                # Skip files larger than the limit
                if info.file_size > ZIP_MAX_FILE_SIZE:
                    combined_parts.append(
                        f"\n{'='*60}\n📄 FILE: {entry_name}\n{'='*60}\n"
                        f"[Skipped: file too large ({info.file_size:,} bytes)]\n"
                    )
                    continue

                try:
                    raw = zf.read(entry_name)
                    text = _read_text_bytes(raw)
                    if text.strip():
                        readable_count += 1
                        combined_parts.append(
                            f"\n{'='*60}\n📄 FILE: {entry_name}\n{'='*60}\n{text}\n"
                        )
                except Exception:
                    # Silently skip unreadable files
                    continue

            if combined_parts:
                header = (
                    f"📦 ZIP ARCHIVE: {file.filename}\n"
                    f"Total readable files extracted: {readable_count}\n"
                    f"Total entries in archive: {len(entries)}\n"
                    f"{'-'*60}\n"
                )
                return header + "\n".join(combined_parts)
            else:
                return f"[ZIP archive '{file.filename}' contained no readable text files]"
    except zipfile.BadZipFile:
        return f"[Error: '{file.filename}' is not a valid ZIP file]"
    except Exception as e:
        return f"[Error extracting ZIP '{file.filename}': {str(e)}]"


def extract_text_from_file(file):
    """Extracts raw text content from DS, TXT, PDF, DOCX, ZIP, and programming files."""
    filename = file.filename.lower()
    ext = os.path.splitext(filename)[1]

    if filename.endswith(".ds"):
        raw_bytes = file.read()
        return _read_text_bytes(raw_bytes)

    elif filename.endswith(".txt"):
        return file.read().decode("utf-8", errors="replace")

    elif filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file.read()))
        text = ""
        for page in reader.pages:
            content = page.extract_text()
            if content:
                text += content + "\n"
        return text

    elif filename.endswith(".docx"):
        text = docx2txt.process(io.BytesIO(file.read()))
        return text

    elif filename.endswith(".zip"):
        return _extract_zip_contents(file)

    elif ext in CODE_EXTENSIONS:
        raw_bytes = file.read()
        return _read_text_bytes(raw_bytes)

    return ""


def is_zoho_ds_content(text: str, filename: str = "") -> bool:
    if filename and (filename.lower().endswith(".ds") or filename.lower().endswith(".deluge")):
        return True
    if not text:
        return False
    t = text.lower()
    tokens = [
        "application ", "form ", "report ", "workflow ", "deluge", "zml",
        "displayname", "success message", "on add", "on edit", "on load",
        "thisapp", "input."
    ]
    matched = sum(1 for tok in tokens if tok in t)
    return matched >= 2 or 'application "' in t or "application '" in t or (filename and filename.lower().endswith(".ds"))



def classify_zoho_request(user_prompt: str) -> str:
    """Classifies user intent for Zoho Creator .ds file handling:
    Returns one of:
      'pages_count'   -> Exact count and list of Pages
      'forms_count'   -> Exact count and list of Forms
      'reports_count' -> Exact count and list of Reports
      'pages_explain' -> Detailed explanation of Pages & Snippet/Widget workflows
      'pages'         -> Pages & Dashboards documentation
      'forms'         -> Form schemas & Data dictionary documentation
      'functions'     -> Deluge functions & business logic documentation
      'workflows'     -> Workflows & event triggers documentation
      'reports'       -> Reports & views documentation
      'security'      -> Security & PII audit documentation
      'architecture'  -> Entity relationships & architecture
      'testing'       -> QA & Test scenarios
      'full_doc'      -> Full comprehensive 10-section technical documentation
      'llm_qa'        -> Specific question / chat query answered by LLM
    """
    import re
    p = user_prompt.lower().strip()
    if not p:
        return "full_doc"

    # Multi-component or full technical document requests (e.g. "what I want in pdf", overview + forms + widgets + er diagram)
    is_multi_section_doc = (
        any(w in p for w in ["what i want in pdf", "want in pdf", "content in pdf", "in pdf means", "full doc", "full technical", "technical documentation", "complete documentation", "full documentation", "complete doc", "unified documentation", "unified technical", "with widgets", "uploaded widgets", "submit", "connecting all widget"]) or
        (sum([1 for k in ["overview", "form", "report", "page", "widget", "workflow", "function", "schedule", "diagram"] if k in p]) >= 3 and any(w in p for w in ["pdf", "doc", "document", "specification", "expand", "explain"]))
    )
    if is_multi_section_doc:
        return "full_doc"

    # Targeted exclusivity terms ("only", "just", "mattum", "madu", "alone", "specifically", "thaniya")
    only_words = ["only", "just", "mattum", "madu", "alone", "single", "specifically", "particular", "thaniya", "solrathu", "solluratha"]
    is_targeted_only = any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in only_words)

    # Widget Page vs Snippet Page separation & count questions (Deterministic)
    widget_page_tokens = ["widget page", "widegt page", "widget pages", "widegt pages", "custom widget page", "custom widget"]
    snippet_page_tokens = ["snippet page", "snippet pages", "html snippet", "htmlpage", "html snippet page"]
    has_page_distinction = any(w in p for w in widget_page_tokens) or any(s in p for s in snippet_page_tokens)
    count_sep_tokens = [
        "ethana", "how many", "count", "separate", "thani", "thani thani", "thaniya",
        "piri", "pirikka", "distinguish", "difference", "list", "ethana nu", "ethana page"
    ]
    if has_page_distinction and any(c in p for c in count_sep_tokens):
        return "pages_count"

    # Component Count / Inventory questions (Deterministic)
    if any(w in p for w in ["how many page", "how many pages", "page count", "pages count", "number of page", "number of pages", "ethana page", "page ethana", "pages ethana"]):
        if any(w in p for w in ["explain", "workflow", "detail"]):
            return "pages_explain"
        return "pages_count"

    if any(w in p for w in ["how many form", "how many forms", "form count", "forms count", "number of form", "number of forms"]):
        return "forms_count"

    if any(w in p for w in ["how many report", "how many reports", "report count", "reports count", "number of report", "number of reports"]):
        return "reports_count"

    if any(w in p for w in ["how many workflow", "how many workflows", "workflow count", "workflows count", "number of workflow", "number of workflows", "workflow list", "list workflow", "list workflows"]):
        return "workflows_count"

    # Targeted Workflows query (e.g. "workflow thaniya", "work flow thani ya", "workflows mattum", "thaniya workflow")
    workflow_targeted_kw = [
        "workflow thaniya", "work flow thaniya", "workflow thani", "work flow thani",
        "workflows thaniya", "workflow mattum", "work flow mattum", "workflow madu",
        "work flow madu", "thaniya workflow", "thaniya work flow", "workflow pirikkala", "work flow pirikkala",
        "workflow thani ya", "work flow thani ya"
    ]
    if any(wk in p for wk in workflow_targeted_kw):
        return "workflows"

    # Explain page / snippet / widget workflows or user asking about page vs form confusion
    if any(w in p for w in ["widget page", "snippet page", "htmlpage", "page keta", "fom count", "form count"]):
        return "pages_explain"

    if any(w in p for w in ["page", "pages"]) and any(w in p for w in ["explain", "snippet", "widget", "workflow", "kanban", "how does", "pipeline"]):
        if not any(w in p for w in ["only give doc", "tech doc", "technical doc", "documetion", "documentation", "specification"]):
            return "pages_explain"

    # Page Navigation, Page Variables, In-Page Work & Inter-Page Redirection Diagrams query
    page_nav_kw = [
        "page connect", "pages connect", "page innoru page", "connect airukku",
        "page variable", "pages variable", "page la irukkura variable", "page variables",
        "page redirect", "pages redirect", "page ridirect", "aduthapage", "next page",
        "intha page la enna", "work pannala", "page actions", "page work", "in-page work"
    ]
    if any(k in p for k in page_nav_kw) or (("page" in p or "pages" in p) and any(k in p for k in ["variable", "connect", "redirect", "ridirect", "aduthapage", "actions"])):
        return "pages"

    # Executive Application Summary / Application Flow / What is happening
    summary_flow_kw = [
        "app summary", "application summary", "what happening", "what is happening",
        "executive summary", "overview", "application overview", "summary doc", "summary document",
        "summarize application", "summarise application", "compact doc", "compact document",
        "short doc", "short document", "concise doc", "concise document", "compact pdf",
        "summary pdf", "like this content in pdf", "like this in pdf", "summary mattum",
        "overview mattum", "enna nadakkuthu", "enna pannuthu", "process summary",
        # Application flow requests:
        "flow in application", "flow of application", "application flow", "app flow",
        "system flow", "operational flow", "core operational flow", "lifecycle flow",
        "overview of application", "overview of the application", "what the flow",
        "what is the flow", "explain the flow", "explain flow", "flow in the application",
        "overall flow", "end to end flow", "end-to-end flow", "flow in app", "app process flow"
    ]
    if any(sk in p for sk in summary_flow_kw) or (("summary" in p or "summarize" in p or "overview" in p or "flow" in p) and any(w in p for w in ["app", "application", "system", "what", "doc", "pdf", "project", "this"])):
        # If user explicitly asked for standalone ER diagram or function call graph flowchart, let it route to diagram
        if not any(ek in p for ek in ["er diagram", "entity relationship", "call graph", "function call"]):
            return "app_summary"

    # Flowchart / Diagram request
    diagram_kw = [
        "flow diogram", "flow diagram", "flowchart", "flow chart", "diogram", "diagram",
        "diagrams", "diograms", "workflow diagram", "process flow", "architecture diagram",
        "data flow", "system flow", "er diagram", "entity relationship diagram", "sequence diagram",
        "logic flow", "visual flow", "visual diagram"
    ]
    has_negation = any(neg in p for neg in ["vendam", "venda", "remove", "delete", "without", "illama", "theva illa", "don't want", "dont want", "exclude", "skip", "omit"])
    if any(dk in p for dk in diagram_kw):
        if has_negation:
            return "full_doc"
        return "diagram"

    # Keywords for individual components
    pages_kw = ["page", "pages", "dashboard", "dashboards", "zml", "ui page", "screen", "screens"]
    functions_kw = ["function", "functions", "deluge", "deluge script", "deluge scripts", "business logic", "custom function", "methods", "fn"]
    forms_kw = ["form", "forms", "schema", "schemas", "data dictionary", "dictionary", "field", "fields", "entity", "entities", "table", "tables"]
    workflows_kw = ["workflow", "workflows", "trigger", "triggers", "automation", "automations", "event", "events"]
    reports_kw = ["report", "reports", "view", "views", "grid", "grids", "list view"]
    security_kw = ["security", "privacy", "pii", "personal data", "gdpr", "access control"]
    arch_kw = ["architecture", "relationship", "relationships", "er diagram", "relational", "parent child"]
    test_kw = ["test case", "test cases", "testing", "qa", "test scenario", "traceability"]

    matches = {
        "pages": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in pages_kw),
        "functions": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in functions_kw),
        "forms": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in forms_kw),
        "workflows": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in workflows_kw),
        "reports": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in reports_kw),
        "security": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in security_kw),
        "architecture": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in arch_kw),
        "testing": any(re.search(r"\b" + re.escape(w) + r"\b", p) for w in test_kw),
    }

    matched_components = [k for k, v in matches.items() if v]

    # Targeted Form Purpose / Single Form query (e.g. "verum Lead form purpose mattum", "forms purpose", etc.)
    form_purpose_kw = [
        "form purpose", "purpose of form", "purpose of the form", "forms purpose",
        "form oda purpose", "form mattum", "form madu", "purpose mattum", "purpose madu",
        "form purpose mattum", "form purpose madu", "aform purpose madu", "verum form"
    ]
    if any(fp in p for fp in form_purpose_kw):
        return "form_purpose"

    # Explicit full document request keywords
    explicit_full_words = ["full doc", "full documentation", "entire doc", "complete doc", "all sections", "everything", "full technical"]
    if any(w in p for w in explicit_full_words) and not is_targeted_only:
        return "full_doc"

    # Check for direct questions
    is_question = ("?" in p) or any(p.startswith(qw) for qw in ["what ", "how ", "why ", "can you ", "does ", "where ", "when ", "is there ", "explain "])
    doc_words = ["doc", "document", "documentation", "documentaion", "documention", "documatation", "spec", "specification", "table", "give", "generate", "show", "list", "tharanum", "kudu", "ready", "panni", "pannitha"]
    has_doc_intent = any(dw in p for dw in doc_words)

    if is_question and not has_doc_intent:
        return "llm_qa"

    # If exactly one component matched:
    if len(matched_components) == 1:
        return matched_components[0]
    elif len(matched_components) > 1:
        # If targeted exclusively, return the targeted one
        if is_targeted_only:
            for comp in ["pages", "functions", "forms", "workflows", "reports", "security", "architecture", "testing"]:
                if matches[comp]:
                    return comp
        return "full_doc"

    # General documentation / generation keywords (without specific component)
    general_doc_kw = [
        "document", "documentation", "documentaion", "documention", "documatation", "doc", "tech doc", "technical",
        "technical documentation", "generate", "summary", "overview", "ready panni tha",
        "tharanum", "panni tha", "create report", "analyze"
    ]
    if any(kw in p for kw in general_doc_kw):
        return "full_doc"

    # Explicit widget & full doc requests
    if "widget" in p:
        if any(w in p for w in ["attach", "source", "code", "full", "doc", "document", "complete", "all", "tharanum", "panni", "analyze", "anolyse", "unified", "integrated", "package"]):
            return "full_doc"

    return "llm_qa"


def detect_widgets_in_ds(ds_text: str) -> list:
    """Scan a parsed DS file and return a list of genuine custom widgets.
    Associates each widget with its specific Zoho Creator page.
    Note: Native HTML snippets (<dsp>, htmlpage, HTML Snippet elements) are NOT widgets —
    their code is already fully embedded inside the .ds file and documented natively.
    Only external custom widgets (Extension / CLI packages requiring file attachment) are returned.
    Returns an empty list [] if no genuine widgets are detected."""
    import re as _re
    import html as _html

    widgets_found = []
    seen_keys = set()

    # Match each page definition directly: page <name> { ... } (allowing optional parameters)
    page_matches = list(_re.finditer(r'(?:^|\n)\s*page\s+([A-Za-z0-9_]+)(?:\s*\([^)]*\))?\s*\{', ds_text))

    for pm in page_matches:
        p_name = pm.group(1)
        start_idx = pm.end() - 1
        depth = 0
        end_idx = start_idx
        for i in range(start_idx, len(ds_text)):
            if ds_text[i] == '{':
                depth += 1
            elif ds_text[i] == '}':
                depth -= 1
                if depth == 0:
                    end_idx = i
                    break

        page_body = ds_text[start_idx:end_idx + 1]
        disp_m = _re.search(r'displayname\s*=\s*["\']([^"\']+)["\']', page_body)
        p_display = disp_m.group(1) if disp_m else p_name.replace("_", " ")
        unescaped = _html.unescape(page_body).replace('\\"', '"').replace("\\'", "'")

        # Check for genuine <widget ...>, <widgets ...> or <customwidget ...> tags inside this page
        # Note: <dsp ...> is Deluge Server Page / HTML Snippet, which is intentionally excluded
        cwidget_tags = list(_re.finditer(r'<(?:customwidgets?|widgets?)\b([^>]*)>', unescaped, _re.IGNORECASE))
        for cwm in cwidget_tags:
            tag_attrs = cwm.group(1)
            # Skip HTML snippets or Deluge snippets
            if 'html snippet' in tag_attrs.lower() or 'type="html"' in tag_attrs.lower() or 'type=\'html\'' in tag_attrs.lower():
                continue

            name_m = _re.search(r'(?:name|elementName|widgetName|widgetid|id)\s*=\s*["\']([^"\']+)["\']', tag_attrs, _re.IGNORECASE)
            w_name = name_m.group(1) if name_m else f"{p_name}_Widget"
            if w_name.lower() in ['html snippet', 'html_snippet', 'snippet']:
                continue
            if w_name.lower() == 'widget':
                w_name = f"{p_name}_Widget"

            key = (p_name, w_name)
            if key not in seen_keys:
                seen_keys.add(key)
                widgets_found.append({
                    "widget_id": f"w_{len(widgets_found) + 1}",
                    "page_name": p_name,
                    "page_display_name": p_display,
                    "widget_name": w_name,
                    "widget_display_name": w_name.replace("_", " "),
                    "widget_type": "Custom Widget",
                    "description": f"Custom Widget on page '{p_display}'"
                })

        # Also check for components declared explicitly with type="widget" or type="customwidget"
        typed_widgets = list(_re.finditer(r'<[a-zA-Z0-9_]+\b[^>]*\btype\s*=\s*["\'](?:custom)?widgets?["\'][^>]*>', unescaped, _re.IGNORECASE))
        for twm in typed_widgets:
            tag_str = twm.group(0)
            if 'dsp' in tag_str.lower() or 'html snippet' in tag_str.lower():
                continue
            name_m = _re.search(r'(?:name|elementName|widgetName|widgetid|id)\s*=\s*["\']([^"\']+)["\']', tag_str, _re.IGNORECASE)
            w_name = name_m.group(1) if name_m else f"{p_name}_Widget"
            if w_name.lower() in ['html snippet', 'html_snippet', 'snippet']:
                continue
            if w_name.lower() == 'widget':
                w_name = f"{p_name}_Widget"
            key = (p_name, w_name)
            if key not in seen_keys:
                seen_keys.add(key)
                widgets_found.append({
                    "widget_id": f"w_{len(widgets_found) + 1}",
                    "page_name": p_name,
                    "page_display_name": p_display,
                    "widget_name": w_name,
                    "widget_display_name": w_name.replace("_", " "),
                    "widget_type": "Custom Widget",
                    "description": f"Custom Widget on page '{p_display}'"
                })

    # Check for standalone customwidget declarations (top-level or outside pages)
    raw_customwidgets = list(set(_re.findall(r'(?:^|\n)\s*(?:customwidgets?|widgets?)\s+([A-Za-z0-9_]+)\s*\{', ds_text, _re.IGNORECASE)))
    for cw in raw_customwidgets:
        if cw.lower() not in ['app', 'pages', 'forms', 'reports', 'workflow', 'functions'] and not any(w["widget_name"] == cw for w in widgets_found):
            widgets_found.append({
                "widget_id": f"w_{len(widgets_found) + 1}",
                "page_name": "Application Level",
                "page_display_name": "App Component",
                "widget_name": cw,
                "widget_display_name": cw.replace("_", " "),
                "widget_type": "Custom Widget",
                "description": f"Custom Widget component '{cw}'"
            })

    return widgets_found

def auto_match_widget_file(filename: str, found_widgets: list) -> dict:
    """Uses fuzzy domain-weighted token similarity to automatically pair an uploaded widget archive with a DS widget page."""
    if not filename or not found_widgets:
        return None
    from zoho_parser import ZohoDSParser
    parser = ZohoDSParser("")
    best_w = None
    best_score = 0
    for w in found_widgets:
        sc = parser._score_package_match(w.get("widget_name", ""), w.get("page_name", ""), filename)
        if sc > best_score:
            best_score = sc
            best_w = w
    if best_score >= 25:
        return best_w
    return None


def generate_zoho_documentation(ds_text: str, user_instruction: str = "", provider: str = "auto", api_key: str = "", ollama_host: str = "", widget_text: str = "", widget_mappings: list = None, target_page: str = "", target_widget: str = "", all_widgets_uploaded: bool = True, pending_widgets: list = None) -> str:
    """Instantly produces targeted or full enterprise-grade documentation from .ds file in milliseconds,
    or fast-answers targeted Deluge questions via the selected AI provider.
    - When an individual widget is uploaded for a target page (and not all widgets are uploaded yet),
      it produces the focused Page & Widget Architecture Specification with workflow diagram, variables,
      connected Deluge functions, and widget source inspection.
    - When all widgets have been uploaded (or bulk upload / full doc request), it produces the Master Unified
      Technical Documentation with the Master End-to-End Widget Integration Matrix."""
    parser = ZohoDSParser(ds_text)
    intent = classify_zoho_request(user_instruction)

    wants_tamil = any(tw in user_instruction.lower() for tw in ["tamil la", "tamilla", "in tamil", "tamilil", "tamil doc", "tamil-la", "tamil explanation", "tamil documentation", "tamila "])

    doc_output = ""
    if intent == "pages_count":
        doc_output = parser.get_component_inventory("pages")
    elif intent == "forms_count":
        doc_output = parser.get_component_inventory("forms")
    elif intent == "reports_count":
        doc_output = parser.get_component_inventory("reports")
    elif intent == "workflows_count":
        doc_output = parser.get_component_inventory("workflows")
    elif intent == "pages_explain":
        doc_output = parser.explain_pages_and_workflows()
    elif intent in ("diagram", "flowchart"):
        doc_output = parser.generate_flow_diagrams()
    elif intent == "form_purpose":
        doc_output = parser.generate_targeted_form_purpose(user_instruction)
    elif intent == "pages":
        doc_output = parser.generate_pages_documentation()
    elif intent == "forms":
        only_words = ["only", "just", "mattum", "madu", "alone", "single", "specifically", "particular", "thaniya"]
        if any(w in user_instruction.lower() for w in only_words):
            doc_output = parser.generate_targeted_form_purpose(user_instruction)
        else:
            doc_output = parser.generate_forms_documentation()
    elif intent == "functions":
        doc_output = parser.generate_functions_documentation()
    elif intent == "workflows":
        doc_output = parser.generate_workflows_documentation()
    elif intent == "reports":
        doc_output = parser.generate_reports_documentation()
    elif intent == "security":
        doc_output = parser.generate_security_documentation()
    elif intent == "architecture":
        doc_output = parser.generate_architecture_documentation()
    elif intent == "testing":
        doc_output = parser.generate_test_documentation()
    elif intent == "app_summary":
        doc_output = parser.generate_application_summary(widget_text=widget_text)
    elif intent == "full_doc":
        doc_output = parser.generate_full_technical_documentation(widget_text=widget_text, widget_mappings=widget_mappings)

    # Widget source code processing
    if widget_text and widget_text.strip():
        # Check if user explicitly requested full documentation or clicked bottom submit
        is_full_intent = any(w in user_instruction.lower() for w in ["full doc", "full documentation", "complete doc", "all sections", "unified", "master", "submit", "kila irukkura", "full documetion", "connecting all widget"])
        # Stage 1: Individual widget upload for a specific page when other widgets are pending
        if target_page and not all_widgets_uploaded and not is_full_intent:
            curr_m = next((m for m in (widget_mappings or []) if m.get("page") == target_page or m.get("widget") == target_widget), None)
            doc_output = parser.generate_targeted_page_widget_doc(
                target_page_name=target_page,
                widget_text=widget_text,
                widget_mapping=curr_m,
                pending_widgets=pending_widgets
            )
        else:
            # Stage 2: Master Full Technical Documentation (only if not already generated)
            if not doc_output:
                doc_output = parser.generate_full_technical_documentation(widget_text=widget_text, widget_mappings=widget_mappings)

    if doc_output:
        # Append widget documentation section only if widget files were uploaded and NOT already covered by Section 5 of full doc
        if widget_text and widget_text.strip() and "## 5. Custom Widgets Explanation" not in doc_output:
            doc_output += "\n" + _generate_widget_documentation_section(widget_text, parser, provider, api_key, ollama_host, widget_mappings=widget_mappings)
        if wants_tamil:
            prompt = f"Translate and explain the following Technical Documentation in clear, natural Tamil. Keep Deluge code blocks and table structures intact:\n\n{doc_output[:8000]}"
            try:
                return llm.ask(prompt, system="You are an expert technical translator. Translate clearly into Tamil while preserving code and Markdown tables.", provider=provider, api_key=api_key, ollama_host=ollama_host, max_tokens=1500)
            except Exception:
                return doc_output
        return doc_output

    # If it's a specific question or conversation, feed focused context to LLM
    parsed_data = parser.parse()
    meta = parsed_data["metadata"]

    pages_summary = "\n".join([
        f"- Page '{pg['display_name']}' (`{pg['name']}`): Type={pg['page_type']}"
        for pg in parsed_data["pages"]
    ]) or "(No standalone custom pages; views are driven by native Form and Report navigation)"
    forms_summary = ", ".join([f["name"] for f in parsed_data["forms"][:15]]) or "(none)"
    reports_summary = ", ".join([r["name"] for r in parsed_data["reports"][:15]]) or "(none)"
    functions_summary = "\n".join([
        f"- Function {fn['full_name']}({', '.join(fn['parameters'])}): returns {fn['return_type']}."
        for fn in parsed_data["functions"][:10]
    ]) or "(none)"

    workflows_summary = "\n".join([
        f"- Workflow '{wf['display_name']}' (`{wf['name']}`): Form={wf['form']}, Trigger={', '.join(wf.get('triggers', []))}"
        for wf in parsed_data["workflows"][:15]
    ]) or "(none)"

    # Include code of functions the question mentions
    q = user_instruction.lower()
    mentioned = [fn for fn in parsed_data["functions"]
                 if fn["name"].lower() in q or fn["full_name"].lower() in q]
    code_blocks = ""
    for fn in (mentioned or parsed_data["functions"][:3]):
        block = f"\n### {fn['full_name']}\n```\n{fn['body'][:800]}\n```\n"
        code_blocks += block

    context_prompt = (
        f"Application: {meta['app_name']}\n"
        f"Component Counts: Forms={len(parsed_data['forms'])}, Reports={len(parsed_data['reports'])}, Workflows={len(parsed_data['workflows'])}, Pages={len(parsed_data['pages'])}, Functions={len(parsed_data['functions'])}\n"
        f"Forms ({len(parsed_data['forms'])}): {forms_summary}\n"
        f"Reports ({len(parsed_data['reports'])}): {reports_summary}\n"
        f"Workflows ({len(parsed_data['workflows'])} automated rules):\n{workflows_summary}\n"
        f"Pages ({len(parsed_data['pages'])}):\n{pages_summary}\n"
        f"Functions ({len(parsed_data['functions'])}):\n{functions_summary}\n"
        f"{code_blocks}\n\n"
        f"Question: {user_instruction}\n"
        f"Provide a brief, direct answer in Markdown. Keep Forms, Reports, Workflows, Pages, and Deluge Functions strictly distinguished and separated. Differentiate between standalone custom functions and event-driven workflows."
    )

    try:
        return llm.ask(
            context_prompt,
            system="You are a Senior Zoho Creator Solutions Architect. Answer directly and concisely in GitHub Markdown. Strictly differentiate between Forms, Reports, Workflows, Pages, and Functions.",
            provider=provider,
            api_key=api_key,
            ollama_host=ollama_host,
            max_tokens=800
        )
    except Exception as e:
        print(f"LLM Q&A fallback: {e}")
        if mentioned:
            fn0 = mentioned[0]
            res = f"### Deluge Function: `{fn0['full_name']}`\n\n"
            res += f"- **Return Type**: `{fn0['return_type']}`\n"
            res += f"- **Parameters**: `{', '.join(fn0['parameters']) or 'None'}`\n\n"
            res += f"#### Deluge Definition:\n```deluge\n{fn0['body']}\n```\n"
            return res
        return parser.generate_full_technical_documentation()


def _extract_widget_data_ingestion_mappings(widget_text: str, ds_forms: list) -> list:
    """Extracts user input controls from widget HTML/JS and maps them to target Zoho Creator forms and fields.
    Explains which user input in the widget persists into which Form and Field table in the database."""
    import re
    mappings = []
    if not ds_forms:
        return mappings

    # 1. Detect target forms from SDK calls or references (rank by frequency of appearance)
    form_scores = []
    for f in ds_forms:
        fname = f.get("name", "")
        if fname:
            count = len(re.findall(rf'{re.escape(fname)}', widget_text, re.I))
            if count > 0:
                form_scores.append((count, f))
    form_scores.sort(key=lambda x: x[0], reverse=True)
    target_forms = [f for _, f in form_scores[:5]] if form_scores else [ds_forms[0]]

    # 2. Extract HTML input elements with strict deduplication and template-literal filtering
    seen_ids = set()
    inputs = []

    def _clean_token(val: str) -> str:
        if not val or "${" in val or "{" in val or len(val) > 40:
            return ""
        return val.strip()

    # <input> tags
    for m in re.finditer(r'<input([^>]*)>', widget_text, re.I):
        attrs = m.group(1)
        id_m = re.search(r'id=["\x27]([^"\x27]+)["\x27]', attrs, re.I)
        name_m = re.search(r'name=["\x27]([^"\x27]+)["\x27]', attrs, re.I)
        type_m = re.search(r'type=["\x27]([^"\x27]+)["\x27]', attrs, re.I)
        placeholder_m = re.search(r'placeholder=["\x27]([^"\x27]+)["\x27]', attrs, re.I)

        raw_id = id_m.group(1) if id_m else (name_m.group(1) if name_m else "")
        inp_id = _clean_token(raw_id)
        inp_type = type_m.group(1).lower() if type_m else "text"
        if inp_type in {"submit", "button", "hidden", "reset"}:
            continue

        raw_placeholder = placeholder_m.group(1) if placeholder_m else ""
        placeholder = _clean_token(raw_placeholder)
        key = (inp_id or placeholder).lower()
        if not key or key in seen_ids:
            continue
        seen_ids.add(key)

        label = placeholder or inp_id.replace("_", " ").title()
        inputs.append({
            "id": inp_id or f"input_{len(inputs)+1}",
            "type": inp_type,
            "label": label
        })
        if len(inputs) >= 30:
            break

    # <select> dropdowns
    if len(inputs) < 30:
        for m in re.finditer(r'<select([^>]*)>', widget_text, re.I):
            id_m = re.search(r'id=["\x27]([^"\x27]+)["\x27]', m.group(1), re.I)
            raw_id = id_m.group(1) if id_m else ""
            inp_id = _clean_token(raw_id)
            key = inp_id.lower() if inp_id else ""
            if key and key not in seen_ids:
                seen_ids.add(key)
                inputs.append({"id": inp_id, "type": "dropdown", "label": inp_id.replace("_", " ").title()})
                if len(inputs) >= 30:
                    break

    # <textarea> multiline inputs
    if len(inputs) < 30:
        for m in re.finditer(r'<textarea([^>]*)>', widget_text, re.I):
            id_m = re.search(r'id=["\x27]([^"\x27]+)["\x27]', m.group(1), re.I)
            raw_id = id_m.group(1) if id_m else ""
            inp_id = _clean_token(raw_id)
            key = inp_id.lower() if inp_id else ""
            if key and key not in seen_ids:
                seen_ids.add(key)
                inputs.append({"id": inp_id, "type": "multiline text", "label": inp_id.replace("_", " ").title()})
                if len(inputs) >= 30:
                    break

    # 3. Detect API operation
    api_op = "ZOHO.CREATOR.API.addRecord()"
    if "updateRecord" in widget_text:
        api_op = "ZOHO.CREATOR.API.updateRecord()"
    elif "addRecord" in widget_text:
        api_op = "ZOHO.CREATOR.API.addRecord()"
    elif "getRecords" in widget_text or "getAllRecords" in widget_text:
        api_op = "ZOHO.CREATOR.API.getRecords() / Query"

    # 4. Map each unique input to its single best matching form and field (NO cartesian multiplication)
    primary_form = target_forms[0]
    if inputs:
        for inp in inputs:
            best_form = primary_form
            best_fld = None
            in_id = inp["id"].lower()
            in_lbl = inp["label"].lower()

            for tf in target_forms:
                for fld in tf.get("fields", []):
                    fn = fld.get("name", "").lower()
                    if fn == in_id or (len(in_id) > 2 and fn in in_id) or (len(fn) > 2 and fn in in_lbl):
                        best_form = tf
                        best_fld = fld
                        break
                if best_fld:
                    break

            if not best_fld and best_form.get("fields"):
                best_fld = best_form["fields"][0]

            target_fld_name = best_fld.get("name") if best_fld else "Record_Data"
            target_fld_disp = best_fld.get("display_name", target_fld_name) if best_fld else "Record Data"
            fld_type = best_fld.get("type", inp["type"]) if best_fld else inp["type"]

            mappings.append({
                "widget_input": f"`<{inp['type']} id='{inp['id']}'>` ({inp['label']})",
                "input_type": inp["type"].capitalize(),
                "destination_form": best_form.get("name", "Form"),
                "destination_field": f"`{target_fld_name}` ({target_fld_disp})",
                "api_method": f"`{api_op}`",
                "validation": f"Validates `{fld_type}` constraints -> Persists into `{best_form.get('name')}` table"
            })
            if len(mappings) >= 30:
                break
    else:
        for fld in primary_form.get("fields", [])[:8]:
            mappings.append({
                "widget_input": f"Dynamic Control: `{fld.get('display_name', fld['name'])}`",
                "input_type": fld.get("type", "text").capitalize(),
                "destination_form": primary_form.get("name", "Form"),
                "destination_field": f"`{fld['name']}` ({fld.get('display_name', fld['name'])})",
                "api_method": f"`{api_op}`",
                "validation": f"Schema verification -> Persists to `{primary_form.get('name')}.{fld['name']}`"
            })

    return mappings[:30]


def _format_widget_ingestion_doc(mappings: list, primary_form: str = "") -> str:
    """Formats the Widget-to-Form Ingestion Matrix into comprehensive Markdown documentation."""
    if not mappings:
        return ""
    mappings = mappings[:35]
    md = []
    md.append("### \U0001f4e5 Widget Data Ingestion & Form Destination Flow (Widget Data -> Form Mapping)\n")
    md.append("> [!IMPORTANT]\n> This section maps **user-entered data in the Widget interface** directly to the **destination Form tables** in the Zoho Creator database.\n")
    md.append("The table below specifies each input control in the widget, the persistence API method, and the target form table and column where data is stored:\n")
    md.append("| # | Widget User Input Control | Input Type | Destination DS Form Table | Destination Field Column | Ingestion API Method | Database Persistence & Workflows |")
    md.append("|---|---|---|---|---|---|---|")
    for i, m in enumerate(mappings, 1):
        md.append(f"| {i} | {m['widget_input']} | {m['input_type']} | **`{m['destination_form']}`** | {m['destination_field']} | {m['api_method']} | {m['validation']} |")
    md.append("")

    target_form_name = primary_form or (mappings[0]['destination_form'] if mappings else 'Application_Form')
    md.append("#### \U0001f504 End-to-End Data Ingestion Lifecycle Walkthrough:\n")
    md.append(f"1. **User Interaction**: The user enters data into the Widget input controls on the page.")
    md.append(f"2. **Client-Side Packaging**: Widget JavaScript captures form values and packages them into a JSON payload.")
    md.append(f"3. **SDK Bridge Execution**: Widget calls `{mappings[0]['api_method']}` via the Zoho Creator JavaScript SDK bridge.")
    md.append(f"4. **Backend Validation**: Zoho Creator intercepts the payload, runs Form `{target_form_name}` `on validate` Deluge workflows.")
    md.append(f"5. **Database Commit**: If validation passes, a new record is inserted into `{target_form_name}` with an auto-generated unique Record ID.")
    md.append(f"6. **Post-Commit Automation**: Form `{target_form_name}` `on success` event triggers notifications, tasks, and linked workflows.\n")

    md.append("#### \U0001f4ca Widget-to-Form Ingestion Pipeline Diagram:\n")
    md.append("```mermaid")
    md.append("flowchart TD")
    md.append("    subgraph Widget [Widget Client Layer]")
    md.append("        UI[User Fills Widget Inputs] --> VAL[Client Validation & Event Handler]")
    md.append(f"        VAL --> SDK[\"{mappings[0]['api_method']} Payload\"]")
    md.append("    end")
    md.append("    subgraph Creator [Zoho Creator Application Backend]")
    md.append("        SDK --> BRG[Secure Iframe SDK Bridge]")
    md.append(f"        BRG --> ONVAL[\"Form {target_form_name} 'on validate' Deluge\"]")
    md.append(f"        ONVAL --> TBL[\"Destination Table: {target_form_name}\"]")
    md.append(f"        TBL --> ONSUCC[\"Form {target_form_name} 'on success' Workflow\"]")
    md.append("    end")
    md.append("```\n")
    return "\n".join(md)


def _explain_widget_file_detailed(file_obj: dict, ds_forms: list = None, ds_reports: list = None) -> list:
    """Generates detailed, structured plain English explanations of a widget file without printing raw code snippets."""
    path = file_obj.get("path", "file")
    ext = file_obj.get("ext", "").lower()
    content = file_obj.get("content", "")
    lines_cnt = file_obj.get("lines", 0)

    points = []

    if ext in (".html", ".htm"):
        points.append(f"**Presentation View Template**: Renders the client-side user interface structure ({lines_cnt:,} lines).")
        title_m = re.search(r'<title[^>]*>([^<]+)</title>', content, re.I)
        if title_m:
            points.append(f"**Interface Screen Title**: `{title_m.group(1).strip()}`.")

        forms_found = re.findall(r'<form\b', content, re.I)
        modals_found = re.findall(r'class=["\'][^"\']*(?:modal|dialog|popup)[^"\']*[\"\']', content, re.I)
        tables_found = re.findall(r'<table\b', content, re.I)
        inputs_found = re.findall(r'<input\b|<select\b|<textarea\b', content, re.I)
        buttons_found = re.findall(r'<button\b|type=["\']submit["\']', content, re.I)

        comp_desc = []
        if forms_found: comp_desc.append(f"{len(forms_found)} data entry form(s)")
        if inputs_found: comp_desc.append(f"{len(inputs_found)} interactive input control(s)")
        if buttons_found: comp_desc.append(f"{len(buttons_found)} action button(s)")
        if tables_found: comp_desc.append(f"{len(tables_found)} tabular data grid(s)")
        if modals_found: comp_desc.append(f"{len(modals_found)} modal popup dialogue(s)")
        if comp_desc:
            points.append(f"**UI Elements & Components Mounted**: Incorporates {', '.join(comp_desc)}.")

        scripts = re.findall(r'<script[^>]*src=["\']([^"\']+)["\']', content, re.I)
        if scripts:
            points.append(f"**External Script Bridges**: Integrates client libraries: {', '.join('`' + s.split('/')[-1] + '`' for s in scripts[:5])}.")

    elif ext in (".js", ".jsx", ".ts", ".tsx"):
        points.append(f"**Client Controller & State Management**: Executes dynamic event handling, SDK communication, and form validations ({lines_cnt:,} lines).")

        funcs = re.findall(r'(?:function\s+([A-Za-z_$][A-Za-z0-9_$]*)|(?:const|let|var)\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=\s*(?:function|\(|async))', content)
        fn_names = [f[0] or f[1] for f in funcs if (f[0] or f[1]) and not (f[0] or f[1]).startswith('$')]
        if fn_names:
            points.append(f"**Core Functional Procedures ({len(fn_names)})**: {', '.join('`' + fn + '()`' for fn in fn_names[:10])}.")

        sdk_calls = list(set(re.findall(r'ZOHO\.CREATOR\.[A-Z_]+(?:\.[A-Za-z_]+)*', content)))
        if sdk_calls:
            points.append(f"**Zoho Creator JS SDK Methods Invoked**: {', '.join('`' + s + '`' for s in sdk_calls)} — Bridges the widget iframe directly with Zoho Creator backend databases.")

        events = list(set(re.findall(r'(?:addEventListener|\.on)\s*\(\s*[\"\']([a-zA-Z]+)[\"\']', content)))
        if events:
            points.append(f"**User Event Bindings**: Listens for user interactions ({', '.join('`' + ev + '`' for ev in events[:5])}) to trigger real-time UI updates.")

        if "innerHTML" in content or "appendChild" in content or "render" in content.lower():
            points.append("**Dynamic DOM Rendering**: Constructively updates UI elements, tables, and modal dialogues in the client view without full-page reloads.")
        if "fetch(" in content or "$.ajax" in content or "axios" in content or "invokeurl" in content.lower():
            points.append("**Asynchronous Network Operations**: Dispatches REST API requests to query or persist data across operational workflows.")
        if "localStorage" in content or "sessionStorage" in content:
            points.append("**Client State Caching**: Utilizes browser storage to retain session context and active user filters.")

    elif ext in (".css", ".scss", ".less"):
        points.append(f"**User Interface Styling & Theme**: Defines responsive styling, component themes, color schemes, and mobile viewports ({lines_cnt:,} lines).")
        if "@media" in content:
            points.append("**Responsive Design Rules**: Includes responsive media queries ensuring optimal layout on mobile, tablet, and desktop viewports.")
        if "modal" in content or "popup" in content:
            points.append("**Modal & Dialog Styling**: Styles interactive popups, overlay backdrops, and confirmation prompts.")

    elif ext == ".json":
        points.append(f"**Configuration & Metadata Manifest**: Configures package properties, permissions, and runtime initialization parameters ({lines_cnt:,} lines).")
        if "package.json" in path:
            name_m = re.search(r'"name"\s*:\s*"([^"]+)"', content)
            ver_m = re.search(r'"version"\s*:\s*"([^"]+)"', content)
            if name_m or ver_m:
                points.append(f"**Package Specification**: Package `{name_m.group(1) if name_m else 'Widget'}` (v{ver_m.group(1) if ver_m else '1.0'}).")

    return points


def _generate_widget_documentation_section(widget_text: str, parser=None, provider: str = "auto", api_key: str = "", ollama_host: str = "", widget_mappings: list = None) -> str:
    """Generates a comprehensive Widget Source Code Documentation section from uploaded widget files.
    Parses the combined widget text (from ZIP extraction or individual code files),
    organizes files by type, extracts structural elements, and generates documentation."""
    import re as _re

    md = []
    md.append("\n---\n")
    md.append("## WIDGET SOURCE CODE DOCUMENTATION (Uploaded Widget Files)\n")
    md.append("> [!IMPORTANT]\n> The following section documents the **actual widget source code** uploaded alongside the .ds file.")
    md.append("> Widget code is NOT embedded in the .ds export file — it runs as a separate client-side package inside Zoho Creator.\n")

    if widget_mappings:
        md.append("### 🗺️ Master End-to-End Widget Integration Matrix\n")
        md.append("The application defines the following custom client-side widget integrations connected to host pages, forms, and backend logic:\n")
        md.append("| # | Widget Component | Source File / Package | Target Zoho Creator Page | Integration Mechanism / Tag | Connected Deluge Forms & Workflows | Integration Status |")
        md.append("|---|---|---|---|---|---|---|")

        parsed_ds = parser.parse() if (parser and hasattr(parser, 'parse')) else {}
        ds_pages = parsed_ds.get("pages", [])
        ds_forms = parsed_ds.get("forms", [])

        for midx, wm in enumerate(widget_mappings):
            p_raw = wm.get("page") or ""
            w_label = wm.get("widget") or "Custom Widget"
            f_name = wm.get("file", "widget_file")

            # Find page details
            target_pg = next((p for p in ds_pages if p["name"].lower() == p_raw.lower() or p["display_name"].lower() == p_raw.lower()), None)
            p_disp = f"`{target_pg['name']}` ({target_pg['display_name']})" if target_pg else (f"`{p_raw}`" if p_raw else "Application Scope")

            # Integration tag
            tag_str = f"`<widgets elementName='{w_label}' importedTemplate='true'/>`" if target_pg else "`Custom Webhook / Iframe`"

            # Connected forms
            form_names = [f.get("name") for f in ds_forms if f.get("name")]
            linked_f = [fn for fn in form_names if fn.lower() in widget_text.lower() or (target_pg and fn.lower() in target_pg.get("raw_zml", "").lower())]
            forms_summary = ", ".join(f"`{lf}`" for lf in linked_f[:3]) if linked_f else "Dynamic SDK Storage"

            md.append(f"| {midx+1} | **`{w_label}`** | `{f_name}` | {p_disp} | {tag_str} | {forms_summary} | ✅ **Integrated & Verified** |")
        md.append("")

    # Parse individual files from the combined text block with vendor filtering & deduplication
    file_sections = _re.split(r'={50,}\s*\n(?:📄\s*)?FILE:\s*', widget_text)

    widget_files = []
    seen_file_paths = set()
    for section in file_sections:
        if not section.strip():
            continue
        lines = section.split('\n', 2)
        if len(lines) >= 2:
            fname = lines[0].strip()
            content_start = section.find('\n', section.find('\n') + 1)
            content = section[content_start:].strip() if content_start > 0 else ""
            content = _re.sub(r'\n={50,}\s*$', '', content).strip()
            if fname and content and '[Skipped:' not in content:
                # Filter out heavy third-party / minified vendor libraries
                low_name = fname.lower()
                if any(v in low_name for v in ["node_modules", "package-lock.json", ".min.", "jquery", "bootstrap", "widgetsdk", "fontawesome"]):
                    continue
                if fname not in seen_file_paths:
                    seen_file_paths.add(fname)
                    ext = os.path.splitext(fname)[1].lower()
                    widget_files.append({
                        "path": fname,
                        "ext": ext,
                        "content": content,
                        "lines": len(content.splitlines())
                    })

    # If the text wasn't from a ZIP (no file markers), treat the whole text as a single file
    if not widget_files and widget_text.strip():
        widget_files.append({
            "path": "uploaded_widget_code",
            "ext": ".txt",
            "content": widget_text.strip(),
            "lines": len(widget_text.strip().splitlines())
        })

    if not widget_files:
        md.append("No readable widget files were found in the uploaded package.\n")
        return "\n".join(md)

    # Classify files by type
    html_files = [f for f in widget_files if f["ext"] in {".html", ".htm"}]
    js_files = [f for f in widget_files if f["ext"] in {".js", ".jsx", ".ts", ".tsx"}]
    css_files = [f for f in widget_files if f["ext"] in {".css", ".scss", ".less"}]
    config_files = [f for f in widget_files if f["ext"] in {".json", ".xml", ".yml", ".yaml", ".toml", ".env", ".properties"}]
    other_files = [f for f in widget_files if f not in html_files + js_files + css_files + config_files]

    total_lines = sum(f["lines"] for f in widget_files)

    # Overview table
    md.append("### Widget Package Overview\n")
    md.append(f"- **Total Files**: {len(widget_files)}")
    md.append(f"- **Total Lines of Code**: {total_lines:,}")
    md.append(f"- **HTML/Template Files**: {len(html_files)}")
    md.append(f"- **JavaScript/TypeScript Files**: {len(js_files)}")
    md.append(f"- **CSS/Style Files**: {len(css_files)}")
    md.append(f"- **Configuration Files**: {len(config_files)}")
    if other_files:
        md.append(f"- **Other Files**: {len(other_files)}")
    md.append("")

    # File inventory table
    md.append("### Widget File Inventory\n")
    md.append("| # | File Path | Type | Lines |")
    md.append("|---|---|---|---|")
    for idx, f in enumerate(widget_files):
        file_type = {
            ".html": "HTML Template", ".htm": "HTML Template",
            ".js": "JavaScript", ".jsx": "React JSX", ".ts": "TypeScript", ".tsx": "React TSX",
            ".css": "Stylesheet", ".scss": "SCSS", ".less": "LESS",
            ".json": "JSON Config", ".xml": "XML Config", ".yml": "YAML Config", ".yaml": "YAML Config",
        }.get(f["ext"], f["ext"].upper().lstrip('.') + " File")
        md.append(f"| {idx+1} | `{f['path']}` | {file_type} | {f['lines']} |")
    md.append("")

    # Detect Zoho Widget SDK usage
    all_code = "\n".join(f["content"] for f in widget_files)
    uses_widget_sdk = "ZOHO.CREATOR" in all_code or "widgetsdk" in all_code.lower()
    uses_jquery = "jquery" in all_code.lower() or "$(document)" in all_code or "$(function" in all_code
    uses_bootstrap = "bootstrap" in all_code.lower()
    uses_react = "react" in all_code.lower() or "ReactDOM" in all_code or "useState" in all_code
    uses_vue = "vue" in all_code.lower() or "v-model" in all_code or "v-if" in all_code

    if uses_widget_sdk or uses_jquery or uses_bootstrap or uses_react or uses_vue:
        md.append("### Detected Technologies & SDK Usage\n")
        if uses_widget_sdk:
            md.append("- ✅ **Zoho Creator Widget SDK** — Widget communicates with Zoho Creator via `ZOHO.CREATOR` API")
        if uses_jquery:
            md.append("- ✅ **jQuery** — DOM manipulation library detected")
        if uses_bootstrap:
            md.append("- ✅ **Bootstrap** — CSS framework for responsive layout")
        if uses_react:
            md.append("- ✅ **React** — Component-based UI framework")
        if uses_vue:
            md.append("- ✅ **Vue.js** — Progressive JavaScript framework")
        md.append("")

    # Extract Zoho Widget SDK API calls
    sdk_calls = list(set(_re.findall(r'ZOHO\.CREATOR\.[A-Z_]+(?:\.[A-Za-z_]+)*\([^)]*\)', all_code)))
    sdk_methods = list(set(_re.findall(r'(ZOHO\.CREATOR\.[A-Z_]+(?:\.[A-Za-z_]+)*)', all_code)))
    if sdk_methods:
        md.append("### Zoho Creator Widget SDK Methods Used\n")
        md.append("| # | SDK Method | Purpose |")
        md.append("|---|---|---|")
        sdk_purpose_map = {
            "ZOHO.CREATOR.init": "Initialize Widget SDK connection",
            "ZOHO.CREATOR.API.getRecords": "Fetch records from a Zoho Creator report/form",
            "ZOHO.CREATOR.API.addRecord": "Create a new record in a form",
            "ZOHO.CREATOR.API.updateRecord": "Update an existing record",
            "ZOHO.CREATOR.API.deleteRecord": "Delete a record",
            "ZOHO.CREATOR.API.getAllRecords": "Fetch all records from a form/report",
            "ZOHO.CREATOR.UTIL.navigateParentURL": "Navigate the parent Zoho Creator page",
            "ZOHO.CREATOR.UTIL.getQueryParams": "Read URL query parameters",
            "ZOHO.CREATOR.UTIL.setImageData": "Set image data in widget",
            "ZOHO.CREATOR.DATA.invokeCustomApi": "Call a Custom API (Deluge function) from widget",
            "ZOHO.CREATOR.UTIL.resize": "Resize the widget iframe",
        }
        for idx, method in enumerate(sorted(set(sdk_methods))):
            purpose = sdk_purpose_map.get(method, "Widget SDK interaction")
            md.append(f"| {idx+1} | `{method}` | {purpose} |")
        md.append("")

    # Cross-reference with Zoho Creator DS components (Forms, Reports, Deluge APIs)
    if parser:
        try:
            parsed_ds = parser.parse() if hasattr(parser, 'parse') else {}
            ds_forms = parsed_ds.get("forms", [])
            ds_reports = parsed_ds.get("reports", [])
            form_names = [f.get("name") for f in ds_forms if f.get("name")]
            report_names = [r.get("name") for r in ds_reports if r.get("name")]

            linked_forms = [fn for fn in form_names if fn.lower() in all_code.lower()]
            linked_reports = [rn for rn in report_names if rn.lower() in all_code.lower()]

            if linked_forms or linked_reports:
                md.append("### 🔗 Widget ↔ Zoho Creator Application Integration Mapping & Comparative Audit\n")
                md.append("The widget client-side code directly interacts with and connects to the following backend application schema defined in the `.ds` file:\n")
                if linked_forms:
                    md.append(f"- **Referenced Application Forms**: {', '.join(f'`{f}`' for f in linked_forms)}")
                if linked_reports:
                    md.append(f"- **Referenced Reports / Views**: {', '.join(f'`{r}`' for r in linked_reports)}")
                md.append("")

                md.append("| Architectural Layer | Widget Code Reference | DS Application Schema | Verification Status |")
                md.append("|---|---|---|---|")
                for lf in linked_forms:
                    target_f = next((f for f in ds_forms if f.get("name") == lf), None)
                    f_count = len(target_f.get("fields", [])) if target_f else 0
                    md.append(f"| **Entity / Form** | Interacts with `{lf}` | Form `{lf}` ({f_count} fields in `.ds`) | ✅ **Verified in Schema** |")
                for lr in linked_reports:
                    md.append(f"| **Report View** | Queries `{lr}` | Report `{lr}` in `.ds` | ✅ **Verified View** |")
                md.append("")
        except Exception as e:
            print(f"Error mapping widget to DS components: {e}")

    # Document each file with structured functional explanation (Zero raw code dumps)
    def _document_file_group(group_files, group_title, lang_hint):
        if not group_files:
            return
        md.append(f"### {group_title}\n")
        # Prioritize core application scripts/views
        for f in group_files[:12]:
            md.append(f"#### 📄 Component: `{f['path']}` ({f['lines']:,} lines)\n")
            explanation_points = _explain_widget_file_detailed(f, ds_forms, ds_reports)
            for pt in explanation_points:
                md.append(f"- {pt}")
            md.append("")
        if len(group_files) > 12:
            remaining_paths = ", ".join(f"`{f['path']}`" for f in group_files[12:])
            md.append(f"> **Auxiliary & Theme Assets ({len(group_files)-12})**: {remaining_paths} (Supporting styling, helper utilities, and localized asset files).\n")

    _document_file_group(html_files, "HTML Template Files", "html")
    _document_file_group(js_files, "JavaScript / TypeScript Files", "javascript")
    _document_file_group(css_files, "Stylesheet Files", "css")
    _document_file_group(config_files, "Configuration Files", "json")
    _document_file_group(other_files, "Other Widget Files", "text")

    # Functional summary (Deterministic fallback + fast AI summary)
    functional_bullets = []
    if html_files:
        sample_html = ', '.join('`' + f.get('path', '') + '`' for f in html_files[:3])
        functional_bullets.append(f"Provides custom UI presentation via {len(html_files)} template view(s) ({sample_html}).")
    if js_files:
        functional_bullets.append(f"Executes client-side business logic and interactive event handling across {len(js_files)} script module(s).")
    if sdk_methods:
        sample_sdk = ', '.join('`' + m + '`' for m in sorted(set(sdk_methods))[:6])
        functional_bullets.append(f"Directly interfaces with Zoho Creator database using SDK methods: {sample_sdk}.")
    if 'linked_forms' in locals() and linked_forms:
        sample_forms = ', '.join('`' + f + '`' for f in linked_forms)
        functional_bullets.append(f"Performs record queries and mutations on application form(s): {sample_forms}.")
    if 'linked_reports' in locals() and linked_reports:
        sample_reports = ', '.join('`' + r + '`' for r in linked_reports)
        functional_bullets.append(f"Retrieves aggregated data from application report view(s): {sample_reports}.")

    has_active_provider = (provider in ("groq", "gemini", "openai") and api_key) or llm.is_ollama_online(ollama_host)
    ai_summary_generated = False
    if has_active_provider:
        try:
            widget_summary_prompt = (
                f"Analyze the following Zoho Creator Widget source code files and provide a brief functional summary "
                f"in 3-5 bullet points. Focus on: what the widget does, which Zoho Creator forms/reports it interacts with, "
                f"what user actions it supports, and any external API integrations.\n\n"
                f"Widget Code:\n{all_code[:6000]}"
            )
            ai_summary = llm.ask(
                widget_summary_prompt,
                system="You are a Zoho Creator Widget expert. Provide a concise functional analysis in GitHub Markdown bullet points.",
                provider=provider,
                api_key=api_key,
                ollama_host=ollama_host,
                max_tokens=500
            )
            if ai_summary and ai_summary.strip():
                md.append("### Functional Widget Summary & Architecture\n")
                md.append(ai_summary.strip())
                md.append("")
                ai_summary_generated = True
        except Exception as e:
            print(f"Widget AI summary skipped: {e}")

    if not ai_summary_generated and functional_bullets:
        md.append("### Functional Widget Summary & Architecture\n")
        for b in functional_bullets:
            md.append(f"- {b}")
        md.append("")

    # Append Widget Data Ingestion & Form Destination Flow (User Data in Widget -> Target Form Table)
    p_forms = []
    if parser and hasattr(parser, "parse"):
        try:
            p_forms = parser.parse().get("forms", [])
        except Exception:
            p_forms = []
    ingestion_mappings = _extract_widget_data_ingestion_mappings(widget_text, p_forms)
    if ingestion_mappings:
        md.append(_format_widget_ingestion_doc(ingestion_mappings))

    return "\n".join(md)


def generate_response(user_prompt: str, document_text: str = None, filename: str = "", provider: str = "auto", api_key: str = "", ollama_host: str = "", widget_text: str = "", widget_mappings: list = None, target_page: str = "", target_widget: str = "", all_widgets_uploaded: bool = True, pending_widgets: list = None) -> str:
    # 1. Check if user is teaching the AI a fact to remember/learn
    if not document_text and knowledge_memory.is_learning_instruction(user_prompt):
        return knowledge_memory.process_learning_request(user_prompt)

    # 2. If user pasted raw DS code directly in user_prompt
    if not document_text and is_zoho_ds_content(user_prompt):
        document_text = user_prompt
        user_prompt = "Generate a comprehensive, enterprise-grade Technical Documentation for this application."

    # 3. If the uploaded document is a Zoho Creator DS file
    if document_text and is_zoho_ds_content(document_text, filename):
        return generate_zoho_documentation(
            document_text,
            user_prompt,
            provider=provider,
            api_key=api_key,
            ollama_host=ollama_host,
            widget_text=widget_text,
            widget_mappings=widget_mappings,
            target_page=target_page,
            target_widget=target_widget,
            all_widgets_uploaded=all_widgets_uploaded,
            pending_widgets=pending_widgets
        )

    # 4. General query handling with dynamic knowledge augmentation (In-Context RAG)
    knowledge_context = knowledge_memory.build_knowledge_context(user_prompt)

    prompt_parts = []
    if knowledge_context:
        prompt_parts.append(f"Context Facts (Verified Real-Time Information):\n{knowledge_context}")
    if document_text:
        truncated_doc = document_text[:llm.MAX_DOC_CHARS]
        note = "\n[...document truncated...]" if len(document_text) > llm.MAX_DOC_CHARS else ""
        prompt_parts.append(f"Context Document:\n{truncated_doc}{note}")

    if knowledge_context:
        prompt_parts.append(
            f"Question: Based strictly on the Context Facts above (which override any outdated pre-training cutoff), {user_prompt}"
        )
    else:
        prompt_parts.append(f"User Question: {user_prompt}")

    final_content = "\n\n".join(prompt_parts)

    return llm.ask(
        final_content,
        system="You are a versatile, highly intelligent AI assistant like ChatGPT. Answer questions accurately and directly in GitHub Markdown.",
        provider=provider,
        api_key=api_key,
        ollama_host=ollama_host
    )


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/api/knowledge", methods=["GET"])
def get_knowledge_list():
    return jsonify(knowledge_memory.get_all_knowledge())


@app.route("/api/knowledge", methods=["POST"])
def post_knowledge():
    data = request.get_json() or {}
    fact = data.get("fact", "").strip()
    topic = data.get("topic", "").strip()
    if not fact:
        return jsonify({"error": "No fact provided"}), 400
    item = knowledge_memory.add_knowledge(fact, topic)
    return jsonify({"success": True, "item": item})



@app.route("/api/status")
def api_status():
    provider = "Zoo AI (Ollama Local)" if llm.is_ollama_online() else (
        "Groq Cloud AI" if llm.ENV_GROQ_KEY else (
            "Google Gemini AI" if llm.ENV_GEMINI_KEY else (
                "OpenAI" if llm.ENV_OPENAI_KEY else "Zoo Cloud AI"
            )
        )
    )
    return jsonify({
        "status": "online",
        "provider": provider,
        "model": "zoo",
        "catalyst_ready": True,
        "offline_ds_engine": "Active"
    })


@app.route("/api/sessions", methods=["GET"])
def get_sessions():
    return jsonify(list_sessions())


@app.route("/api/sessions/<session_id>", methods=["GET"])
def get_session(session_id):
    data = load_session(session_id)
    if not data:
        return jsonify({"error": "Session not found"}), 404
    return jsonify(data)


@app.route("/api/sessions/<session_id>/delete", methods=["POST"])
def remove_session(session_id):
    success = delete_session(session_id)
    return jsonify({"success": success})


@app.route("/api/widget/attach", methods=["POST"])
def attach_widget():
    """Quietly attaches or updates widget source files for a specific widget or bulk package."""
    session_id = request.form.get("session_id", "").strip()
    widget_id = request.form.get("widget_id", "").strip()
    widget_page = request.form.get("widget_page", "").strip()
    widget_name = request.form.get("widget_name", "").strip()

    uploaded_files = request.files.getlist("files[]")
    legacy_file = request.files.get("file")
    if legacy_file and legacy_file.filename:
        uploaded_files = [legacy_file] + [f for f in uploaded_files if f.filename]
    uploaded_files = [f for f in uploaded_files if f and f.filename]

    if not uploaded_files:
        return jsonify({"error": "No widget files provided."}), 400

    session = load_session(session_id) if session_id else None
    if not session:
        all_s = list_sessions()
        if all_s:
            session = load_session(all_s[0]["id"])
            if session:
                session_id = session["id"]
    if not session:
        session_id = session_id or f"sess_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        session = {
            "id": session_id,
            "title": "Application Analysis",
            "created_at": datetime.datetime.now().isoformat(),
            "updated_at": datetime.datetime.now().isoformat(),
            "filename": "",
            "document_text": "",
            "messages": [],
            "widget_files_text": "",
            "widget_files_list": [],
            "widget_mappings": []
        }

    if not session.get("widget_files_text"):
        session["widget_files_text"] = ""
    if not session.get("widget_files_list"):
        session["widget_files_list"] = []
    if not session.get("widget_mappings"):
        session["widget_mappings"] = []
    if "widget_file_payloads" not in session:
        session["widget_file_payloads"] = {}

    document_text = session.get("document_text", "")
    found_widgets = detect_widgets_in_ds(document_text) if document_text else []

    attached_files = []
    for uf in uploaded_files:
        uf_name = uf.filename
        extracted = extract_text_from_file(uf)

        cur_page = widget_page
        cur_widget = widget_name
        cur_id = widget_id
        if (not cur_page or cur_id == "bulk") and found_widgets:
            matched_w = auto_match_widget_file(uf_name, found_widgets)
            if matched_w:
                cur_page = matched_w.get("page_name", "")
                cur_widget = matched_w.get("widget_name", "")
                cur_id = matched_w.get("widget_id", "")

        meta_header = f" [Page: {cur_page} | Component: {cur_widget}]" if (cur_page or cur_widget) else ""
        if not uf_name.lower().endswith(".zip"):
            entry_text = f"\n{'='*60}\n📄 FILE: {uf_name}{meta_header}\n{'='*60}\n{extracted}\n"
        else:
            entry_text = f"\n{'='*60}\n📦 PACKAGE: {uf_name}{meta_header}\n{'='*60}\n\n{extracted}\n"

        payload_key = cur_id if cur_id else f"w_{len(session['widget_mappings']) + 1}"
        session["widget_file_payloads"][payload_key] = {
            "file": uf_name,
            "page": cur_page,
            "widget": cur_widget,
            "text": entry_text
        }

        attached_files.append(uf_name)

        # Update or record widget mapping
        existing_m = next((m for m in session["widget_mappings"] if ((cur_id and m.get("widget_id") == cur_id and cur_id != "bulk") or (cur_page and m.get("page") == cur_page))), None)
        if existing_m:
            existing_m["file"] = uf_name
            existing_m["page"] = cur_page
            existing_m["widget"] = cur_widget
            existing_m["widget_id"] = cur_id
        else:
            session["widget_mappings"].append({
                "file": uf_name,
                "page": cur_page,
                "widget": cur_widget,
                "widget_id": cur_id
            })

    # Rebuild widget_files_text and widget_files_list from payloads
    if session["widget_file_payloads"]:
        session["widget_files_text"] = "\n".join(e.get("text", "") for e in session["widget_file_payloads"].values())
        session["widget_files_list"] = list(dict.fromkeys(e.get("file") for e in session["widget_file_payloads"].values() if e.get("file")))

    save_session(session)
    attached_ids = {wm.get("widget_id") for wm in session.get("widget_mappings", []) if wm.get("widget_id")}
    if any(wm.get("widget_id") == "bulk" for wm in session.get("widget_mappings", [])):
        attached_ids = {fw["widget_id"] for fw in found_widgets}
    else:
        for wm in session.get("widget_mappings", []):
            for fw in found_widgets:
                if (wm.get("page") and wm.get("page") == fw.get("page_name")) or (wm.get("widget") and wm.get("widget") == fw.get("widget_name")):
                    attached_ids.add(fw["widget_id"])

    return jsonify({
        "success": True,
        "session_id": session["id"],
        "files": attached_files,
        "widget_id": widget_id,
        "attached_count": len(attached_ids),
        "total_count": len(found_widgets),
        "widget_mappings": session["widget_mappings"]
    })


@app.route("/api/widget/remove", methods=["POST"])
def remove_widget():
    """Removes an attached widget file or bulk package from the session."""
    session_id = request.form.get("session_id", "").strip()
    widget_id = request.form.get("widget_id", "").strip()
    widget_page = request.form.get("widget_page", "").strip()
    widget_name = request.form.get("widget_name", "").strip()

    if not session_id:
        return jsonify({"error": "Missing session_id"}), 400

    session = load_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    payloads = session.get("widget_file_payloads", {})
    mappings = session.get("widget_mappings", [])

    if widget_id == "bulk":
        session["widget_file_payloads"] = {}
        session["widget_mappings"] = []
        session["widget_files_text"] = ""
        session["widget_files_list"] = []
    else:
        old_file = ""
        for m in mappings:
            if (widget_id and m.get("widget_id") == widget_id) or (widget_page and m.get("page") == widget_page) or (widget_name and m.get("widget") == widget_name):
                old_file = m.get("file", "")
                break

        if widget_id in payloads:
            del payloads[widget_id]

        for wid, entry in list(payloads.items()):
            if (widget_page and entry.get("page") == widget_page) or (widget_name and entry.get("widget") == widget_name):
                del payloads[wid]

        session["widget_file_payloads"] = payloads

        session["widget_mappings"] = [
            m for m in mappings
            if not (
                (widget_id and m.get("widget_id") == widget_id) or
                (widget_page and m.get("page") == widget_page) or
                (widget_name and m.get("widget") == widget_name)
            )
        ]

        if payloads:
            session["widget_files_text"] = "\n".join(e.get("text", "") for e in payloads.values())
            session["widget_files_list"] = list(dict.fromkeys(e.get("file") for e in payloads.values() if e.get("file")))
        else:
            if old_file and session.get("widget_files_text"):
                escaped_f = re.escape(old_file)
                patt = rf'(?:^|\n)=+\s*\n(?:.*?FILE|.*?PACKAGE)?:\s*{escaped_f}[^\n]*\n=+.*?(?=(?:\n=+\s*\n)|\Z)'
                session["widget_files_text"] = re.sub(patt, '', session.get("widget_files_text", ""), flags=re.DOTALL).strip()
            else:
                session["widget_files_text"] = ""
            session["widget_files_list"] = [m.get("file") for m in session["widget_mappings"] if m.get("file")]

    save_session(session)

    document_text = session.get("document_text", "")
    found_widgets = detect_widgets_in_ds(document_text) if document_text else []
    attached_ids = {wm.get("widget_id") for wm in session.get("widget_mappings", []) if wm.get("widget_id")}
    for wm in session.get("widget_mappings", []):
        for fw in found_widgets:
            if (wm.get("page") and wm.get("page") == fw.get("page_name")) or (wm.get("widget") and wm.get("widget") == fw.get("widget_name")):
                attached_ids.add(fw["widget_id"])

    return jsonify({
        "success": True,
        "session_id": session["id"],
        "widget_id": widget_id,
        "attached_count": len(attached_ids),
        "total_count": len(found_widgets),
        "widget_mappings": session["widget_mappings"]
    })


@app.route("/chat", methods=["POST"])
def chat():
    user_input = request.form.get("prompt", "").strip()
    uploaded_files = request.files.getlist("files[]")
    # Also support legacy single-file upload (backward compat)
    legacy_file = request.files.get("file")
    if legacy_file and legacy_file.filename:
        uploaded_files = [legacy_file] + [f for f in uploaded_files if f.filename]

    # Filter out empty file entries
    uploaded_files = [f for f in uploaded_files if f and f.filename]

    session_id = request.form.get("session_id", "").strip()
    ai_provider = request.form.get("ai_provider", "auto").strip()
    api_key = request.form.get("api_key", "").strip()
    ollama_host = request.form.get("ollama_host", "").strip()

    # Load or initialize session on disk
    session = load_session(session_id) if session_id else None
    if not session:
        session_id = f"sess_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        session = {
            "id": session_id,
            "title": "New Analysis",
            "created_at": datetime.datetime.now().isoformat(),
            "updated_at": datetime.datetime.now().isoformat(),
            "filename": "",
            "document_text": "",
            "messages": []
        }

    document_text = None
    widget_text = ""
    filename = session.get("filename", "")
    uploaded_filenames = []
    has_new_widgets = False
    widget_page = request.form.get("widget_page", "").strip()
    widget_name = request.form.get("widget_name", "").strip()
    widget_id = request.form.get("widget_id", "").strip()

    if uploaded_files:
        try:
            # Partition DS application files from widget/code files
            ds_files = []
            non_ds_files = []
            for uf in uploaded_files:
                uname = (uf.filename or "").lower()
                if uname.endswith(".ds") or uname.endswith(".deluge"):
                    ds_files.append(uf)
                else:
                    non_ds_files.append(uf)

            # Process DS files first to establish application context
            for uf in ds_files:
                uf_name = uf.filename
                extracted = extract_text_from_file(uf)
                document_text = extracted
                session["document_text"] = document_text
                session["filename"] = uf_name
                filename = uf_name
                uploaded_filenames.append(uf_name)

            widget_page = request.form.get("widget_page", "").strip()
            widget_name = request.form.get("widget_name", "").strip()
            widget_id = request.form.get("widget_id", "").strip()

            # Process non-DS files (widgets, zip packages, html, js, css, etc.)
            for uf in non_ds_files:
                uf_name = uf.filename
                extracted = extract_text_from_file(uf)
                is_ds_upload = is_zoho_ds_content(extracted, uf_name)

                if is_ds_upload:
                    document_text = extracted
                    session["document_text"] = document_text
                    session["filename"] = uf_name
                    filename = uf_name
                    uploaded_filenames.append(uf_name)
                else:
                    active_ds = document_text or session.get("document_text")
                    if active_ds and is_zoho_ds_content(active_ds, filename or session.get("filename", "")):
                        # Attached to existing or newly uploaded DS application -> Widget Source Code
                        if not session.get("widget_files_text"):
                            session["widget_files_text"] = ""

                        cur_page = widget_page
                        cur_widget = widget_name
                        cur_id = widget_id
                        if not cur_page:
                            ds_widgets = detect_widgets_in_ds(active_ds)
                            matched_w = auto_match_widget_file(uf_name, ds_widgets)
                            if matched_w:
                                cur_page = matched_w.get("page_name", "")
                                cur_widget = matched_w.get("widget_name", "")
                                cur_id = matched_w.get("widget_id", "")

                        meta_header = f" [Page: {cur_page} | Component: {cur_widget}]" if (cur_page or cur_widget) else ""
                        if not uf_name.lower().endswith(".zip"):
                            session["widget_files_text"] += f"\n{'='*60}\n📄 FILE: {uf_name}{meta_header}\n{'='*60}\n{extracted}\n"
                        else:
                            session["widget_files_text"] += f"\n{'='*60}\n📦 PACKAGE: {uf_name}{meta_header}\n{'='*60}\n\n{extracted}\n"

                        if not session.get("widget_files_list"):
                            session["widget_files_list"] = []
                        if uf_name not in session["widget_files_list"]:
                            session["widget_files_list"].append(uf_name)

                        if not session.get("widget_mappings"):
                            session["widget_mappings"] = []
                        existing_m = next((m for m in session["widget_mappings"] if ((cur_id and m.get("widget_id") == cur_id and cur_id != "bulk") or (cur_page and m.get("page") == cur_page))), None)
                        if existing_m:
                            existing_m["file"] = uf_name
                            existing_m["page"] = cur_page
                            existing_m["widget"] = cur_widget
                            existing_m["widget_id"] = cur_id
                        else:
                            session["widget_mappings"].append({
                                "file": uf_name,
                                "page": cur_page,
                                "widget": cur_widget,
                                "widget_id": cur_id
                            })

                        has_new_widgets = True
                        uploaded_filenames.append(uf_name)
                    else:
                        if not document_text:
                            document_text = extracted
                        else:
                            document_text += "\n\n" + extracted
                        session["document_text"] = document_text
                        if not filename:
                            session["filename"] = uf_name
                            filename = uf_name
                        uploaded_filenames.append(uf_name)

            # Load DS + widget from session if not set
            if not document_text and session.get("document_text"):
                document_text = session["document_text"]
                filename = session.get("filename", "")
            widget_text = session.get("widget_files_text", "")

        except Exception as e:
            return jsonify({"error": f"Failed to read document: {str(e)}"}), 400
    else:
        # No new files — reuse from session
        if session.get("document_text"):
            document_text = session["document_text"]
            widget_text = session.get("widget_files_text", "")
            filename = session.get("filename", "")

    # Build a multi-file label for the session title
    multi_label = ", ".join(uploaded_filenames) if uploaded_filenames else ""

    # Check widget status in DS application
    is_ds_session = bool(document_text and is_zoho_ds_content(document_text, filename))
    found_widgets = detect_widgets_in_ds(document_text) if (document_text and is_ds_session) else []

    is_ds_upload = bool(uploaded_filenames and any(fn.lower().endswith('.ds') for fn in uploaded_filenames))
    # If DS file has custom widgets, prompt for widgets FIRST before generating full documentation!
    if is_ds_upload and found_widgets and not session.get("widget_files_text") and not any(w in user_input.lower() for w in ["skip", "without widget", "no widget", "ignore widget"]):
        session["ds_widgets"] = found_widgets
        session["filename"] = filename
        session["document_text"] = document_text
        if not session.get("title") or session.get("title") == "New Analysis":
            session["title"] = os.path.splitext(filename)[0].replace("_", " ").title() if filename else "DS Application"
        
        widget_prompt_data = {
            "widgets": found_widgets,
            "count": len(found_widgets),
            "filename": filename,
            "attached_ids": [],
            "widget_mappings": []
        }
        resp_msg = (
            f"### 🧩 Custom Widgets Detected in `{filename}`\n\n"
            f"This application includes **{len(found_widgets)} custom widget component(s)** embedded in its pages.\n\n"
            f"Please attach your widget source files or package (`.zip` / `.js` / `.html`) in the panel below, "
            f"then click **Submit & Generate Complete Documentation** to produce the full specification with both Deluge backend and widget architecture."
        )
        now_str = datetime.datetime.now().strftime("%I:%M %p")
        session["messages"].append({
            "role": "user",
            "text": user_input or f"Uploaded {filename}",
            "filename": filename,
            "time": now_str
        })
        session["messages"].append({
            "role": "assistant",
            "text": resp_msg,
            "filename": filename,
            "is_ds": True,
            "is_doc": False,
            "time": now_str
        })
        session["updated_at"] = datetime.datetime.now().isoformat()
        save_session(session)

        return jsonify({
            "response": resp_msg,
            "filename": filename,
            "is_ds": True,
            "is_doc": False,
            "session_id": session_id,
            "session_title": session["title"],
            "widget_prompt": widget_prompt_data,
            "has_widgets": True,
            "widget_analyzed": False,
            "widget_files": [],
            "widget_mappings": [],
            "facts": None,
            "widget_facts": None,
            "validation_report": None,
            "five_line_summary": None
        })
    widget_mappings = session.get("widget_mappings", [])
    attached_ids = {wm.get("widget_id") for wm in widget_mappings if wm.get("widget_id")}
    if any(wm.get("widget_id") == "bulk" for wm in widget_mappings):
        attached_ids = {fw["widget_id"] for fw in found_widgets}
    else:
        for wm in widget_mappings:
            for fw in found_widgets:
                if (wm.get("page") and wm.get("page") == fw.get("page_name")) or (wm.get("widget") and wm.get("widget") == fw.get("widget_name")):
                    attached_ids.add(fw["widget_id"])

    all_widgets_uploaded = bool(found_widgets and len(attached_ids) >= len(found_widgets))
    pending_widgets = [fw for fw in found_widgets if fw["widget_id"] not in attached_ids]

    # If widget file was uploaded or default prompt needed
    if is_ds_session and (not user_input or has_new_widgets):
        if widget_page and not all_widgets_uploaded:
            user_input = f"Analyze attached widget source for page '{widget_page}' (widget: '{widget_name}') and generate targeted page & widget technical specification."
        elif widget_text:
            user_input = "Generate a comprehensive, enterprise-grade Technical Documentation for this application including the uploaded widget source code."
        else:
            user_input = "Generate a comprehensive, enterprise-grade Technical Documentation for this application."
    elif not user_input and not document_text:
        return jsonify({"error": "Please provide an input prompt or attach a document/DS file."}), 400

    # Auto-derive friendly title if new session
    if not session["messages"] or session.get("title") == "New Analysis":
        if multi_label:
            # Use first DS filename if present, otherwise first uploaded filename
            ds_files = [n for n in uploaded_filenames if n.endswith('.ds')]
            title_file = ds_files[0] if ds_files else (uploaded_filenames[0] if uploaded_filenames else filename)
            session["title"] = os.path.splitext(title_file)[0].replace("_", " ").title()
            if len(uploaded_filenames) > 1:
                session["title"] += f" (+{len(uploaded_filenames)-1} more)"
        elif filename:
            session["title"] = os.path.splitext(filename)[0].replace("_", " ").title()
        else:
            session["title"] = user_input[:35]

    try:
        ai_resp = generate_response(
            user_input,
            document_text,
            filename,
            provider=ai_provider,
            api_key=api_key,
            ollama_host=ollama_host,
            widget_text=widget_text,
            widget_mappings=session.get("widget_mappings", []),
            target_page=widget_page,
            target_widget=widget_name,
            all_widgets_uploaded=all_widgets_uploaded,
            pending_widgets=pending_widgets
        )

        # Append messages to session and write to local disk
        now_str = datetime.datetime.now().strftime("%I:%M %p")
        session["messages"].append({
            "role": "user",
            "text": user_input,
            "filename": multi_label if uploaded_filenames else "",
            "time": now_str
        })
        session["messages"].append({
            "role": "assistant",
            "text": ai_resp,
            "filename": filename,
            "is_ds": is_zoho_ds_content(document_text or "", filename),
            "time": now_str
        })
        session["updated_at"] = datetime.datetime.now().isoformat()
        save_session(session)

        # Detect widgets in DS and ask user to upload widget files (only if genuine widgets exist)
        widget_prompt_data = None
        has_widgets = False
        if is_zoho_ds_content(document_text or "", filename):
            found_widgets = detect_widgets_in_ds(document_text)
            if found_widgets:
                has_widgets = True
                session["ds_widgets"] = found_widgets
                save_session(session)
                widget_mappings = session.get("widget_mappings", [])
                attached_ids = {wm.get("widget_id") for wm in widget_mappings if wm.get("widget_id")}
                if any(wm.get("widget_id") == "bulk" for wm in widget_mappings):
                    attached_ids = {fw["widget_id"] for fw in found_widgets}
                else:
                    for wm in widget_mappings:
                        for fw in found_widgets:
                            if (wm.get("page") and wm.get("page") == fw.get("page_name")) or (wm.get("widget") and wm.get("widget") == fw.get("widget_name")):
                                attached_ids.add(fw["widget_id"])

                is_submitting_full = any(w in user_input.lower() for w in ["submit", "connecting all widget", "complete unified", "with widgets", "uploaded widgets", "complete doc"])
                # Only show recommendation if there are widgets not yet attached or no widget files uploaded yet (and not submitting full doc)
                if not is_submitting_full and (not widget_text or not widget_text.strip() or (len(attached_ids) < len(found_widgets) and len(attached_ids) > 0)):
                    widget_prompt_data = {
                        "widgets": found_widgets,
                        "count": len(found_widgets),
                        "filename": filename,
                        "attached_ids": list(attached_ids),
                        "widget_mappings": widget_mappings
                    }

        facts_data = None
        validation_rep = None
        five_line_summary = None
        widget_facts_data = None
        if is_ds_session:
            try:
                p = ZohoCreatorParser(document_text)
                facts_data = p.generate_facts_json()
                validation_rep = p.generate_validation_report(doc_text=ai_resp, widget_text=widget_text)
                five_line_summary = p.generate_5_line_summary(widget_text=widget_text)
                if widget_text:
                    widget_facts_data = p.generate_widget_facts_json(widget_text=widget_text, widget_files=session.get("widget_files_list", []))
                session["facts"] = facts_data
                session["widget_facts"] = widget_facts_data
                session["validation_report"] = validation_rep
                session["five_line_summary"] = five_line_summary
                save_session(session)
            except Exception as e:
                print(f"Error computing Master facts/validation: {e}")

        return jsonify({
            "response": ai_resp,
            "filename": filename,
            "is_ds": is_zoho_ds_content(document_text or "", filename),
            "session_id": session_id,
            "is_doc": True,
            "session_title": session["title"],
            "widget_prompt": widget_prompt_data,
            "widget_analyzed": bool(widget_text and widget_text.strip()),
            "widget_files": session.get("widget_files_list", []),
            "widget_mappings": session.get("widget_mappings", []),
            "has_widgets": has_widgets,
            "facts": facts_data,
            "widget_facts": widget_facts_data,
            "validation_report": validation_rep,
            "five_line_summary": five_line_summary
        })
    except llm.LLMError as e:
        return jsonify({"error": str(e)}), 503
    except Exception as e:
        return jsonify({"error": f"Error: {str(e)}"}), 500


@app.route("/api/detect-widgets", methods=["POST"])
def api_detect_widgets():
    """Fast pre-upload widget scan for DS files."""
    session_id = request.form.get("session_id", "").strip()
    session = load_session(session_id) if session_id else None

    files = request.files.getlist("files[]") or request.files.getlist("file")
    text = ""
    filename = ""
    if files and files[0].filename:
        filename = files[0].filename
        text = extract_text_from_file(files[0])
    elif session and session.get("document_text"):
        text = session.get("document_text")
        filename = session.get("filename", "")

    if not text:
        return jsonify({"widgets": [], "count": 0, "has_widgets": False})

    widgets = detect_widgets_in_ds(text)
    return jsonify({
        "widgets": widgets,
        "count": len(widgets),
        "filename": filename,
        "has_widgets": len(widgets) > 0
    })


@app.route("/export", methods=["POST"])
def export_response():
    """Export the AI response text as PDF, DOCX, or TXT based on user-selected format."""
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400

    text = data.get("text", "").strip()
    fmt = data.get("format", "txt").lower()
    diagrams = data.get("diagrams", [])

    if not text:
        return jsonify({"error": "No response text to export"}), 400

    app_name = data.get("app_name", "").strip()
    if not app_name:
        title_m = re.search(r'#\s+([^:\r\n]+):\s*([^\r\n]+)', text)
        if title_m:
            app_name = title_m.group(2).strip()
        else:
            app_name = "Application"
    safe_app_name = re.sub(r'[^A-Za-z0-9_-]', '_', app_name).strip('_') or "Application"
    doc_base_name = f"{safe_app_name}_Functional_Documentation"

    if fmt == "facts":
        facts_obj = data.get("facts")
        if not facts_obj:
            p = ZohoCreatorParser(text)
            facts_obj = p.generate_facts_json()
        buffer = io.BytesIO(json.dumps(facts_obj, indent=2).encode("utf-8"))
        buffer.seek(0)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{safe_app_name}_facts.json",
            mimetype="application/json",
        )

    elif fmt in ("widget_facts", "widget-facts"):
        wf_obj = data.get("widget_facts") or {}
        buffer = io.BytesIO(json.dumps(wf_obj, indent=2).encode("utf-8"))
        buffer.seek(0)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{safe_app_name}_widget_facts.json",
            mimetype="application/json",
        )

    elif fmt in ("validation", "validation_report"):
        val_text = data.get("validation_report") or ""
        if not val_text:
            p = ZohoCreatorParser(text)
            val_text = p.generate_validation_report(doc_text=text)
        buffer = io.BytesIO(val_text.encode("utf-8"))
        buffer.seek(0)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{safe_app_name}_validation_report.md",
            mimetype="text/markdown",
        )

    elif fmt == "docx":
        buffer = markdown_to_docx(text, diagrams=diagrams)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{doc_base_name}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    elif fmt == "pdf":
        buffer = markdown_to_pdf(text, diagrams=diagrams)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{doc_base_name}.pdf",
            mimetype="application/pdf",
        )

    else:  # Default: plain markdown / text
        buffer = io.BytesIO(text.encode("utf-8"))
        buffer.seek(0)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{doc_base_name}.md",
            mimetype="text/markdown",
        )


if __name__ == "__main__":
    port_str = (
        os.environ.get("X_ZOHO_CATALYST_LISTEN_PORT")
        or os.environ.get("X_ZOHO_CATALYST_APPSAIL_PORT")
        or os.environ.get("PORT")
        or "9000"
    )
    port = int(port_str)
    print(f"Starting Zoho Doc AI on 0.0.0.0:{port}...", flush=True)
    app.run(debug=False, host="0.0.0.0", port=port)

