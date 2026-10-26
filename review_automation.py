#!/usr/bin/env python3
"""
AI-Powered Review Automation and Candidate Guidance System  (single-file edition)
=================================================================================
Run:
    pip install flask pandas openpyxl            (pypdf optional, for reading PDF submissions)
    python review_automation.py                  -> starts the server on http://127.0.0.1:5000
    python review_automation.py sample           -> writes sample_candidates.xlsx
    python review_automation.py reminders        -> sends reminders for reviews in the next 24h (use with cron/Task Scheduler)

Configuration (environment variables or a .env file next to this script):
    SECRET_KEY, APP_URL, ADMIN_EMAIL, ADMIN_PASSWORD
    ANTHROPIC_API_KEY, ANTHROPIC_MODEL            (AI features; without a key a simple offline fallback is used)
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_FROM   (without SMTP, emails are only logged to the Outbox page)
    DEFAULT_TZ, PARTIAL_CREDIT, MAX_UPLOAD_MB

Safety: submitted code is NEVER executed. It is only read as text and scanned statically.
"""
import os, re, io, sys, json, uuid, time, ast, hashlib, secrets, sqlite3, smtplib, zipfile, urllib.request
from datetime import datetime, timedelta, date
from email.message import EmailMessage
from concurrent.futures import ThreadPoolExecutor
from functools import wraps
from pathlib import Path

import pandas as pd
from flask import (Flask, g, request, session, redirect, url_for, flash, render_template,
                   send_file, abort, jsonify)
from jinja2 import DictLoader
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

# --------------------------------------------------------------------------- config
BASE = Path(__file__).resolve().parent


def _load_env():
    f = BASE / ".env"
    if f.exists():
        for line in f.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env()


def env(k, d=""):
    return os.environ.get(k, d)


DB_PATH = env("DATABASE_PATH", str(BASE / "review.db"))
UPLOAD_DIR = Path(env("UPLOAD_DIR", str(BASE / "uploads")))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
APP_URL = env("APP_URL", "http://127.0.0.1:5000").rstrip("/")
MODEL = env("ANTHROPIC_MODEL", "claude-sonnet-5-5")
DEFAULT_TZ = env("DEFAULT_TZ", "Asia/Kolkata (IST)")
PARTIAL_CREDIT = float(env("PARTIAL_CREDIT", "0.5"))
MAX_UPLOAD_MB = int(env("MAX_UPLOAD_MB", "25"))
SYNC_JOBS = env("SYNC_JOBS") == "1"           # run background jobs inline (used by tests)
FMT = "%Y-%m-%d %H:%M:%S"
ALLOWED_UPLOAD = {".zip", ".py", ".pdf", ".docx", ".txt", ".md", ".png", ".jpg", ".jpeg", ".js", ".html", ".css", ".java", ".sql", ".json", ".ipynb"}
TEXT_EXT = {".py", ".js", ".ts", ".jsx", ".tsx", ".html", ".css", ".java", ".c", ".cpp", ".cs", ".php", ".sql", ".json",
            ".md", ".txt", ".yml", ".yaml", ".ini", ".sh", ".ipynb"}
CODE_EXT = {".py", ".js", ".ts", ".jsx", ".tsx", ".java", ".php", ".c", ".cpp", ".cs", ".html"}
SKIP_DIRS = ("node_modules/", ".git/", "venv/", "__pycache__/", ".venv/", "dist/", "build/")
MAX_CORPUS = 60000


def _secret():
    if env("SECRET_KEY"):
        return env("SECRET_KEY")
    f = BASE / ".secret_key"
    if not f.exists():
        f.write_text(secrets.token_hex(32))
    return f.read_text().strip()


app = Flask(__name__)
app.config.update(SECRET_KEY=_secret(), MAX_CONTENT_LENGTH=MAX_UPLOAD_MB * 1024 * 1024,
                  SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  PERMANENT_SESSION_LIFETIME=timedelta(hours=8))

# --------------------------------------------------------------------------- database
SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY, email TEXT UNIQUE NOT NULL, name TEXT, password_hash TEXT NOT NULL,
  role TEXT NOT NULL CHECK(role IN ('admin','reviewer','candidate')), active INTEGER DEFAULT 1,
  failed_attempts INTEGER DEFAULT 0, locked_until TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS candidates(
  id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL,
  track TEXT, stage INTEGER DEFAULT 1, deadline TEXT, status TEXT DEFAULT 'Active', active INTEGER DEFAULT 1,
  user_id INTEGER REFERENCES users(id), created_at TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS tasks(
  id INTEGER PRIMARY KEY, title TEXT NOT NULL, problem TEXT, objectives TEXT, modules TEXT, tech TEXT,
  expected TEXT, submission_format TEXT, criteria TEXT, deadline TEXT, allow_resubmit INTEGER DEFAULT 1,
  created_by TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS assignments(
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL REFERENCES tasks(id), candidate_id INTEGER NOT NULL REFERENCES candidates(id),
  assigned_at TEXT, explanation TEXT, explanation_source TEXT, UNIQUE(task_id, candidate_id));
CREATE TABLE IF NOT EXISTS submissions(
  id INTEGER PRIMARY KEY, assignment_id INTEGER NOT NULL REFERENCES assignments(id), version INTEGER NOT NULL,
  github_url TEXT, notes TEXT, submitted_at TEXT, late INTEGER DEFAULT 0, status TEXT DEFAULT 'queued', error TEXT);
CREATE TABLE IF NOT EXISTS submission_files(
  id INTEGER PRIMARY KEY, submission_id INTEGER NOT NULL REFERENCES submissions(id), orig_name TEXT, stored_name TEXT, size INTEGER);
CREATE TABLE IF NOT EXISTS ai_reports(
  id INTEGER PRIMARY KEY, submission_id INTEGER UNIQUE NOT NULL REFERENCES submissions(id), data TEXT, completion REAL,
  source TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS review_schedules(
  id INTEGER PRIMARY KEY, candidate_id INTEGER NOT NULL REFERENCES candidates(id), stage INTEGER NOT NULL,
  review_date TEXT, review_time TEXT, tz TEXT, meet_link TEXT, created_at TEXT, updated_at TEXT, UNIQUE(candidate_id, stage));
CREATE TABLE IF NOT EXISTS review_feedback(
  id INTEGER PRIMARY KEY, submission_id INTEGER NOT NULL REFERENCES submissions(id), reviewer TEXT,
  decision TEXT CHECK(decision IN ('approved','needs_improvement','pending')), comments TEXT, adjusted_score REAL, created_at TEXT);
CREATE TABLE IF NOT EXISTS emails(
  id INTEGER PRIMARY KEY, to_email TEXT, subject TEXT, body TEXT, kind TEXT, status TEXT, error TEXT,
  dedupe_key TEXT, created_at TEXT, sent_at TEXT);
CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY, actor TEXT, action TEXT, detail TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS tokens(
  id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id), token_hash TEXT NOT NULL, purpose TEXT,
  expires_at TEXT, used INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS chats(id INTEGER PRIMARY KEY, assignment_id INTEGER NOT NULL, role TEXT, content TEXT, created_at TEXT);
CREATE INDEX IF NOT EXISTS ix_assign_c ON assignments(candidate_id);
CREATE INDEX IF NOT EXISTS ix_sub_a ON submissions(assignment_id);
CREATE INDEX IF NOT EXISTS ix_email_k ON emails(dedupe_key);
"""


def now():
    return datetime.now().strftime(FMT)


def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH, timeout=30)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys=ON")
    return g.db


@app.teardown_appcontext
def close_db(_=None):
    d = g.pop("db", None)
    if d:
        d.close()


def q(sql, args=(), one=False):
    rows = db().execute(sql, args).fetchall()
    return (rows[0] if rows else None) if one else rows


def ex(sql, args=()):
    cur = db().execute(sql, args)
    db().commit()
    return cur.lastrowid


def init_db():
    c = sqlite3.connect(DB_PATH)
    c.execute("PRAGMA journal_mode=WAL")
    c.executescript(SCHEMA)
    c.commit()
    if not c.execute("SELECT 1 FROM users WHERE role='admin'").fetchone():
        email = env("ADMIN_EMAIL", "admin@example.com").lower()
        pw = env("ADMIN_PASSWORD") or secrets.token_urlsafe(9)
        c.execute("INSERT INTO users(email,name,password_hash,role,created_at) VALUES(?,?,?,?,?)",
                  (email, "Administrator", generate_password_hash(pw), "admin", now()))
        c.commit()
        print(f"\n*** First run: admin account created -> {email} / {pw}  (change it; shown only once) ***\n")
    c.close()


def audit(action, detail=""):
    actor = "system"
    try:
        actor = session.get("email") or "system"
    except RuntimeError:
        pass
    ex("INSERT INTO audit_log(actor,action,detail,created_at) VALUES(?,?,?,?)", (actor, action, detail[:500], now()))


# --------------------------------------------------------------------------- background jobs
EMAIL_POOL = ThreadPoolExecutor(1)
AI_POOL = ThreadPoolExecutor(4)


def job(pool, fn, *a):
    def run():
        with app.app_context():
            try:
                fn(*a)
            except Exception as e:  # never let a worker thread die silently
                app.logger.exception("background job failed: %s", e)
    run() if SYNC_JOBS else pool.submit(run)


# --------------------------------------------------------------------------- email
def queue_email(to, subject, body, kind, dedupe=None):
    """Queue an email. Returns False if an identical notification (same dedupe key) was already sent/queued."""
    if dedupe and q("SELECT 1 FROM emails WHERE dedupe_key=? AND status IN ('queued','sent','console')", (dedupe,), one=True):
        return False
    eid = ex("INSERT INTO emails(to_email,subject,body,kind,status,dedupe_key,created_at) VALUES(?,?,?,?,?,?,?)",
             (to, subject, body, kind, "queued", dedupe, now()))
    job(EMAIL_POOL, deliver_email, eid)
    return True


def deliver_email(eid):
    m = q("SELECT * FROM emails WHERE id=?", (eid,), one=True)
    if not env("SMTP_HOST"):
        print(f"\n[EMAIL not sent - SMTP not configured] To: {m['to_email']}\nSubject: {m['subject']}\n{m['body']}\n")
        ex("UPDATE emails SET status='console', sent_at=? WHERE id=?", (now(), eid))
        return
    msg = EmailMessage()
    msg["From"] = env("SMTP_FROM", env("SMTP_USER"))
    msg["To"] = m["to_email"]
    msg["Subject"] = m["subject"]
    msg.set_content(m["body"])
    err = ""
    for attempt in range(3):
        try:
            port = int(env("SMTP_PORT", "587"))
            if port == 465:
                s = smtplib.SMTP_SSL(env("SMTP_HOST"), port, timeout=30)
            else:
                s = smtplib.SMTP(env("SMTP_HOST"), port, timeout=30)
                s.starttls()
            if env("SMTP_USER"):
                s.login(env("SMTP_USER"), env("SMTP_PASSWORD"))
            s.send_message(msg)
            s.quit()
            ex("UPDATE emails SET status='sent', sent_at=?, error=NULL WHERE id=?", (now(), eid))
            return
        except Exception as e:
            err = str(e)[:300]
            time.sleep(2 * (attempt + 1))
    ex("UPDATE emails SET status='failed', error=? WHERE id=?", (err, eid))


# --------------------------------------------------------------------------- security helpers
HITS = {}


def rate_limited(key, limit=10, window=60):
    t = time.time()
    lst = [x for x in HITS.get(key, []) if t - x < window] + [t]
    HITS[key] = lst
    return len(lst) > limit


def sha(s):
    return hashlib.sha256(s.encode()).hexdigest()


def make_token(uid, purpose, minutes):
    ex("UPDATE tokens SET used=1 WHERE user_id=? AND purpose=? AND used=0", (uid, purpose))
    t = secrets.token_urlsafe(32)
    ex("INSERT INTO tokens(user_id,token_hash,purpose,expires_at) VALUES(?,?,?,?)",
       (uid, sha(t), purpose, (datetime.now() + timedelta(minutes=minutes)).strftime(FMT)))
    return t


def current_user():
    uid = session.get("uid")
    return q("SELECT * FROM users WHERE id=? AND active=1", (uid,), one=True) if uid else None


STAFF = ("admin", "reviewer")


def login_required(*roles):
    def deco(f):
        @wraps(f)
        def w(*a, **k):
            u = current_user()
            if not u:
                return redirect(url_for("admin_login" if request.path.startswith("/admin") else "login"))
            if roles and u["role"] not in roles:
                abort(403)
            g.user = u
            return f(*a, **k)
        return w
    return deco


def my_candidate():
    return q("SELECT * FROM candidates WHERE user_id=?", (g.user["id"],), one=True)


@app.before_request
def csrf_protect():
    if request.method == "POST":
        tok, sent = session.get("csrf"), request.form.get("csrf") or request.headers.get("X-CSRF-Token")
        if not tok or not sent or not secrets.compare_digest(tok, sent):
            abort(400, "CSRF validation failed. Please go back, refresh the page and try again.")


@app.after_request
def sec_headers(r):
    r.headers.setdefault("X-Content-Type-Options", "nosniff")
    r.headers.setdefault("X-Frame-Options", "DENY")
    return r


@app.context_processor
def inject():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return dict(csrf=lambda: session["csrf"], me=current_user(), today=date.today().isoformat())


def first_name(n):
    return (n or "there").split()[0]


# --------------------------------------------------------------------------- AI layer
TONE = """You are a friendly, experienced senior developer mentoring a junior candidate. Write the way a kind mentor talks.
Rules: use simple, warm, natural English. Address the candidate by first name. No robotic phrases such as "As an AI",
"Based on the analysis", "In conclusion" or "insufficient compliance". Short paragraphs, not walls of bullets. Be honest but
encouraging. Never invent requirements that are not in the task. Never claim something is missing or broken unless you have
evidence in the files you were given - otherwise say you could not confirm it and that the reviewer will check. Never claim
that code was executed or tested; it was only read."""


def llm(system, messages, max_tokens=2500):
    key = env("ANTHROPIC_API_KEY")
    if not key:
        return None
    body = json.dumps({"model": MODEL, "max_tokens": max_tokens, "system": system, "messages": messages}).encode()
    for attempt in range(2):
        try:
            req = urllib.request.Request("https://api.anthropic.com/v1/messages", data=body, headers={
                "content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01"})
            with urllib.request.urlopen(req, timeout=120) as r:
                d = json.load(r)
            return "".join(b.get("text", "") for b in d.get("content", []) if b.get("type") == "text")
        except Exception as e:
            app.logger.warning("LLM call failed: %s", e)
            time.sleep(2)
    return None


def task_text(t):
    parts = [("Title", "title"), ("Problem statement", "problem"), ("Objectives", "objectives"), ("Required modules/features", "modules"),
             ("Technology requirements", "tech"), ("Expected output", "expected"), ("Submission format", "submission_format"),
             ("Evaluation criteria", "criteria"), ("Deadline", "deadline")]
    return "\n".join(f"{lbl}: {t[k]}" for lbl, k in parts if t[k])


def parse_modules(text):
    """One requirement per line; optional weight after a pipe:  'Login validation | 2'."""
    out = []
    for line in (text or "").splitlines():
        line = re.sub(r"^\s*(?:[-*\u2022]|\d+[.)])\s*", "", line).strip()
        if not line:
            continue
        name, _, w = line.partition("|")
        try:
            wt = float(w) if w.strip() else 1.0
        except ValueError:
            wt = 1.0
        out.append({"name": name.strip(), "weight": max(wt, 0.0)})
    return out


def fallback_explanation(t, cand):
    mods = "\n".join(f"  - {m['name']}" for m in parse_modules(t["modules"])) or "  - (see the task description)"
    return (f"Hi {first_name(cand['name'])}, here is your task in simple words.\n\n"
            f"What it is: {t['title']}. {t['problem'] or ''}\n\nWhat you need to build:\n{t['objectives'] or ''}\n\n"
            f"Main modules to implement:\n{mods}\n\nTechnology: {t['tech'] or 'as you prefer'}\n\n"
            f"What to submit: {t['expected'] or ''} ({t['submission_format'] or 'see task'})\n\n"
            f"How it will be evaluated: {t['criteria'] or 'based on the modules above'}\n\n"
            "Tip: build the smallest working version first, then add one module at a time and test as you go.\n\n"
            "(AI explanations are not enabled on this server yet, so this is a plain summary of the task. Use 'Ask a question' once enabled.)")


def gen_explanation(aid):
    a = q("SELECT * FROM assignments WHERE id=?", (aid,), one=True)
    t = q("SELECT * FROM tasks WHERE id=?", (a["task_id"],), one=True)
    c = q("SELECT * FROM candidates WHERE id=?", (a["candidate_id"],), one=True)
    prompt = (f"Explain this task to {first_name(c['name'])} in simple, natural, conversational English. Cover, in a flowing way "
              "(light headings are fine): 1) what the task means, 2) why it is required, 3) what to build, 4) the main modules, "
              "5) step-by-step guidance on how to approach it (do NOT write the full solution), 6) expected output and submission "
              "requirements, 7) common mistakes to avoid, 8) how it will be evaluated. Only use requirements present in the task.\n\n"
              + task_text(t))
    text = llm(TONE, [{"role": "user", "content": prompt}], 2500)
    ex("UPDATE assignments SET explanation=?, explanation_source=? WHERE id=?",
       (text or fallback_explanation(t, c), "ai" if text else "fallback", aid))


# ---- submission reading + static analysis (never executes anything)
def static_scan(name, text):
    ext = Path(name).suffix.lower()
    if ext not in CODE_EXT:
        return []
    out = []
    if ext == ".py":
        try:
            ast.parse(text)
        except SyntaxError as e:
            out.append(dict(file=name, line=e.lineno, title="Python syntax error", severity="high", confirmed=True, detail=str(e.msg)))
    pats = [(r"\b(eval|exec)\s*\(", "Use of eval/exec", "high", "Executing dynamic strings can run unexpected code."),
            (r"(?i)\b\w*(password|passwd|secret|api_?key|token)\w*\s*[=:]\s*['\"][^'\"\s]{4,}['\"]", "Possible hard-coded secret", "high",
             "Secrets should live in environment variables."),
            (r"debug\s*=\s*True", "Debug mode switched on", "medium", "Debug mode should be off outside development."),
            (r"(?i)(select|insert|update|delete)\b[^\n]*(\"\s*\+|'\s*\+|%s[^\n]*%\s*\w|\.format\(|\bf['\"])",
             "SQL query built with string formatting", "high", "This can allow SQL injection; use parameterised queries.")]
    for i, line in enumerate(text.splitlines(), 1):
        for rx, title, sev, det in pats:
            if re.search(rx, line):
                out.append(dict(file=name, line=i, title=title, severity=sev, confirmed=False, detail=det + " (pattern match - verify manually)"))
    return out[:15]


def _read_zip(fobj, label, add, notes, tree):
    try:
        with zipfile.ZipFile(fobj) as z:
            infos = [i for i in z.infolist() if not i.is_dir()]
            if sum(i.file_size for i in infos) > 150 * 1024 * 1024:
                notes.append(f"{label}: archive is too large when unpacked; skipped for safety.")
                return
            for i in infos:
                n = i.filename
                if any(s in n.lower() for s in SKIP_DIRS) or ".." in n.split("/"):
                    continue
                tree.append(n)
                if Path(n).suffix.lower() in TEXT_EXT and i.file_size < 300 * 1024:
                    add(f"{label}/{n}", z.read(i).decode("utf-8", "ignore"))
    except zipfile.BadZipFile:
        notes.append(f"{label}: not a valid zip file.")


def build_corpus(sid):
    s = q("SELECT * FROM submissions WHERE id=?", (sid,), one=True)
    parts, notes, tree, static, total = [], [], [], [], [0]

    def add(name, text):
        if total[0] >= MAX_CORPUS:
            return
        text = text[:12000]
        total[0] += len(text)
        parts.append(f"### FILE: {name}\n{text}")
        static.extend(static_scan(name, text))

    for f in q("SELECT * FROM submission_files WHERE submission_id=?", (sid,)):
        p, ext = UPLOAD_DIR / f["stored_name"], Path(f["orig_name"]).suffix.lower()
        if ext == ".zip":
            _read_zip(str(p), f["orig_name"], add, notes, tree)
        elif ext in TEXT_EXT:
            tree.append(f["orig_name"])
            add(f["orig_name"], p.read_text("utf-8", "ignore"))
        elif ext == ".pdf":
            try:
                from pypdf import PdfReader
                add(f["orig_name"], "\n".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages[:30]))
            except Exception:
                notes.append(f"{f['orig_name']}: PDF text could not be read (install pypdf); reviewer should open it manually.")
        else:
            notes.append(f"{f['orig_name']}: {ext} file not analysed as text (reviewer should open it manually).")
    if s["github_url"]:
        m = re.match(r"https://github\.com/([\w.-]+)/([\w.-]+)", s["github_url"])
        try:
            if not m:
                raise ValueError("not a github repo url")
            req = urllib.request.Request(f"https://api.github.com/repos/{m.group(1)}/{m.group(2).removesuffix('.git')}/zipball",
                                         headers={"User-Agent": "review-automation"})
            with urllib.request.urlopen(req, timeout=40) as r:
                data = r.read(15 * 1024 * 1024 + 1)
            if len(data) > 15 * 1024 * 1024:
                raise ValueError("repository larger than 15 MB")
            _read_zip(io.BytesIO(data), "github", add, notes, tree)
        except Exception as e:
            notes.append(f"GitHub repository could not be read ({e}). Only the link is available; reviewer should open it manually.")
    if tree:
        parts.insert(0, "### FILE TREE\n" + "\n".join(tree[:200]))
    return "\n\n".join(parts), notes, static[:30], bool(tree)


STATUS_SCORE = {"done": 1.0, "partial": PARTIAL_CREDIT, "missing": 0.0, "unverified": 0.0}


def completion_of(reqs):
    tw = sum(r["weight"] for r in reqs)
    return round(100 * sum(r["weight"] * STATUS_SCORE.get(r["status"], 0) for r in reqs) / tw, 1) if tw else 0.0


def heuristic_report(t, cand, reqs, corpus, static, notes, readable):
    low = corpus.lower()
    out = []
    for r in reqs:
        words = [w for w in re.findall(r"[a-z]{4,}", r["name"].lower())] or [r["name"].lower()]
        hit = sum(1 for w in words if w in low) / len(words)
        st = "done" if hit >= 0.7 else "partial" if hit >= 0.35 else "unverified"
        out.append(dict(name=r["name"], weight=r["weight"], status=st, confidence="needs_manual_verification",
                        evidence="Keyword match only (no AI model configured)." if st != "unverified" else "No clear evidence found by the simple keyword check."))
    issues = [dict(title=s["title"], type="static", severity=s["severity"], confirmed=s["confirmed"],
                   location=f"{s['file']}:{s['line']}", why_it_matters=s["detail"], how_to_fix="", snippet="") for s in static]
    fn = first_name(cand["name"])
    msg = (f"Hi {fn}, thanks for submitting your work. I only ran a basic automatic check on it (no AI model is connected right now), "
           "so please treat this as a first look and not a final verdict. Your reviewer will go through everything properly.\n\n"
           "Below you can see which modules look present and which ones I couldn't find clear evidence for. Where I'm unsure I've "
           "said so rather than guessing. A good next step is to open each unchecked module and make sure it's clearly visible in your files.")
    return dict(message=msg, requirements=out, issues=issues, next_steps=[f"Double-check: {r['name']}" for r in out if r["status"] != "done"][:6],
                documentation_gaps=[], example_fixes=[], notes=notes, tests_executed=False)


def ai_report(t, cand, reqs, corpus, static, notes, readable, prev):
    if not env("ANTHROPIC_API_KEY"):
        return heuristic_report(t, cand, reqs, corpus, static, notes, readable), "heuristic"
    schema = {"message": "friendly mentor feedback for the candidate in 3-6 short paragraphs: what they did well, what is missing and why it matters, how to fix, what to do next",
              "requirements": [{"name": "must match a required module name exactly", "status": "done|partial|missing|unverified",
                                "confidence": "confirmed|needs_manual_verification", "evidence": "file/function you saw, or why you could not tell"}],
              "issues": [{"title": "", "type": "bug|validation|security|quality|performance|error_handling", "severity": "low|medium|high",
                          "confirmed": True, "location": "file:line if known", "why_it_matters": "", "how_to_fix": "", "snippet": "optional short example"}],
              "next_steps": ["..."], "documentation_gaps": ["..."]}
    prompt = (f"Analyse {cand['name']}'s submission against the task. The code was only READ, never run - do not say it was tested.\n"
              "Use status 'missing' only when you looked at the relevant files and found clear evidence it is absent; if unsure use 'unverified'. "
              "Mark 'confirmed' only for things you can point to in the files. Return ONLY valid JSON matching this shape (no markdown):\n"
              f"{json.dumps(schema)}\n\nREQUIRED MODULES (use these exact names):\n" + "\n".join(f"- {r['name']}" for r in reqs) +
              f"\n\nTASK:\n{task_text(t)}\n\nSTATIC FINDINGS (automatic scan; 'confirmed' false means pattern match only):\n{json.dumps(static)}\n"
              f"\nNOTES ABOUT THE SUBMISSION: {json.dumps(notes)}\n"
              + (f"\nPREVIOUS FEEDBACK GIVEN TO THIS CANDIDATE (mention progress naturally):\n{prev.get('message', '')[:1500]}\n" if prev else "")
              + f"\nSUBMITTED FILES:\n{corpus or '(no readable files)'}")
    text = llm(TONE, [{"role": "user", "content": prompt}], 4000)
    try:
        d = json.loads(text[text.index("{"): text.rindex("}") + 1])
        assert isinstance(d.get("requirements"), list)
    except Exception:
        return heuristic_report(t, cand, reqs, corpus, static, notes, readable), "heuristic (AI reply could not be parsed)"
    by = {str(r.get("name", "")).strip().lower(): r for r in d["requirements"]}
    final = []
    for r in reqs:
        m = by.get(r["name"].lower(), {})
        st = m.get("status") if m.get("status") in STATUS_SCORE else "unverified"
        conf = m.get("confidence", "needs_manual_verification")
        if (st == "missing" and conf != "confirmed") or (not readable and st != "done"):
            st = "unverified"                       # evidence guard: never assert "missing" without evidence
        final.append(dict(name=r["name"], weight=r["weight"], status=st, confidence=conf, evidence=m.get("evidence", "")))
    return dict(message=d.get("message", ""), requirements=final, issues=d.get("issues", []) or [],
                next_steps=d.get("next_steps", []) or [], documentation_gaps=d.get("documentation_gaps", []) or [],
                example_fixes=[], notes=notes, tests_executed=False), "ai"


RANK = {"missing": 0, "unverified": 0, "partial": 1, "done": 2}


def compare_reports(prev, cur):
    pr = {r["name"].lower(): r["status"] for r in prev["requirements"]}
    improved = [r["name"] for r in cur["requirements"] if RANK[r["status"]] > RANK.get(pr.get(r["name"].lower(), "missing"), 0)]
    pi = {i["title"].strip().lower(): i["title"] for i in prev.get("issues", [])}
    ci = {i["title"].strip().lower(): i["title"] for i in cur.get("issues", [])}
    return dict(requirements_improved=improved, issues_fixed=[pi[k] for k in pi if k not in ci],
                issues_remaining=[ci[k] for k in ci if k in pi], issues_new=[ci[k] for k in ci if k not in pi])


def analyse_submission(sid):
    s = q("SELECT * FROM submissions WHERE id=?", (sid,), one=True)
    a = q("SELECT * FROM assignments WHERE id=?", (s["assignment_id"],), one=True)
    t = q("SELECT * FROM tasks WHERE id=?", (a["task_id"],), one=True)
    c = q("SELECT * FROM candidates WHERE id=?", (a["candidate_id"],), one=True)
    ex("UPDATE submissions SET status='analysing' WHERE id=?", (sid,))
    try:
        reqs = parse_modules(t["modules"]) or parse_modules(t["objectives"])
        corpus, notes, static, readable = build_corpus(sid)
        prevrow = q("SELECT r.* FROM ai_reports r JOIN submissions s ON s.id=r.submission_id WHERE s.assignment_id=? AND s.version<? "
                    "ORDER BY s.version DESC LIMIT 1", (a["id"], s["version"]), one=True)
        prev = json.loads(prevrow["data"]) if prevrow else None
        rep, source = ai_report(t, c, reqs, corpus, static, notes, readable, prev)
        rep["completion"] = completion_of(rep["requirements"])
        rep["comparison"] = compare_reports(prev, rep) if prev else None
        rep["previous_completion"] = prevrow["completion"] if prevrow else None
        ex("INSERT OR REPLACE INTO ai_reports(submission_id,data,completion,source,created_at) VALUES(?,?,?,?,?)",
           (sid, json.dumps(rep), rep["completion"], source, now()))
        ex("UPDATE submissions SET status='done', error=NULL WHERE id=?", (sid,))
        audit("ai_analysis", f"submission {sid} ({source})")
        queue_email(c["email"], f"Your feedback is ready: {t['title']}",
                    f"Hi {first_name(c['name'])},\n\nI've gone through your submission (version {s['version']}) and your feedback is ready. "
                    f"Log in to read it:\n{APP_URL}/submission/{sid}\n\nYour reviewer will also look at it before the review.\n",
                    "feedback", dedupe=f"feedback:{sid}")
    except Exception as e:
        app.logger.exception("analysis failed")
        ex("UPDATE submissions SET status='failed', error=? WHERE id=?", (str(e)[:300], sid))


# --------------------------------------------------------------------------- notifications
def activation_email(cand, user_id):
    tok = make_token(user_id, "activate", 72 * 60)
    queue_email(cand["email"], "Set up your Review Portal account",
                f"Hi {first_name(cand['name'])},\n\nYour account for the Review Portal is ready. Please set your password using this link "
                f"(valid for 72 hours, can be used once):\n{APP_URL}/reset/{tok}\n\nAfter that, log in at {APP_URL}/login with this email address.\n",
                "activation")


def schedule_email(cand, rv, kind):
    a = q("SELECT t.title, (SELECT status FROM submissions s WHERE s.assignment_id=a.id ORDER BY version DESC LIMIT 1) AS sub "
          "FROM assignments a JOIN tasks t ON t.id=a.task_id WHERE a.candidate_id=? ORDER BY a.id DESC LIMIT 1", (cand["id"],), one=True)
    task = a["title"] if a else "(task not yet assigned)"
    subst = "Not submitted yet" if not a or not a["sub"] else "Submitted"
    head = "Your review has been scheduled" if kind == "new" else "Update: your review details have changed"
    body = (f"Hi {first_name(cand['name'])},\n\n{head}.\n\nReview stage: Review {rv['stage']}\nDate: {rv['review_date']}\n"
            f"Time: {rv['review_time']} ({rv['tz']})\nGoogle Meet: {rv['meet_link']}\nTask / project: {task}\nSubmission status: {subst}\n\n"
            "How to prepare: keep your project running and ready to demo, be ready to walk through your code and explain your decisions, "
            "and join the Meet link 5 minutes early with a working microphone.\n")
    key = f"sched:{cand['id']}:{rv['stage']}:{rv['review_date']}:{rv['review_time']}:{rv['tz']}:{sha(rv['meet_link'] or '')[:8]}"
    queue_email(cand["email"], f"{'Review scheduled' if kind == 'new' else 'Review updated'}: Review {rv['stage']} on {rv['review_date']}", body, "schedule_" + kind, dedupe=key)


def upsert_schedule(cid, stage, d, t, tz, link):
    cand = q("SELECT * FROM candidates WHERE id=?", (cid,), one=True)
    row = q("SELECT * FROM review_schedules WHERE candidate_id=? AND stage=?", (cid, stage), one=True)
    if row and (row["review_date"], row["review_time"], row["tz"], row["meet_link"]) == (d, t, tz, link):
        return False
    if row:
        ex("UPDATE review_schedules SET review_date=?,review_time=?,tz=?,meet_link=?,updated_at=? WHERE id=?", (d, t, tz, link, now(), row["id"]))
    else:
        ex("INSERT INTO review_schedules(candidate_id,stage,review_date,review_time,tz,meet_link,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)",
           (cid, stage, d, t, tz, link, now(), now()))
    rv = q("SELECT * FROM review_schedules WHERE candidate_id=? AND stage=?", (cid, stage), one=True)
    schedule_email(cand, rv, "updated" if row else "new")
    audit("schedule", f"candidate {cid} stage {stage} {d} {t}")
    return True


def send_reminder(rv):
    c = q("SELECT * FROM candidates WHERE id=?", (rv["candidate_id"],), one=True)
    body = (f"Hi {first_name(c['name'])},\n\nFriendly reminder about your Review {rv['stage']}:\n{rv['review_date']} at {rv['review_time']} "
            f"({rv['tz']})\nGoogle Meet: {rv['meet_link']}\n\nPlease be ready to demo your work. All the best!\n")
    return queue_email(c["email"], f"Reminder: Review {rv['stage']} on {rv['review_date']}", body, "reminder",
                       dedupe=f"remind:{rv['id']}:{rv['review_date']}:{rv['review_time']}:{date.today()}")


def assign_task(tid, cids):
    t = q("SELECT * FROM tasks WHERE id=?", (tid,), one=True)
    n = 0
    for cid in cids:
        c = q("SELECT * FROM candidates WHERE id=? AND active=1", (cid,), one=True)
        if not c or q("SELECT 1 FROM assignments WHERE task_id=? AND candidate_id=?", (tid, cid), one=True):
            continue
        aid = ex("INSERT INTO assignments(task_id,candidate_id,assigned_at) VALUES(?,?,?)", (tid, cid, now()))
        n += 1
        job(AI_POOL, gen_explanation, aid)
        queue_email(c["email"], f"New task assigned: {t['title']}",
                    f"Hi {first_name(c['name'])},\n\nYou have been assigned a new task: {t['title']}\nDeadline: {t['deadline'] or 'not set'}\n\n"
                    f"Log in to read a simple explanation of the task and to ask questions:\n{APP_URL}/task/{aid}\n", "assignment", dedupe=f"assign:{aid}")
    audit("assign_task", f"task {tid} -> {n} candidate(s)")
    return n


def create_candidate(code, name, email, track, stage, deadline, status):
    uid = ex("INSERT INTO users(email,name,password_hash,role,created_at) VALUES(?,?,?,?,?)",
             (email, name, generate_password_hash(secrets.token_urlsafe(24)), "candidate", now()))
    cid = ex("INSERT INTO candidates(code,name,email,track,stage,deadline,status,user_id,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
             (code, name, email, track, stage, deadline, status, uid, now(), now()))
    activation_email(q("SELECT * FROM candidates WHERE id=?", (cid,), one=True), uid)
    return cid


# --------------------------------------------------------------------------- auth routes
EMAIL_RX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def do_login(kind):
    if request.method == "POST":
        if rate_limited("login:" + (request.remote_addr or ""), 15):
            flash("Too many attempts. Please wait a minute and try again.", "danger")
            return render_template("login.html", kind=kind)
        email, pw = request.form.get("email", "").strip().lower(), request.form.get("password", "")
        u = q("SELECT * FROM users WHERE email=?", (email,), one=True)
        allowed = ("candidate",) if kind == "candidate" else STAFF
        if u and u["role"] in allowed and u["active"]:
            if u["locked_until"] and u["locked_until"] > now():
                flash("This account is temporarily locked after too many wrong attempts. Try again in a few minutes or reset your password.", "danger")
            elif check_password_hash(u["password_hash"], pw):
                ex("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?", (u["id"],))
                session.clear()
                session.permanent = True
                session.update(uid=u["id"], email=u["email"], csrf=secrets.token_hex(16))
                audit("login", u["role"])
                return redirect(url_for("admin_dash" if kind == "admin" else "dashboard"))
            else:
                fails = u["failed_attempts"] + 1
                lock = (datetime.now() + timedelta(minutes=10)).strftime(FMT) if fails >= 5 else None
                ex("UPDATE users SET failed_attempts=?, locked_until=? WHERE id=?", (0 if lock else fails, lock, u["id"]))
                flash("That email or password is not correct. If you've forgotten it, use 'Forgot password?' below.", "danger")
        else:
            check_password_hash(generate_password_hash("x"), pw)   # keep timing similar
            flash("That email or password is not correct. If you've forgotten it, use 'Forgot password?' below.", "danger")
    return render_template("login.html", kind=kind)


@app.route("/")
def index():
    u = current_user()
    if not u:
        return redirect(url_for("login"))
    return redirect(url_for("admin_dash" if u["role"] in STAFF else "dashboard"))


@app.route("/login", methods=["GET", "POST"])
def login():
    return do_login("candidate")


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    return do_login("admin")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/forgot", methods=["GET", "POST"])
def forgot():
    if request.method == "POST":
        if not rate_limited("forgot:" + (request.remote_addr or ""), 5, 300):
            email = request.form.get("email", "").strip().lower()
            u = q("SELECT * FROM users WHERE email=? AND active=1", (email,), one=True)
            if u:
                tok = make_token(u["id"], "reset", 30)
                queue_email(email, "Reset your Review Portal password",
                            f"Hi {first_name(u['name'])},\n\nUse this link to set a new password (valid for 30 minutes, one use only):\n{APP_URL}/reset/{tok}\n\n"
                            "If you didn't ask for this, you can ignore this email.\n", "reset")
                audit("password_reset_requested", email)
        flash("If that email is registered, a reset link has been sent. Please check your inbox.", "info")   # same answer either way
        return redirect(url_for("forgot"))
    return render_template("forgot.html")


@app.route("/reset/<token>", methods=["GET", "POST"])
def reset(token):
    tk = q("SELECT * FROM tokens WHERE token_hash=? AND used=0 AND expires_at>?", (sha(token), now()), one=True)
    if not tk:
        return render_template("message.html", title="Link expired", msg="This link is invalid, already used or expired. Please request a new one.",
                               link=url_for("forgot"), label="Request a new link"), 400
    if request.method == "POST":
        p1, p2 = request.form.get("password", ""), request.form.get("password2", "")
        if p1 != p2 or len(p1) < 8 or not re.search(r"\d", p1) or not re.search(r"[A-Za-z]", p1):
            flash("Passwords must match and be at least 8 characters with letters and numbers.", "danger")
        else:
            ex("UPDATE users SET password_hash=?, failed_attempts=0, locked_until=NULL WHERE id=?", (generate_password_hash(p1), tk["user_id"]))
            ex("UPDATE tokens SET used=1 WHERE id=?", (tk["id"],))
            audit("password_set", f"user {tk['user_id']}")
            flash("Password saved. You can log in now.", "success")
            return redirect(url_for("login"))
    return render_template("reset.html")


# --------------------------------------------------------------------------- admin: dashboard & candidates
@app.route("/admin")
@login_required(*STAFF)
def admin_dash():
    one = lambda s: q(s, one=True)[0]
    stats = dict(total=one("SELECT COUNT(*) FROM candidates WHERE active=1"), assigned=one("SELECT COUNT(*) FROM assignments"),
                 submitted=one("SELECT COUNT(DISTINCT assignment_id) FROM submissions"),
                 reviewed=one("SELECT COUNT(DISTINCT submission_id) FROM review_feedback WHERE decision IN ('approved','needs_improvement')"))
    stats["pending"] = stats["assigned"] - stats["submitted"]
    upcoming = q("SELECT r.*, c.name FROM review_schedules r JOIN candidates c ON c.id=r.candidate_id WHERE r.review_date>=? ORDER BY r.review_date, r.review_time LIMIT 8", (date.today().isoformat(),))
    recent = q("SELECT s.*, c.name, t.title, r.completion FROM submissions s JOIN assignments a ON a.id=s.assignment_id JOIN candidates c ON c.id=a.candidate_id "
               "JOIN tasks t ON t.id=a.task_id LEFT JOIN ai_reports r ON r.submission_id=s.id ORDER BY s.id DESC LIMIT 8")
    return render_template("admin_dash.html", stats=stats, upcoming=upcoming, recent=recent)


@app.route("/admin/candidates")
@login_required(*STAFF)
def admin_candidates():
    s, track, stage = request.args.get("q", "").strip(), request.args.get("track", ""), request.args.get("stage", "")
    sql, args = "SELECT * FROM candidates WHERE 1=1", []
    if s:
        sql += " AND (name LIKE ? OR email LIKE ? OR code LIKE ?)"
        args += [f"%{s}%"] * 3
    if track:
        sql += " AND track=?"
        args.append(track)
    if stage.isdigit():
        sql += " AND stage=?"
        args.append(int(stage))
    rows = q(sql + " ORDER BY id DESC", args)
    tracks = [r[0] for r in q("SELECT DISTINCT track FROM candidates WHERE track IS NOT NULL AND track!=''")]
    return render_template("admin_candidates.html", rows=rows, tracks=tracks, f=dict(q=s, track=track, stage=stage))


def _cand_form():
    f = request.form
    email, name = f.get("email", "").strip().lower(), f.get("name", "").strip()
    if not name or not EMAIL_RX.match(email):
        return None
    stage = int(f["stage"]) if f.get("stage", "").isdigit() else 1
    return dict(code=f.get("code", "").strip() or None, name=name, email=email, track=f.get("track", "").strip(),
                stage=stage, deadline=f.get("deadline") or None, status=f.get("status", "Active"))


@app.route("/admin/candidates/new", methods=["GET", "POST"])
@login_required(*STAFF)
def admin_candidate_new():
    if request.method == "POST":
        d = _cand_form()
        if not d:
            flash("Name and a valid email are required.", "danger")
        elif q("SELECT 1 FROM candidates WHERE email=? OR (code IS NOT NULL AND code=?)", (d["email"], d["code"]), one=True):
            flash("A candidate with that email or ID already exists.", "danger")
        else:
            cid = create_candidate(d["code"], d["name"], d["email"], d["track"], d["stage"], d["deadline"], d["status"])
            audit("candidate_created", d["email"])
            flash("Candidate added and activation email queued.", "success")
            return redirect(url_for("admin_candidate_view", cid=cid))
    return render_template("candidate_form.html", c=None)


@app.route("/admin/candidates/<int:cid>/edit", methods=["GET", "POST"])
@login_required(*STAFF)
def admin_candidate_edit(cid):
    c = q("SELECT * FROM candidates WHERE id=?", (cid,), one=True) or abort(404)
    if request.method == "POST":
        d = _cand_form()
        dup = d and q("SELECT 1 FROM candidates WHERE id!=? AND (email=? OR (code IS NOT NULL AND code=?))", (cid, d["email"], d["code"]), one=True)
        if not d or dup:
            flash("Check the details: name and a valid, unique email/ID are required.", "danger")
        else:
            ex("UPDATE candidates SET code=?,name=?,email=?,track=?,stage=?,deadline=?,status=?,updated_at=? WHERE id=?",
               (d["code"], d["name"], d["email"], d["track"], d["stage"], d["deadline"], d["status"], now(), cid))
            ex("UPDATE users SET email=?, name=? WHERE id=?", (d["email"], d["name"], c["user_id"]))
            audit("candidate_edited", d["email"])
            flash("Saved.", "success")
            return redirect(url_for("admin_candidate_view", cid=cid))
    return render_template("candidate_form.html", c=c)


@app.route("/admin/candidates/<int:cid>")
@login_required(*STAFF)
def admin_candidate_view(cid):
    c = q("SELECT * FROM candidates WHERE id=?", (cid,), one=True) or abort(404)
    asg = q("SELECT a.*, t.title FROM assignments a JOIN tasks t ON t.id=a.task_id WHERE a.candidate_id=?", (cid,))
    subs = q("SELECT s.*, t.title, r.completion FROM submissions s JOIN assignments a ON a.id=s.assignment_id JOIN tasks t ON t.id=a.task_id "
             "LEFT JOIN ai_reports r ON r.submission_id=s.id WHERE a.candidate_id=? ORDER BY s.id DESC", (cid,))
    revs = q("SELECT * FROM review_schedules WHERE candidate_id=? ORDER BY stage", (cid,))
    return render_template("candidate_view.html", c=c, asg=asg, subs=subs, revs=revs, tz=DEFAULT_TZ)


@app.route("/admin/candidates/<int:cid>/toggle", methods=["POST"])
@login_required("admin")
def admin_candidate_toggle(cid):
    c = q("SELECT * FROM candidates WHERE id=?", (cid,), one=True) or abort(404)
    v = 0 if c["active"] else 1
    ex("UPDATE candidates SET active=? WHERE id=?", (v, cid))
    ex("UPDATE users SET active=? WHERE id=?", (v, c["user_id"]))
    audit("candidate_active", f"{c['email']} -> {v}")
    flash("Account " + ("activated." if v else "deactivated."), "success")
    return redirect(request.referrer or url_for("admin_candidates"))


@app.route("/admin/candidates/<int:cid>/resend", methods=["POST"])
@login_required(*STAFF)
def admin_resend(cid):
    c = q("SELECT * FROM candidates WHERE id=?", (cid,), one=True) or abort(404)
    activation_email(c, c["user_id"])
    audit("activation_resent", c["email"])
    flash("A fresh set-password link has been emailed.", "success")
    return redirect(url_for("admin_candidate_view", cid=cid))


@app.route("/admin/candidates/<int:cid>/schedule", methods=["POST"])
@login_required(*STAFF)
def admin_schedule(cid):
    f = request.form
    stage = int(f["stage"]) if f.get("stage", "").isdigit() else 1
    link = f.get("meet_link", "").strip()
    if not (f.get("date") and re.match(r"^\d{1,2}:\d{2}$", f.get("time", "")) and link.startswith("https://")):
        flash("Enter a date, a time (HH:MM) and a Google Meet link starting with https://", "danger")
    else:
        changed = upsert_schedule(cid, stage, f["date"], f["time"], f.get("tz") or DEFAULT_TZ, link)
        flash("Review saved and the candidate has been emailed." if changed else "No change - nothing to send.", "success")
    return redirect(url_for("admin_candidate_view", cid=cid))


# --------------------------------------------------------------------------- admin: Excel import
COLS = {"candidate id": "code", "candidate name": "name", "email address": "email", "assigned task or track": "track",
        "current review stage": "stage", "task deadline": "deadline", "review date": "rdate", "review time": "rtime",
        "google meet link": "link", "candidate status": "status"}
SAMPLE_ROWS = [[f"C{100 + i}", n, f"{n.split()[0].lower()}{i}@example.com", "Web Development", 1, "2026-11-15",
                "2026-11-20", "10:30", "https://meet.google.com/abc-defg-hij", "Active"]
               for i, n in enumerate(["Aarav Kumar", "Priya Sharma", "Karthik Raj", "Divya Lakshmi", "Mohammed Rafi"], 1)]


def sample_df():
    return pd.DataFrame(SAMPLE_ROWS, columns=["Candidate ID", "Candidate Name", "Email Address", "Assigned Task or Track", "Current Review Stage",
                                              "Task Deadline", "Review Date", "Review Time", "Google Meet Link", "Candidate Status"])


def _d(v):
    if not v:
        return None
    x = pd.to_datetime(v, errors="coerce")
    return None if pd.isna(x) else x.strftime("%Y-%m-%d")


@app.route("/admin/sample-excel")
@login_required(*STAFF)
def admin_sample():
    bio = io.BytesIO()
    sample_df().to_excel(bio, index=False)
    bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="sample_candidates.xlsx")


@app.route("/admin/import", methods=["GET", "POST"])
@login_required("admin")
def admin_import():
    result = None
    if request.method == "POST":
        f = request.files.get("file")
        if not f or not f.filename.lower().endswith((".xlsx", ".xls")):
            flash("Please choose an .xlsx file.", "danger")
            return render_template("admin_import.html", result=None)
        try:
            df = pd.read_excel(f, dtype=str)
        except Exception as e:
            flash(f"Could not read that file: {e}", "danger")
            return render_template("admin_import.html", result=None)
        df.columns = [COLS.get(str(c).strip().lower(), str(c)) for c in df.columns]
        missing = [c for c in ("code", "name", "email") if c not in df.columns]
        if missing:
            flash("Missing required column(s): " + ", ".join(k for k, v in COLS.items() if v in missing), "danger")
            return render_template("admin_import.html", result=None)
        created = updated = 0
        errors, seen = [], set()
        for i, row in df.iterrows():
            n = i + 2
            v = {k: (str(row[k]).strip() if k in df.columns and pd.notna(row[k]) else "") for k in COLS.values()}
            v["email"] = v["email"].lower()
            prob = []
            if not v["code"]: prob.append("Candidate ID missing")
            if not v["name"]: prob.append("Candidate Name missing")
            if not EMAIL_RX.match(v["email"]): prob.append("Email invalid or missing")
            if v["email"] in seen: prob.append("Duplicate email inside this file")
            dl, rd = _d(v["deadline"]), _d(v["rdate"])
            if v["deadline"] and not dl: prob.append("Task Deadline is not a valid date")
            if v["rdate"] and not rd: prob.append("Review Date is not a valid date")
            if v["rtime"] and not re.match(r"^\d{1,2}:\d{2}", v["rtime"]): prob.append("Review Time must look like 10:30")
            if v["link"] and not v["link"].startswith("https://"): prob.append("Google Meet Link must start with https://")
            m = re.search(r"\d+", v["stage"])
            stage = int(m.group()) if m else 1
            by_email = q("SELECT * FROM candidates WHERE email=?", (v["email"],), one=True)
            by_code = q("SELECT * FROM candidates WHERE code=?", (v["code"],), one=True)
            if by_code and by_email is None:
                prob.append("Candidate ID already belongs to a different email")
            if prob:
                errors.append({"Row": n, "Candidate ID": v["code"], "Email": v["email"], "Problems": "; ".join(prob)})
                continue
            seen.add(v["email"])
            status = v["status"] or "Active"
            if by_email:
                ex("UPDATE candidates SET code=?,name=?,track=?,stage=?,deadline=COALESCE(?,deadline),status=?,updated_at=? WHERE id=?",
                   (v["code"], v["name"], v["track"], stage, dl, status, now(), by_email["id"]))
                cid, updated = by_email["id"], updated + 1
            else:
                cid, created = create_candidate(v["code"], v["name"], v["email"], v["track"], stage, dl, status), created + 1
            tk = v["track"] and q("SELECT id FROM tasks WHERE lower(title)=lower(?)", (v["track"],), one=True)
            if tk:
                assign_task(tk["id"], [cid])
            if rd and v["rtime"] and v["link"]:
                upsert_schedule(cid, stage, rd, v["rtime"][:5].zfill(5), DEFAULT_TZ, v["link"])
        err_file = None
        if errors:
            err_file = f"import_errors_{uuid.uuid4().hex[:8]}.xlsx"
            pd.DataFrame(errors).to_excel(UPLOAD_DIR / err_file, index=False)
        audit("excel_import", f"created {created}, updated {updated}, errors {len(errors)}")
        result = dict(created=created, updated=updated, errors=errors[:50], error_count=len(errors), err_file=err_file, total=len(df))
    return render_template("admin_import.html", result=result)


@app.route("/admin/import/errors/<name>")
@login_required("admin")
def admin_import_errors(name):
    if not re.fullmatch(r"import_errors_[0-9a-f]{8}\.xlsx", name) or not (UPLOAD_DIR / name).exists():
        abort(404)
    return send_file(UPLOAD_DIR / name, as_attachment=True)


# --------------------------------------------------------------------------- admin: tasks, submissions, reviews, reports
@app.route("/admin/tasks", methods=["GET", "POST"])
@login_required(*STAFF)
def admin_tasks():
    if request.method == "POST":
        f = request.form
        if not f.get("title", "").strip():
            flash("Task title is required.", "danger")
        else:
            tid = ex("INSERT INTO tasks(title,problem,objectives,modules,tech,expected,submission_format,criteria,deadline,allow_resubmit,created_by,created_at) "
                     "VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                     (f["title"].strip(), f.get("problem"), f.get("objectives"), f.get("modules"), f.get("tech"), f.get("expected"),
                      f.get("submission_format"), f.get("criteria"), f.get("deadline") or None, 1 if f.get("allow_resubmit") else 0, session["email"], now()))
            audit("task_created", f["title"])
            return redirect(url_for("admin_task_view", tid=tid))
    rows = q("SELECT t.*, (SELECT COUNT(*) FROM assignments a WHERE a.task_id=t.id) AS n FROM tasks t ORDER BY id DESC")
    return render_template("admin_tasks.html", rows=rows)


@app.route("/admin/tasks/<int:tid>")
@login_required(*STAFF)
def admin_task_view(tid):
    t = q("SELECT * FROM tasks WHERE id=?", (tid,), one=True) or abort(404)
    cands = q("SELECT c.*, (SELECT 1 FROM assignments a WHERE a.task_id=? AND a.candidate_id=c.id) AS has FROM candidates c WHERE c.active=1 ORDER BY c.name", (tid,))
    return render_template("admin_task.html", t=t, cands=cands, mods=parse_modules(t["modules"]))


@app.route("/admin/tasks/<int:tid>/assign", methods=["POST"])
@login_required(*STAFF)
def admin_task_assign(tid):
    ids = [int(x) for x in request.form.getlist("cids") if x.isdigit()]
    n = assign_task(tid, ids)
    flash(f"Task assigned to {n} candidate(s). Notifications and AI explanations are being prepared.", "success")
    return redirect(url_for("admin_task_view", tid=tid))


@app.route("/admin/submissions")
@login_required(*STAFF)
def admin_submissions():
    rows = q("SELECT s.*, c.name, c.id AS cid, t.title, r.completion, (SELECT decision FROM review_feedback f WHERE f.submission_id=s.id ORDER BY f.id DESC LIMIT 1) AS decision "
             "FROM submissions s JOIN assignments a ON a.id=s.assignment_id JOIN candidates c ON c.id=a.candidate_id JOIN tasks t ON t.id=a.task_id "
             "LEFT JOIN ai_reports r ON r.submission_id=s.id ORDER BY s.id DESC")
    return render_template("admin_submissions.html", rows=rows)


def load_submission(sid):
    return q("SELECT s.*, a.candidate_id, a.task_id, c.name, c.email, t.title FROM submissions s JOIN assignments a ON a.id=s.assignment_id "
             "JOIN candidates c ON c.id=a.candidate_id JOIN tasks t ON t.id=a.task_id WHERE s.id=?", (sid,), one=True) or abort(404)


@app.route("/admin/submissions/<int:sid>")
@login_required(*STAFF)
def admin_submission(sid):
    s = load_submission(sid)
    r = q("SELECT * FROM ai_reports WHERE submission_id=?", (sid,), one=True)
    return render_template("admin_submission.html", s=s, r=r, rep=json.loads(r["data"]) if r else None,
                           files=q("SELECT * FROM submission_files WHERE submission_id=?", (sid,)),
                           fb=q("SELECT * FROM review_feedback WHERE submission_id=? ORDER BY id DESC", (sid,)))


@app.route("/admin/submissions/<int:sid>/decision", methods=["POST"])
@login_required(*STAFF)
def admin_decision(sid):
    load_submission(sid)
    d = request.form.get("decision")
    if d not in ("approved", "needs_improvement", "pending"):
        abort(400)
    adj = request.form.get("adjusted_score", "")
    ex("INSERT INTO review_feedback(submission_id,reviewer,decision,comments,adjusted_score,created_at) VALUES(?,?,?,?,?,?)",
       (sid, session["email"], d, request.form.get("comments", "")[:3000], float(adj) if re.fullmatch(r"\d+(\.\d+)?", adj) else None, now()))
    audit("review_decision", f"submission {sid}: {d}")
    flash("Reviewer decision saved. (AI scores only assist you - this decision is yours.)", "success")
    return redirect(url_for("admin_submission", sid=sid))


@app.route("/admin/submissions/<int:sid>/reanalyse", methods=["POST"])
@login_required(*STAFF)
def admin_reanalyse(sid):
    load_submission(sid)
    ex("UPDATE submissions SET status='queued' WHERE id=?", (sid,))
    job(AI_POOL, analyse_submission, sid)
    flash("Re-analysis started.", "info")
    return redirect(url_for("admin_submission", sid=sid))


@app.route("/admin/reviews")
@login_required(*STAFF)
def admin_reviews():
    rows = q("SELECT r.*, c.name, c.email FROM review_schedules r JOIN candidates c ON c.id=r.candidate_id ORDER BY r.review_date DESC, r.review_time")
    return render_template("admin_reviews.html", rows=rows)


@app.route("/admin/reviews/remind", methods=["POST"])
@login_required(*STAFF)
def admin_remind():
    n = send_due_reminders()
    flash(f"{n} reminder(s) queued (reviews in the next 24 hours; duplicates are skipped).", "success")
    return redirect(url_for("admin_reviews"))


@app.route("/admin/reviews/<int:rid>/remind", methods=["POST"])
@login_required(*STAFF)
def admin_remind_one(rid):
    rv = q("SELECT * FROM review_schedules WHERE id=?", (rid,), one=True) or abort(404)
    flash("Reminder queued." if send_reminder(rv) else "A reminder was already sent today.", "info")
    return redirect(url_for("admin_reviews"))


def send_due_reminders():
    lim = datetime.now() + timedelta(hours=24)
    n = 0
    for rv in q("SELECT * FROM review_schedules WHERE review_date>=?", (date.today().isoformat(),)):
        try:
            when = datetime.strptime(f"{rv['review_date']} {rv['review_time']}", "%Y-%m-%d %H:%M")
        except ValueError:
            continue
        if datetime.now() <= when <= lim and send_reminder(rv):
            n += 1
    return n


def report_rows(args):
    sql = ("SELECT c.code, c.name, c.email, t.title AS task, a.id AS aid, s.version, s.status AS submission_status, s.submitted_at, c.stage, r.completion, r.data, "
           "(SELECT decision FROM review_feedback f WHERE f.submission_id=s.id ORDER BY f.id DESC LIMIT 1) AS decision, "
           "(SELECT comments FROM review_feedback f WHERE f.submission_id=s.id ORDER BY f.id DESC LIMIT 1) AS comments, s.id AS sid, c.id AS cid, t.id AS tid "
           "FROM submissions s JOIN assignments a ON a.id=s.assignment_id JOIN candidates c ON c.id=a.candidate_id JOIN tasks t ON t.id=a.task_id "
           "LEFT JOIN ai_reports r ON r.submission_id=s.id WHERE 1=1")
    p = []
    for k, col in (("candidate", "c.id"), ("task", "t.id"), ("stage", "c.stage")):
        if args.get(k, "").isdigit():
            sql += f" AND {col}=?"
            p.append(int(args[k]))
    if args.get("from"):
        sql += " AND date(s.submitted_at)>=?"; p.append(args["from"])
    if args.get("to"):
        sql += " AND date(s.submitted_at)<=?"; p.append(args["to"])
    return q(sql + " ORDER BY s.id DESC", p)


def summarize(row):
    if not row["data"]:
        return "", ""
    rep = json.loads(row["data"])
    return (rep.get("message", "")[:300],
            ", ".join(r["name"] for r in rep["requirements"] if r["status"] in ("missing", "unverified", "partial")))


@app.route("/admin/reports")
@login_required(*STAFF)
def admin_reports():
    rows = report_rows(request.args)
    return render_template("admin_reports.html", rows=rows, cands=q("SELECT id,name FROM candidates ORDER BY name"),
                           tasks=q("SELECT id,title FROM tasks"), f=request.args, qs=request.query_string.decode())


@app.route("/admin/export/<kind>")
@login_required(*STAFF)
def admin_export(kind):
    if kind == "candidates":
        df = pd.DataFrame([dict(r) for r in q("SELECT code AS 'Candidate ID', name AS 'Candidate Name', email AS 'Email Address', track AS 'Assigned Task or Track', "
                                              "stage AS 'Current Review Stage', deadline AS 'Task Deadline', status AS 'Candidate Status', active AS Active FROM candidates")])
    elif kind == "reports":
        data = []
        for r in report_rows(request.args):
            summ, miss = summarize(r)
            data.append({"Candidate ID": r["code"], "Candidate": r["name"], "Task": r["task"], "Review stage": r["stage"], "Version": r["version"],
                         "Submission status": r["submission_status"], "Submitted at": r["submitted_at"], "AI completion %": r["completion"],
                         "AI summary": summ, "Missing / unverified requirements": miss, "Reviewer comments": r["comments"], "Final review status": r["decision"] or "pending"})
        df = pd.DataFrame(data)
    else:
        abort(404)
    bio = io.BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)
    return send_file(bio, as_attachment=True, download_name=f"{kind}_{date.today()}.xlsx")


@app.route("/admin/emails")
@login_required(*STAFF)
def admin_emails():
    return render_template("admin_emails.html", rows=q("SELECT * FROM emails ORDER BY id DESC LIMIT 200"), smtp=bool(env("SMTP_HOST")))


@app.route("/admin/audit")
@login_required("admin")
def admin_audit():
    return render_template("admin_audit.html", rows=q("SELECT * FROM audit_log ORDER BY id DESC LIMIT 300"))


# --------------------------------------------------------------------------- candidate side
def my_assignment(aid):
    c = my_candidate() or abort(403)
    a = q("SELECT a.*, t.title, t.problem, t.objectives, t.modules, t.tech, t.expected, t.submission_format, t.criteria, t.deadline, t.allow_resubmit "
          "FROM assignments a JOIN tasks t ON t.id=a.task_id WHERE a.id=? AND a.candidate_id=?", (aid, c["id"]), one=True)
    return (c, a) if a else abort(404)           # another candidate's task looks like "not found"


@app.route("/dashboard")
@login_required("candidate")
def dashboard():
    c = my_candidate() or abort(403)
    asg = q("SELECT a.id, t.title, t.deadline, (SELECT status FROM submissions s WHERE s.assignment_id=a.id ORDER BY version DESC LIMIT 1) AS sub_status, "
            "(SELECT COUNT(*) FROM submissions s WHERE s.assignment_id=a.id) AS versions FROM assignments a JOIN tasks t ON t.id=a.task_id WHERE a.candidate_id=? ORDER BY a.id DESC", (c["id"],))
    revs = q("SELECT * FROM review_schedules WHERE candidate_id=? ORDER BY stage", (c["id"],))
    fb = q("SELECT s.id, s.version, t.title, r.completion, (SELECT decision FROM review_feedback f WHERE f.submission_id=s.id ORDER BY f.id DESC LIMIT 1) AS decision "
           "FROM submissions s JOIN assignments a ON a.id=s.assignment_id JOIN tasks t ON t.id=a.task_id LEFT JOIN ai_reports r ON r.submission_id=s.id "
           "WHERE a.candidate_id=? ORDER BY s.id DESC", (c["id"],))
    return render_template("dashboard.html", c=c, asg=asg, revs=revs, fb=fb)


@app.route("/task/<int:aid>")
@login_required("candidate")
def task_view(aid):
    c, a = my_assignment(aid)
    if not a["explanation"]:
        job(AI_POOL, gen_explanation, aid)
        a = my_assignment(aid)[1]
    subs = q("SELECT s.*, r.completion FROM submissions s LEFT JOIN ai_reports r ON r.submission_id=s.id WHERE assignment_id=? ORDER BY version DESC", (aid,))
    can_submit = (not subs) or a["allow_resubmit"]
    chats = q("SELECT * FROM chats WHERE assignment_id=? ORDER BY id", (aid,))
    return render_template("task_view.html", a=a, subs=subs, can_submit=can_submit, chats=chats, mods=parse_modules(a["modules"]),
                           allowed=", ".join(sorted(ALLOWED_UPLOAD)))


@app.route("/task/<int:aid>/ask", methods=["POST"])
@login_required("candidate")
def ask(aid):
    c, a = my_assignment(aid)
    msg = (request.get_json(silent=True) or {}).get("q", "").strip()[:1000]
    if not msg:
        return jsonify(error="Please type a question."), 400
    if rate_limited(f"ask:{g.user['id']}", 10, 60):
        return jsonify(error="Please slow down a little and try again in a minute."), 429
    hist = q("SELECT role, content FROM chats WHERE assignment_id=? ORDER BY id DESC LIMIT 8", (aid,))[::-1]
    system = (TONE + "\nYou are answering follow-up questions about ONE assigned task only. Stay on this task. Explain, clarify and guide, but do not write the "
              "complete solution for the candidate. If the question is unrelated to the task, politely steer back. If the answer is not covered by the "
              "task description, say so and suggest asking the reviewer.\n\nTASK:\n" + task_text(a))
    reply = llm(system, [{"role": r["role"], "content": r["content"]} for r in hist] + [{"role": "user", "content": msg}], 900)
    if not reply:
        reply = "I can't answer questions right now because the AI assistant isn't connected on this server. Please ask your reviewer, or try again later."
    ex("INSERT INTO chats(assignment_id,role,content,created_at) VALUES(?,?,?,?)", (aid, "user", msg, now()))
    ex("INSERT INTO chats(assignment_id,role,content,created_at) VALUES(?,?,?,?)", (aid, "assistant", reply, now()))
    return jsonify(reply=reply)


@app.route("/task/<int:aid>/submit", methods=["POST"])
@login_required("candidate")
def task_submit(aid):
    c, a = my_assignment(aid)
    prior = q("SELECT MAX(version) FROM submissions WHERE assignment_id=?", (aid,), one=True)[0] or 0
    if prior and not a["allow_resubmit"]:
        flash("Resubmission is not enabled for this task.", "danger")
        return redirect(url_for("task_view", aid=aid))
    gh = request.form.get("github_url", "").strip()
    files = [f for f in request.files.getlist("files") if f and f.filename]
    if gh and not re.match(r"^https://github\.com/[\w.-]+/[\w.-]+", gh):
        flash("The GitHub link should look like https://github.com/username/repository", "danger")
        return redirect(url_for("task_view", aid=aid))
    if not gh and not files:
        flash("Please attach at least one file or paste a GitHub link.", "danger")
        return redirect(url_for("task_view", aid=aid))
    for f in files:
        if Path(f.filename).suffix.lower() not in ALLOWED_UPLOAD:
            flash(f"'{f.filename}' is not an allowed file type.", "danger")
            return redirect(url_for("task_view", aid=aid))
    late = 1 if a["deadline"] and date.today().isoformat() > a["deadline"] else 0
    sid = ex("INSERT INTO submissions(assignment_id,version,github_url,notes,submitted_at,late,status) VALUES(?,?,?,?,?,?,?)",
             (aid, prior + 1, gh or None, request.form.get("notes", "")[:2000], now(), late, "queued"))
    for f in files:
        stored = f"{uuid.uuid4().hex}_{secure_filename(f.filename) or 'file'}"
        f.save(UPLOAD_DIR / stored)
        ex("INSERT INTO submission_files(submission_id,orig_name,stored_name,size) VALUES(?,?,?,?)",
           (sid, secure_filename(f.filename), stored, (UPLOAD_DIR / stored).stat().st_size))
    audit("submission", f"submission {sid} v{prior + 1}")
    job(AI_POOL, analyse_submission, sid)
    flash("Submitted! Your work is being analysed now - this page will update by itself.", "success")
    return redirect(url_for("submission_view", sid=sid))


def my_submission(sid):
    c = my_candidate() or abort(403)
    s = q("SELECT s.*, t.title, a.candidate_id FROM submissions s JOIN assignments a ON a.id=s.assignment_id JOIN tasks t ON t.id=a.task_id "
          "WHERE s.id=? AND a.candidate_id=?", (sid, c["id"]), one=True)
    return s or abort(404)


@app.route("/submission/<int:sid>")
@login_required("candidate")
def submission_view(sid):
    s = my_submission(sid)
    r = q("SELECT * FROM ai_reports WHERE submission_id=?", (sid,), one=True)
    fb = q("SELECT * FROM review_feedback WHERE submission_id=? AND decision!='pending' ORDER BY id DESC LIMIT 1", (sid,), one=True)
    return render_template("submission_view.html", s=s, r=r, rep=json.loads(r["data"]) if r else None, fb=fb)


@app.route("/submission/<int:sid>/detail", methods=["POST"])
@login_required("candidate")
def detail(sid):
    s = my_submission(sid)
    r = q("SELECT * FROM ai_reports WHERE submission_id=?", (sid,), one=True) or abort(404)
    rep = json.loads(r["data"])
    if not rep.get("detailed"):
        txt = llm(TONE, [{"role": "user", "content": "The candidate asked for a more detailed technical explanation of their feedback. For each issue and missing module "
                          "below, explain the technical reason, show a short example of the fix, and say how to test it. Keep the friendly mentor tone. Do not invent new "
                          "problems.\n\n" + json.dumps({"requirements": rep["requirements"], "issues": rep["issues"]})[:8000]}], 3000)
        rep["detailed"] = txt or "Detailed explanations need the AI assistant, which isn't connected on this server. Please ask your reviewer in the review call."
        ex("UPDATE ai_reports SET data=? WHERE submission_id=?", (json.dumps(rep), sid))
    return redirect(url_for("submission_view", sid=sid) + "#detailed")


# --------------------------------------------------------------------------- templates
TEMPLATES = {}
TEMPLATES["base.html"] = r"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{% block title %}Review Portal{% endblock %}</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f5f7fb}.navbar{background:#1f2a44}.card{border:0;box-shadow:0 1px 4px rgba(0,0,0,.08)}.pre{white-space:pre-wrap}
.stat{font-size:2rem;font-weight:700}</style></head><body>
<nav class="navbar navbar-expand-lg navbar-dark mb-4"><div class="container">
<a class="navbar-brand" href="/">Review Portal</a>
{% if me %}<ul class="navbar-nav me-auto">
{% if me.role in ('admin','reviewer') %}
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_dash') }}">Dashboard</a></li>
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_candidates') }}">Candidates</a></li>
{% if me.role=='admin' %}<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_import') }}">Import</a></li>{% endif %}
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_tasks') }}">Tasks</a></li>
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_submissions') }}">Submissions</a></li>
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_reviews') }}">Reviews</a></li>
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_reports') }}">Reports</a></li>
<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_emails') }}">Outbox</a></li>
{% if me.role=='admin' %}<li class="nav-item"><a class="nav-link" href="{{ url_for('admin_audit') }}">Audit</a></li>{% endif %}
{% else %}<li class="nav-item"><a class="nav-link" href="{{ url_for('dashboard') }}">My dashboard</a></li>{% endif %}
</ul><form method="post" action="{{ url_for('logout') }}"><input type="hidden" name="csrf" value="{{ csrf() }}">
<span class="text-white-50 me-2 small">{{ me.name }}</span><button class="btn btn-sm btn-outline-light">Log out</button></form>{% endif %}
</div></nav>
<div class="container pb-5">
{% for cat, m in get_flashed_messages(with_categories=true) %}<div class="alert alert-{{ cat }}">{{ m }}</div>{% endfor %}
{% block body %}{% endblock %}</div></body></html>"""

TEMPLATES["macros.html"] = r"""{% macro badge(s) %}{% set m={'done':('success','Completed'),'partial':('warning','Partially done'),'missing':('danger','Missing'),
'unverified':('secondary','Needs manual check'),'approved':('success','Approved'),'needs_improvement':('warning','Needs improvement'),
'pending':('secondary','Pending'),'done_':('success','Done')} %}{% set x=m.get(s,('secondary',s or '-')) %}<span class="badge bg-{{x[0]}}">{{x[1]}}</span>{% endmacro %}
{% macro report(rep, completion, source) %}
<div class="card mb-3"><div class="card-body"><h5>Overall completion: {{ completion }}%
{% if rep.previous_completion is not none and rep.previous_completion is defined %}<small class="text-muted">(was {{ rep.previous_completion }}%)</small>{% endif %}</h5>
<div class="progress mb-2"><div class="progress-bar" style="width:{{ completion }}%"></div></div>
<small class="text-muted">Calculated from the task's listed modules (partial = {{ 50 }}% credit by default). It is a guide for your reviewer, not a final grade.
Files were read, not run{% if source and 'ai' not in source %}; this report used a basic keyword check because AI analysis was unavailable{% endif %}.</small></div></div>
<div class="card mb-3"><div class="card-body"><h5>Mentor feedback</h5><div class="pre">{{ rep.message }}</div></div></div>
<div class="card mb-3"><div class="card-body"><h5>Requirements checklist</h5><table class="table table-sm align-middle mb-0"><thead><tr><th>Requirement</th><th>Status</th><th>What I saw</th></tr></thead>
{% for r in rep.requirements %}<tr><td>{{ r.name }}</td><td>{{ badge(r.status) }}{% if r.confidence!='confirmed' and r.status!='unverified' %} <small class="text-muted">to verify</small>{% endif %}</td><td class="small">{{ r.evidence }}</td></tr>{% endfor %}</table></div></div>
{% if rep.issues %}<div class="card mb-3"><div class="card-body"><h5>Things to improve</h5>
{% for i in rep.issues %}<div class="mb-3"><strong>{{ i.title }}</strong> <span class="badge bg-{{ 'danger' if i.severity=='high' else 'warning' if i.severity=='medium' else 'info' }}">{{ i.severity }}</span>
{% if i.confirmed %}<span class="badge bg-dark">confirmed</span>{% else %}<span class="badge bg-secondary">possible - needs checking</span>{% endif %}
{% if i.location %}<div class="small text-muted">{{ i.location }}</div>{% endif %}
{% if i.why_it_matters %}<div><em>Why it matters:</em> {{ i.why_it_matters }}</div>{% endif %}{% if i.how_to_fix %}<div><em>How to fix:</em> {{ i.how_to_fix }}</div>{% endif %}
{% if i.snippet %}<pre class="bg-light p-2 small mt-1">{{ i.snippet }}</pre>{% endif %}</div>{% endfor %}</div></div>{% endif %}
{% if rep.documentation_gaps %}<div class="card mb-3"><div class="card-body"><h5>Documentation</h5><ul class="mb-0">{% for d in rep.documentation_gaps %}<li>{{ d }}</li>{% endfor %}</ul></div></div>{% endif %}
{% if rep.next_steps %}<div class="card mb-3"><div class="card-body"><h5>Your next steps</h5><ol class="mb-0">{% for d in rep.next_steps %}<li>{{ d }}</li>{% endfor %}</ol></div></div>{% endif %}
{% if rep.comparison %}<div class="card mb-3"><div class="card-body"><h5>Compared with your previous version</h5>
<p class="mb-1"><strong>Improved requirements:</strong> {{ rep.comparison.requirements_improved|join(', ') or 'none yet' }}</p>
<p class="mb-1"><strong>Issues fixed:</strong> {{ rep.comparison.issues_fixed|join(', ') or 'none' }}</p>
<p class="mb-1"><strong>Still remaining:</strong> {{ rep.comparison.issues_remaining|join(', ') or 'none' }}</p>
<p class="mb-0"><strong>New issues:</strong> {{ rep.comparison.issues_new|join(', ') or 'none' }}</p></div></div>{% endif %}
{% if rep.notes %}<div class="alert alert-secondary small"><strong>Notes:</strong> {{ rep.notes|join(' ') }}</div>{% endif %}
{% endmacro %}"""

TEMPLATES["login.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="row justify-content-center"><div class="col-md-5"><div class="card"><div class="card-body p-4">
<h4>{{ 'Reviewer / Admin login' if kind=='admin' else 'Candidate login' }}</h4>
<form method="post"><input type="hidden" name="csrf" value="{{ csrf() }}">
<div class="mb-3"><label class="form-label">Email</label><input class="form-control" type="email" name="email" required autofocus></div>
<div class="mb-3"><label class="form-label">Password</label><input class="form-control" type="password" name="password" required></div>
<button class="btn btn-primary w-100">Log in</button></form>
<div class="d-flex justify-content-between mt-3 small"><a href="{{ url_for('forgot') }}">Forgot password?</a>
{% if kind=='admin' %}<a href="{{ url_for('login') }}">Candidate login</a>{% else %}<a href="{{ url_for('admin_login') }}">Reviewer login</a>{% endif %}</div>
</div></div></div></div>{% endblock %}"""

TEMPLATES["forgot.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="row justify-content-center"><div class="col-md-5"><div class="card"><div class="card-body p-4">
<h4>Forgot your password?</h4><p class="text-muted">Enter your registered email and we'll send you a reset link.</p>
<form method="post"><input type="hidden" name="csrf" value="{{ csrf() }}"><div class="mb-3"><input class="form-control" type="email" name="email" required></div>
<button class="btn btn-primary w-100">Send reset link</button></form><a class="small d-block mt-3" href="{{ url_for('login') }}">Back to login</a></div></div></div></div>{% endblock %}"""

TEMPLATES["reset.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="row justify-content-center"><div class="col-md-5"><div class="card"><div class="card-body p-4">
<h4>Set your password</h4><form method="post"><input type="hidden" name="csrf" value="{{ csrf() }}">
<div class="mb-3"><label class="form-label">New password</label><input class="form-control" type="password" name="password" minlength="8" required></div>
<div class="mb-3"><label class="form-label">Confirm password</label><input class="form-control" type="password" name="password2" minlength="8" required></div>
<p class="small text-muted">At least 8 characters, with letters and numbers.</p><button class="btn btn-primary w-100">Save password</button></form></div></div></div></div>{% endblock %}"""

TEMPLATES["message.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="card"><div class="card-body"><h4>{{ title }}</h4><p>{{ msg }}</p><a class="btn btn-primary" href="{{ link }}">{{ label }}</a></div></div>{% endblock %}"""

TEMPLATES["admin_dash.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge %}{% block body %}<h3 class="mb-3">Dashboard</h3>
<div class="row g-3 mb-4">{% for lbl,k in [('Candidates','total'),('Tasks assigned','assigned'),('Submitted','submitted'),('Pending','pending'),('Reviews completed','reviewed')] %}
<div class="col"><div class="card"><div class="card-body text-center"><div class="stat">{{ stats[k] }}</div><div class="text-muted">{{ lbl }}</div></div></div></div>{% endfor %}</div>
<div class="row g-3"><div class="col-lg-6"><div class="card"><div class="card-body"><h5>Upcoming reviews</h5><table class="table table-sm"><tr><th>Candidate</th><th>Stage</th><th>When</th></tr>
{% for r in upcoming %}<tr><td>{{ r.name }}</td><td>{{ r.stage }}</td><td>{{ r.review_date }} {{ r.review_time }}</td></tr>{% else %}<tr><td colspan="3" class="text-muted">None scheduled</td></tr>{% endfor %}</table></div></div></div>
<div class="col-lg-6"><div class="card"><div class="card-body"><h5>Recent submissions</h5><table class="table table-sm"><tr><th>Candidate</th><th>Task</th><th>v</th><th>AI %</th></tr>
{% for r in recent %}<tr><td><a href="{{ url_for('admin_submission', sid=r.id) }}">{{ r.name }}</a></td><td>{{ r.title }}</td><td>{{ r.version }}</td><td>{{ r.completion if r.completion is not none else r.status }}</td></tr>
{% else %}<tr><td colspan="4" class="text-muted">No submissions yet</td></tr>{% endfor %}</table></div></div></div></div>{% endblock %}"""

TEMPLATES["admin_candidates.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="d-flex justify-content-between mb-3"><h3>Candidates ({{ rows|length }})</h3>
<div><a class="btn btn-outline-secondary" href="{{ url_for('admin_export', kind='candidates') }}">Export Excel</a> <a class="btn btn-primary" href="{{ url_for('admin_candidate_new') }}">Add candidate</a></div></div>
<form class="row g-2 mb-3"><div class="col-md-4"><input class="form-control" name="q" placeholder="Search name, email or ID" value="{{ f.q }}"></div>
<div class="col-md-3"><select class="form-select" name="track"><option value="">All tracks</option>{% for t in tracks %}<option {{ 'selected' if t==f.track }}>{{ t }}</option>{% endfor %}</select></div>
<div class="col-md-2"><input class="form-control" name="stage" placeholder="Stage" value="{{ f.stage }}"></div><div class="col"><button class="btn btn-dark">Filter</button></div></form>
<div class="card"><div class="table-responsive"><table class="table table-hover mb-0"><thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Track</th><th>Stage</th><th>Status</th><th></th></tr></thead>
{% for c in rows %}<tr class="{{ '' if c.active else 'table-secondary' }}"><td>{{ c.code }}</td><td><a href="{{ url_for('admin_candidate_view', cid=c.id) }}">{{ c.name }}</a></td><td>{{ c.email }}</td>
<td>{{ c.track }}</td><td>{{ c.stage }}</td><td>{{ c.status if c.active else 'Deactivated' }}</td><td><a href="{{ url_for('admin_candidate_edit', cid=c.id) }}">Edit</a></td></tr>
{% else %}<tr><td colspan="7" class="text-muted p-4">No candidates found.</td></tr>{% endfor %}</table></div></div>{% endblock %}"""

TEMPLATES["candidate_form.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="card"><div class="card-body"><h4>{{ 'Edit' if c else 'Add' }} candidate</h4>
<form method="post" class="row g-3"><input type="hidden" name="csrf" value="{{ csrf() }}">
<div class="col-md-3"><label class="form-label">Candidate ID</label><input class="form-control" name="code" value="{{ c.code if c }}"></div>
<div class="col-md-5"><label class="form-label">Name *</label><input class="form-control" name="name" value="{{ c.name if c }}" required></div>
<div class="col-md-4"><label class="form-label">Email *</label><input class="form-control" type="email" name="email" value="{{ c.email if c }}" required></div>
<div class="col-md-4"><label class="form-label">Task / track</label><input class="form-control" name="track" value="{{ c.track if c }}"></div>
<div class="col-md-2"><label class="form-label">Review stage</label><input class="form-control" type="number" min="1" name="stage" value="{{ c.stage if c else 1 }}"></div>
<div class="col-md-3"><label class="form-label">Task deadline</label><input class="form-control" type="date" name="deadline" value="{{ c.deadline if c and c.deadline }}"></div>
<div class="col-md-3"><label class="form-label">Status</label><select class="form-select" name="status">{% for s in ['Active','Inactive','Completed'] %}<option {{ 'selected' if c and c.status==s }}>{{ s }}</option>{% endfor %}</select></div>
<div class="col-12"><button class="btn btn-primary">Save</button> <a class="btn btn-link" href="{{ url_for('admin_candidates') }}">Cancel</a></div></form></div></div>{% endblock %}"""

TEMPLATES["candidate_view.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge %}{% block body %}
<div class="d-flex justify-content-between"><h3>{{ c.name }} <small class="text-muted">{{ c.code }}</small></h3><div class="d-flex gap-2">
<form method="post" action="{{ url_for('admin_resend', cid=c.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}"><button class="btn btn-outline-primary btn-sm">Resend set-password link</button></form>
{% if me.role=='admin' %}<form method="post" action="{{ url_for('admin_candidate_toggle', cid=c.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}"><button class="btn btn-outline-danger btn-sm">{{ 'Deactivate' if c.active else 'Activate' }}</button></form>{% endif %}</div></div>
<p class="text-muted">{{ c.email }} - Track: {{ c.track or '-' }} - Stage {{ c.stage }} - Deadline {{ c.deadline or '-' }} - {{ c.status }}</p>
<div class="row g-3"><div class="col-lg-6"><div class="card"><div class="card-body"><h5>Tasks</h5><ul>{% for a in asg %}<li>{{ a.title }}</li>{% else %}<li class="text-muted">None assigned</li>{% endfor %}</ul>
<h5>Submissions</h5><table class="table table-sm">{% for s in subs %}<tr><td><a href="{{ url_for('admin_submission', sid=s.id) }}">{{ s.title }} v{{ s.version }}</a></td><td>{{ s.status }}</td><td>{{ s.completion if s.completion is not none }}</td></tr>
{% else %}<tr><td class="text-muted">None yet</td></tr>{% endfor %}</table></div></div></div>
<div class="col-lg-6"><div class="card"><div class="card-body"><h5>Review schedule</h5>
{% for r in revs %}<p class="mb-1">Review {{ r.stage }}: {{ r.review_date }} {{ r.review_time }} ({{ r.tz }}) - <a href="{{ r.meet_link }}" target="_blank" rel="noopener">Meet</a></p>{% endfor %}
<hr><form method="post" action="{{ url_for('admin_schedule', cid=c.id) }}" class="row g-2"><input type="hidden" name="csrf" value="{{ csrf() }}">
<div class="col-4"><input class="form-control" type="number" min="1" name="stage" value="{{ c.stage }}" title="Review stage"></div>
<div class="col-4"><input class="form-control" type="date" name="date" required></div><div class="col-4"><input class="form-control" type="time" name="time" required></div>
<div class="col-12"><input class="form-control" name="tz" value="{{ tz }}"></div><div class="col-12"><input class="form-control" name="meet_link" placeholder="https://meet.google.com/..." required></div>
<div class="col-12"><button class="btn btn-primary">Schedule / update &amp; email candidate</button></div></form></div></div></div></div>{% endblock %}"""

TEMPLATES["admin_import.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="card mb-3"><div class="card-body"><h4>Import candidates from Excel</h4>
<p class="text-muted">Required columns: <strong>Candidate ID, Candidate Name, Email Address</strong>. Optional: Assigned Task or Track, Current Review Stage, Task Deadline, Review Date, Review Time, Google Meet Link, Candidate Status.
Existing candidates (same email) are updated, never duplicated. No passwords are stored in Excel - each new candidate receives a set-password link by email.</p>
<form method="post" enctype="multipart/form-data" class="d-flex gap-2"><input type="hidden" name="csrf" value="{{ csrf() }}"><input class="form-control" type="file" name="file" accept=".xlsx,.xls" required>
<button class="btn btn-primary">Import</button></form><a class="small" href="{{ url_for('admin_sample') }}">Download sample Excel file</a></div></div>
{% if result %}<div class="card"><div class="card-body"><h5>Import summary</h5><p>Rows read: {{ result.total }} - <span class="text-success">Created: {{ result.created }}</span> - Updated: {{ result.updated }} - <span class="text-danger">Rejected: {{ result.error_count }}</span></p>
{% if result.error_count %}<a class="btn btn-sm btn-outline-danger mb-2" href="{{ url_for('admin_import_errors', name=result.err_file) }}">Download error report</a>
<table class="table table-sm"><tr><th>Row</th><th>ID</th><th>Email</th><th>Problems</th></tr>{% for e in result.errors %}<tr><td>{{ e['Row'] }}</td><td>{{ e['Candidate ID'] }}</td><td>{{ e['Email'] }}</td><td>{{ e['Problems'] }}</td></tr>{% endfor %}</table>{% endif %}</div></div>{% endif %}{% endblock %}"""

TEMPLATES["admin_tasks.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="row g-3"><div class="col-lg-7"><div class="card"><div class="card-body"><h5>Tasks</h5><table class="table">
<tr><th>Title</th><th>Deadline</th><th>Assigned</th></tr>{% for t in rows %}<tr><td><a href="{{ url_for('admin_task_view', tid=t.id) }}">{{ t.title }}</a></td><td>{{ t.deadline or '-' }}</td><td>{{ t.n }}</td></tr>
{% else %}<tr><td colspan="3" class="text-muted">No tasks yet - create one on the right.</td></tr>{% endfor %}</table></div></div></div>
<div class="col-lg-5"><div class="card"><div class="card-body"><h5>Create task</h5><form method="post"><input type="hidden" name="csrf" value="{{ csrf() }}">
<input class="form-control mb-2" name="title" placeholder="Task title *" required>
<textarea class="form-control mb-2" name="problem" rows="3" placeholder="Problem statement"></textarea><textarea class="form-control mb-2" name="objectives" rows="2" placeholder="Objectives"></textarea>
<textarea class="form-control mb-2" name="modules" rows="4" placeholder="Required modules - one per line. Optional weight: Login validation | 2"></textarea>
<input class="form-control mb-2" name="tech" placeholder="Technology requirements"><input class="form-control mb-2" name="expected" placeholder="Expected output">
<input class="form-control mb-2" name="submission_format" placeholder="Submission format (ZIP / GitHub link / PDF ...)"><textarea class="form-control mb-2" name="criteria" rows="2" placeholder="Evaluation criteria"></textarea>
<input class="form-control mb-2" type="date" name="deadline"><div class="form-check mb-2"><input class="form-check-input" type="checkbox" name="allow_resubmit" id="rs" checked><label class="form-check-label" for="rs">Allow resubmission</label></div>
<button class="btn btn-primary">Create task</button></form></div></div></div></div>{% endblock %}"""

TEMPLATES["admin_task.html"] = r"""{% extends 'base.html' %}{% block body %}<h3>{{ t.title }}</h3><div class="row g-3"><div class="col-lg-6"><div class="card"><div class="card-body">
<p class="pre">{{ t.problem }}</p><h6>Required modules (used for the AI checklist &amp; score)</h6><ul>{% for m in mods %}<li>{{ m.name }} <small class="text-muted">weight {{ m.weight }}</small></li>{% endfor %}</ul>
<p><strong>Tech:</strong> {{ t.tech }}<br><strong>Expected:</strong> {{ t.expected }}<br><strong>Format:</strong> {{ t.submission_format }}<br><strong>Criteria:</strong> {{ t.criteria }}<br><strong>Deadline:</strong> {{ t.deadline or '-' }}</p></div></div></div>
<div class="col-lg-6"><div class="card"><div class="card-body"><h5>Assign to candidates</h5><form method="post" action="{{ url_for('admin_task_assign', tid=t.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}">
<div class="form-check mb-2"><input class="form-check-input" type="checkbox" onclick="document.querySelectorAll('.cb:not(:disabled)').forEach(x=>x.checked=this.checked)" id="all"><label for="all" class="form-check-label">Select all</label></div>
<div style="max-height:360px;overflow:auto">{% for c in cands %}<div class="form-check"><input class="form-check-input cb" type="checkbox" name="cids" value="{{ c.id }}" id="c{{ c.id }}" {{ 'disabled checked' if c.has }}>
<label class="form-check-label" for="c{{ c.id }}">{{ c.name }} <small class="text-muted">{{ c.email }} {{ '(assigned)' if c.has }}</small></label></div>{% endfor %}</div>
<button class="btn btn-primary mt-3">Assign &amp; notify</button></form></div></div></div></div>{% endblock %}"""

TEMPLATES["admin_submissions.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge %}{% block body %}<h3>Submissions</h3><div class="card"><table class="table table-hover mb-0">
<tr><th>Candidate</th><th>Task</th><th>Ver</th><th>Submitted</th><th>Analysis</th><th>AI %</th><th>Decision</th></tr>
{% for r in rows %}<tr><td><a href="{{ url_for('admin_submission', sid=r.id) }}">{{ r.name }}</a></td><td>{{ r.title }}</td><td>{{ r.version }}{{ ' (late)' if r.late }}</td><td>{{ r.submitted_at }}</td><td>{{ r.status }}</td>
<td>{{ r.completion if r.completion is not none }}</td><td>{{ badge(r.decision or 'pending') }}</td></tr>{% else %}<tr><td colspan="7" class="p-4 text-muted">No submissions yet.</td></tr>{% endfor %}</table></div>{% endblock %}"""

TEMPLATES["admin_submission.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge, report %}{% block body %}
<div class="d-flex justify-content-between"><h3>{{ s.name }} - {{ s.title }} (v{{ s.version }})</h3>
<form method="post" action="{{ url_for('admin_reanalyse', sid=s.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}"><button class="btn btn-outline-secondary btn-sm">Re-run AI analysis</button></form></div>
<p class="text-muted">Submitted {{ s.submitted_at }} {{ '(LATE)' if s.late }} - analysis: {{ s.status }} {{ s.error or '' }} {% if s.github_url %}- <a href="{{ s.github_url }}" target="_blank" rel="noopener">GitHub</a>{% endif %}
{% for f in files %}<br>File: {{ f.orig_name }} ({{ f.size }} bytes){% endfor %}{% if s.notes %}<br>Candidate notes: {{ s.notes }}{% endif %}</p>
<div class="row"><div class="col-lg-8">{% if rep %}{{ report(rep, r.completion, r.source) }}{% else %}<div class="alert alert-info">AI report not ready yet.</div>{% endif %}</div>
<div class="col-lg-4"><div class="card mb-3"><div class="card-body"><h5>Reviewer decision</h5><p class="small text-muted">The AI assists; you decide.</p>
<form method="post" action="{{ url_for('admin_decision', sid=s.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}">
<select class="form-select mb-2" name="decision"><option value="pending">Pending</option><option value="needs_improvement">Needs improvement</option><option value="approved">Approved</option></select>
<input class="form-control mb-2" name="adjusted_score" placeholder="Adjusted score % (optional)"><textarea class="form-control mb-2" name="comments" rows="4" placeholder="Comments for the candidate"></textarea>
<button class="btn btn-primary w-100">Save decision</button></form></div></div>
{% for f in fb %}<div class="card mb-2"><div class="card-body small">{{ badge(f.decision) }} by {{ f.reviewer }} on {{ f.created_at }}{% if f.adjusted_score is not none %} - adjusted {{ f.adjusted_score }}%{% endif %}<div class="pre mt-1">{{ f.comments }}</div></div></div>{% endfor %}</div></div>{% endblock %}"""

TEMPLATES["admin_reviews.html"] = r"""{% extends 'base.html' %}{% block body %}<div class="d-flex justify-content-between mb-3"><h3>Review schedule</h3>
<form method="post" action="{{ url_for('admin_remind') }}"><input type="hidden" name="csrf" value="{{ csrf() }}"><button class="btn btn-primary">Send reminders for next 24h</button></form></div>
<div class="card"><table class="table mb-0"><tr><th>Candidate</th><th>Stage</th><th>Date</th><th>Time</th><th>Meet</th><th></th></tr>
{% for r in rows %}<tr><td><a href="{{ url_for('admin_candidate_view', cid=r.candidate_id) }}">{{ r.name }}</a></td><td>{{ r.stage }}</td><td>{{ r.review_date }}</td><td>{{ r.review_time }} {{ r.tz }}</td>
<td><a href="{{ r.meet_link }}" target="_blank" rel="noopener">link</a></td><td><form method="post" action="{{ url_for('admin_remind_one', rid=r.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}"><button class="btn btn-sm btn-outline-secondary">Remind</button></form></td></tr>
{% else %}<tr><td colspan="6" class="p-4 text-muted">Nothing scheduled.</td></tr>{% endfor %}</table></div>{% endblock %}"""

TEMPLATES["admin_reports.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge %}{% block body %}<div class="d-flex justify-content-between mb-3"><h3>Reports</h3>
<a class="btn btn-outline-secondary" href="{{ url_for('admin_export', kind='reports') }}?{{ qs }}">Export to Excel</a></div>
<form class="row g-2 mb-3"><div class="col-md-3"><select class="form-select" name="candidate"><option value="">All candidates</option>{% for c in cands %}<option value="{{ c.id }}" {{ 'selected' if f.candidate==c.id|string }}>{{ c.name }}</option>{% endfor %}</select></div>
<div class="col-md-3"><select class="form-select" name="task"><option value="">All tasks</option>{% for t in tasks %}<option value="{{ t.id }}" {{ 'selected' if f.task==t.id|string }}>{{ t.title }}</option>{% endfor %}</select></div>
<div class="col-md-1"><input class="form-control" name="stage" placeholder="Stage" value="{{ f.stage }}"></div><div class="col-md-2"><input class="form-control" type="date" name="from" value="{{ f['from'] }}"></div>
<div class="col-md-2"><input class="form-control" type="date" name="to" value="{{ f.to }}"></div><div class="col"><button class="btn btn-dark">Filter</button></div></form>
<div class="card"><table class="table mb-0"><tr><th>Candidate</th><th>Task</th><th>Stage</th><th>v</th><th>AI %</th><th>Status</th></tr>
{% for r in rows %}<tr><td>{{ r.name }}</td><td>{{ r.task }}</td><td>{{ r.stage }}</td><td><a href="{{ url_for('admin_submission', sid=r.sid) }}">{{ r.version }}</a></td><td>{{ r.completion if r.completion is not none }}</td><td>{{ badge(r.decision or 'pending') }}</td></tr>
{% else %}<tr><td colspan="6" class="p-4 text-muted">No data for these filters.</td></tr>{% endfor %}</table></div>{% endblock %}"""

TEMPLATES["admin_emails.html"] = r"""{% extends 'base.html' %}{% block body %}<h3>Email outbox</h3>{% if not smtp %}<div class="alert alert-warning">SMTP is not configured, so emails are only recorded here (status "console") and printed in the server window. Set SMTP_HOST etc. to actually send.</div>{% endif %}
<div class="card"><table class="table table-sm mb-0"><tr><th>When</th><th>To</th><th>Subject</th><th>Kind</th><th>Status</th></tr>
{% for e in rows %}<tr><td>{{ e.created_at }}</td><td>{{ e.to_email }}</td><td><details><summary>{{ e.subject }}</summary><pre class="small mb-0">{{ e.body }}</pre></details></td><td>{{ e.kind }}</td><td>{{ e.status }} {{ e.error or '' }}</td></tr>
{% else %}<tr><td colspan="5" class="p-4 text-muted">No emails yet.</td></tr>{% endfor %}</table></div>{% endblock %}"""

TEMPLATES["admin_audit.html"] = r"""{% extends 'base.html' %}{% block body %}<h3>Audit log</h3><div class="card"><table class="table table-sm mb-0"><tr><th>When</th><th>Who</th><th>Action</th><th>Detail</th></tr>
{% for r in rows %}<tr><td>{{ r.created_at }}</td><td>{{ r.actor }}</td><td>{{ r.action }}</td><td>{{ r.detail }}</td></tr>{% endfor %}</table></div>{% endblock %}"""

TEMPLATES["dashboard.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge %}{% block body %}<h3>Hi {{ c.name.split()[0] }} 👋</h3>
<p class="text-muted">Review stage {{ c.stage }} - {{ c.status }}</p><div class="row g-3"><div class="col-lg-7"><div class="card mb-3"><div class="card-body"><h5>My tasks</h5>
<table class="table">{% for a in asg %}<tr><td><a href="{{ url_for('task_view', aid=a.id) }}">{{ a.title }}</a></td><td>{{ a.deadline or '' }}</td>
<td>{% if a.sub_status %}<span class="badge bg-info">{{ a.versions }} submission(s)</span>{% else %}<span class="badge bg-warning text-dark">Not submitted</span>{% endif %}</td></tr>
{% else %}<tr><td class="text-muted">No task assigned yet. You'll get an email when one is.</td></tr>{% endfor %}</table></div></div>
<div class="card"><div class="card-body"><h5>Feedback &amp; previous reports</h5><table class="table">{% for s in fb %}<tr><td><a href="{{ url_for('submission_view', sid=s.id) }}">{{ s.title }} - version {{ s.version }}</a></td>
<td>{{ (s.completion|string + '%') if s.completion is not none else 'analysing...' }}</td><td>{{ badge(s.decision) if s.decision }}</td></tr>{% else %}<tr><td class="text-muted">Nothing yet.</td></tr>{% endfor %}</table></div></div></div>
<div class="col-lg-5"><div class="card"><div class="card-body"><h5>My reviews</h5>{% for r in revs %}<div class="mb-3"><strong>Review {{ r.stage }}</strong><br>{{ r.review_date }} at {{ r.review_time }} ({{ r.tz }})<br>
<a class="btn btn-sm btn-success mt-1" href="{{ r.meet_link }}" target="_blank" rel="noopener">Join Google Meet</a></div>{% else %}<p class="text-muted">No review scheduled yet.</p>{% endfor %}</div></div></div></div>{% endblock %}"""

TEMPLATES["task_view.html"] = r"""{% extends 'base.html' %}{% block body %}<h3>{{ a.title }}</h3><p class="text-muted">Deadline: {{ a.deadline or 'not set' }}</p>
<div class="row g-3"><div class="col-lg-7"><div class="card mb-3"><div class="card-body"><h5>Your task, explained simply</h5>
{% if a.explanation %}<div class="pre">{{ a.explanation }}</div>{% else %}<p class="text-muted">Your mentor is preparing the explanation... <a href="">refresh</a></p>{% endif %}</div></div>
<div class="card mb-3"><div class="card-body"><h5>Original task details</h5><p class="pre">{{ a.problem }}</p><ul>{% for m in mods %}<li>{{ m.name }}</li>{% endfor %}</ul><p class="mb-0"><strong>Submit as:</strong> {{ a.submission_format }}</p></div></div>
<div class="card"><div class="card-body"><h5>Ask a question</h5><div id="chat" class="mb-2" style="max-height:300px;overflow:auto">{% for m in chats %}<div class="mb-2"><strong>{{ 'You' if m.role=='user' else 'Mentor' }}:</strong> <span class="pre">{{ m.content }}</span></div>{% endfor %}</div>
<div class="input-group"><input id="q" class="form-control" maxlength="1000" placeholder="e.g. What does the validation module need to do?"><button class="btn btn-primary" id="send">Ask</button></div></div></div></div>
<div class="col-lg-5"><div class="card mb-3"><div class="card-body"><h5>Submit your work</h5>
{% if can_submit %}<form method="post" action="{{ url_for('task_submit', aid=a.id) }}" enctype="multipart/form-data"><input type="hidden" name="csrf" value="{{ csrf() }}">
<label class="form-label">Files ({{ allowed }})</label><input class="form-control mb-2" type="file" name="files" multiple><input class="form-control mb-2" name="github_url" placeholder="or GitHub link: https://github.com/you/repo">
<textarea class="form-control mb-2" name="notes" rows="2" placeholder="Notes for the reviewer (optional)"></textarea><button class="btn btn-primary w-100">{{ 'Submit new version' if subs else 'Submit' }}</button></form>
<p class="small text-muted mt-2 mb-0">Your code is only read, never run.</p>{% else %}<p class="mb-0">Resubmission isn't enabled for this task.</p>{% endif %}</div></div>
<div class="card"><div class="card-body"><h5>Submission history</h5>{% for s in subs %}<p class="mb-1"><a href="{{ url_for('submission_view', sid=s.id) }}">Version {{ s.version }}</a> - {{ s.submitted_at }}{{ ' (late)' if s.late }} -
{{ (s.completion|string + '%') if s.completion is not none else s.status }}</p>{% else %}<p class="text-muted mb-0">Nothing submitted yet.</p>{% endfor %}</div></div></div></div>
<script>const csrf="{{ csrf() }}",chat=document.getElementById('chat'),qi=document.getElementById('q');
function add(w,t){const d=document.createElement('div');d.className='mb-2';d.innerHTML='<strong></strong> <span class="pre"></span>';d.children[0].textContent=w+':';d.children[1].textContent=t;chat.appendChild(d);chat.scrollTop=chat.scrollHeight}
async function send(){const t=qi.value.trim();if(!t)return;add('You',t);qi.value='';add('Mentor','Thinking...');
const r=await fetch("{{ url_for('ask', aid=a.id) }}",{method:'POST',headers:{'Content-Type':'application/json','X-CSRF-Token':csrf},body:JSON.stringify({q:t})});
const j=await r.json();chat.lastChild.remove();add('Mentor',j.reply||j.error||'Something went wrong.')}
document.getElementById('send').onclick=send;qi.onkeydown=e=>{if(e.key==='Enter')send()}</script>{% endblock %}"""

TEMPLATES["submission_view.html"] = r"""{% extends 'base.html' %}{% from 'macros.html' import badge, report %}{% block body %}
{% if s.status in ('queued','analysing') %}<meta http-equiv="refresh" content="5">{% endif %}
<h3>{{ s.title }} - version {{ s.version }}</h3><p class="text-muted">Submitted {{ s.submitted_at }}{{ ' (after the deadline)' if s.late }}</p>
{% if s.status in ('queued','analysing') %}<div class="alert alert-info">Your submission is being analysed. This page refreshes automatically...</div>
{% elif s.status=='failed' %}<div class="alert alert-danger">The automatic analysis hit a problem. Your reviewer can still see your work and will check it.</div>{% endif %}
{% if fb %}<div class="alert alert-{{ 'success' if fb.decision=='approved' else 'warning' }}"><strong>Reviewer: {{ badge(fb.decision) }}</strong><div class="pre mt-1">{{ fb.comments }}</div></div>{% endif %}
{% if rep %}{{ report(rep, r.completion, r.source) }}
<div id="detailed">{% if rep.detailed %}<div class="card mb-3"><div class="card-body"><h5>Detailed technical explanation</h5><div class="pre">{{ rep.detailed }}</div></div></div>
{% else %}<form method="post" action="{{ url_for('detail', sid=s.id) }}"><input type="hidden" name="csrf" value="{{ csrf() }}"><button class="btn btn-outline-primary">Explain this in more technical detail</button></form>{% endif %}</div>{% endif %}
<a class="btn btn-link" href="{{ url_for('dashboard') }}">Back to dashboard</a>{% endblock %}"""


app.jinja_env.loader = DictLoader(TEMPLATES)


@app.errorhandler(403)
def e403(_):
    return render_template("message.html", title="Not allowed", msg="You don't have access to this page.", link="/", label="Home"), 403


@app.errorhandler(404)
def e404(_):
    return render_template("message.html", title="Not found", msg="We couldn't find that page.", link="/", label="Home"), 404


@app.errorhandler(413)
def e413(_):
    return render_template("message.html", title="File too large", msg=f"Uploads can be at most {MAX_UPLOAD_MB} MB in total.", link="/", label="Home"), 413


@app.errorhandler(500)
def e500(_):
    return render_template("message.html", title="Something went wrong", msg="Sorry, that didn't work. The error has been logged.", link="/", label="Home"), 500


# --------------------------------------------------------------------------- entry point
if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "run"
    init_db()
    if cmd == "sample":
        sample_df().to_excel("sample_candidates.xlsx", index=False)
        print("Wrote sample_candidates.xlsx")
    elif cmd == "reminders":
        with app.app_context():
            print(f"{send_due_reminders()} reminder(s) queued")
        EMAIL_POOL.shutdown(wait=True)
    else:
        if not env("ANTHROPIC_API_KEY"):
            print("NOTE: ANTHROPIC_API_KEY not set -> AI features use a basic offline fallback.")
        app.run(host=env("HOST", "127.0.0.1"), port=int(env("PORT", "5000")), debug=False, threaded=True)
