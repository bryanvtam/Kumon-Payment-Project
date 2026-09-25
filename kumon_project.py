"""
Kumon Project -- Payment Reconciliation Tool
============================================

Standalone desktop tool (Tkinter GUI + local SQLite database) that
reconciles a student roster against several monthly payment files and
produces a multi-tab Excel report. Designed to be packaged into a
single macOS app that runs with no Python install required.

--------------------------------------------------------------------
INTERFACE (Tkinter window titled "Kumon Project")
--------------------------------------------------------------------
Three steps, same as previous iterations:
  Step 1  Select the student database CSV  (First Name, Last Name)
  Step 2  Select the payment files          (one or many, mixed types)
  Step 3  Run Program & Create Output File   -> builds the .xlsx

--------------------------------------------------------------------
DATABASES (all local SQLite tables, rebuilt fresh on every run)
--------------------------------------------------------------------
  master          first_name, last_name, cash, credit, check, prepayment, notes
  paid_not_found  first_name, last_name, cash, credit, check, prepayment, notes
  cash_db         first_name, last_name, payment_amount, notes
  credit_db       first_name, last_name, payment_amount, notes
  check_db        first_name, last_name, payment_amount, notes
  prepayment_db   first_name, last_name, amount_paid, monthly_paid, notes

--------------------------------------------------------------------
FILE ROUTING  (by keyword in the file name, case-insensitive)
--------------------------------------------------------------------
  "cash"        -> cash file        (tab-separated .txt)
  "credit"      -> credit card file (.rtf)
  "deposit"     -> check/EFT file   (.rtf)
  "prepayment"  -> prepayment file  (.csv)
A file whose name matches none of these is reported as an error naming
the file and why, and is skipped -- the rest of the run continues.

--------------------------------------------------------------------
AMOUNT LOGIC (how dollars are computed)
--------------------------------------------------------------------
Check & Credit descriptions state a per-subject price and subject codes:
  M = Math, R = Reading, ELM = Early-Learning Math, ELR = Early-Learning
  Reading;  a slash (M/R, ELR/M, ELM/R ...) means BOTH subjects.
A student's amount = (number of subjects) x (stated per-subject price),
and for the CREDIT file the $5-per-subject service fee is added to the
price first. When no subject code is present, the stated dollar amount
is used as-is. When several students share one transaction, each gets
their own computed amount; as a safety net, if the per-student amounts
don't reconcile to the transaction total, the total is split evenly and
a note is added so the row can be checked by hand.

Only external dependency: openpyxl (for the .xlsx). RTF is parsed by a
small built-in reader (see rtf_to_text) so no Word/LibreOffice/pandoc
is needed on the machine that runs the packaged app.
"""

import csv
import os
import re
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kumon.db")
CREDIT_SERVICE_FEE = 5.0  # per-subject service fee added in the credit file


# =====================================================================
# RTF -> TEXT  (dependency-free; see module docstring for the why)
# =====================================================================

_GROUP_HEADERS_TO_DROP = ("fonttbl", "colortbl", "*\\expandedcolortbl", "*\\generator",
                           "stylesheet", "info", "pict", "object", "*\\listtable",
                           "*\\listoverridetable")
_TOKEN_RE = re.compile(r"\\(?:([a-zA-Z]+)(-?\d+)?(\s)?|(.))", re.DOTALL)


def _strip_special_groups(text):
    out = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == '\\' and i + 1 < n:
            out.append(ch)
            out.append(text[i + 1])
            i += 2
            continue
        if ch == '{':
            depth = 1
            j = i + 1
            while j < n and depth > 0:
                if text[j] == '\\' and j + 1 < n:
                    j += 2
                    continue
                elif text[j] == '{':
                    depth += 1
                elif text[j] == '}':
                    depth -= 1
                j += 1
            inner = text[i + 1:j - 1]
            if any(h in inner[:40] for h in _GROUP_HEADERS_TO_DROP):
                i = j
                continue
            out.append(_strip_special_groups(inner))
            i = j
            continue
        out.append(ch)
        i += 1
    return ''.join(out)


def rtf_to_text(rtf_content):
    text = _strip_special_groups(rtf_content)

    def _hex_sub(m):
        try:
            return bytes([int(m.group(1), 16)]).decode('cp1252', errors='replace')
        except Exception:
            return ''
    text = re.sub(r"\\'([0-9a-fA-F]{2})", _hex_sub, text)

    out, pos = [], 0
    for m in _TOKEN_RE.finditer(text):
        out.append(text[pos:m.start()])
        word, digits, _space, symbol = m.groups()
        if word:
            wl = word.lower()
            if wl in ('par', 'line', 'row'):
                out.append('\n')
            elif wl in ('cell', 'tab'):
                out.append('\t')
            elif wl == 'u':
                try:
                    code = int(digits) if digits else 0
                    if code < 0:
                        code += 65536
                    out.append(chr(code))
                except (ValueError, TypeError):
                    pass
        elif symbol is not None:
            if symbol == '\\':
                out.append('\\')
            elif symbol in '{}':
                out.append(symbol)
            elif symbol == '~':
                out.append(' ')
            elif symbol == '_':
                out.append('-')
            elif symbol in ('\n', '\r'):
                out.append('\n')
        pos = m.end()
    out.append(text[pos:])
    text = ''.join(out)

    text = text.replace('{', '').replace('}', '')
    text = re.sub(r'[ \xa0]+\n', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text


def read_rtf_file(path):
    with open(path, encoding='cp1252', errors='replace') as f:
        return rtf_to_text(f.read())


# =====================================================================
# NAME MATCHING
# =====================================================================

def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip().lower())


def build_master_index(rows):
    """rows: list of (first_name, last_name). Returns lookup records
    used to match free-text / structured names back to the roster."""
    index = []
    for first, last in rows:
        toks = first.split()
        index.append({
            'first_full': norm(first), 'first_tok': norm(toks[0]) if toks else '',
            'last': norm(last), 'first': first, 'last_orig': last,
        })
    return index


def match_pair(first, last, master_index):
    """Match a (first, last) pair against the roster, tolerating a
    middle name present on one side only."""
    rf, rl = norm(first), norm(last)
    if not rf or not rl:
        return None
    for m in master_index:
        if m['last'] != rl:
            continue
        if (rf == m['first_full'] or rf == m['first_tok']
                or m['first_full'].startswith(rf + ' ') or rf.startswith(m['first_full'] + ' ')
                or (m['first_tok'] and m['first_tok'] == rf.split()[0])):
            return (m['first'], m['last_orig'])
    return None


# =====================================================================
# SUBJECT-CODE -> AMOUNT
# =====================================================================

# Ordered so multi-letter codes match before single letters.
_SUBJECT_CODE_RE = re.compile(
    r'\b(ELR/M|ELM/R|EL[MR]|R/M|M/R|MR|RM|M|R|Reading|Math|Rdg)\b', re.IGNORECASE)


def count_subjects(code_text):
    """Number of subjects a code like 'M/R', 'ELR/M', 'MR', 'M'
    represents. A slash, or a two-letter combo like 'MR'/'RM', means two
    subjects; a bare single code means one."""
    t = code_text.upper().replace(' ', '')
    if '/' in t or t in ('MR', 'RM'):
        return 2
    return 1


def compute_amount(description, add_fee=0.0):
    """Estimate a single student's amount from their fragment of a
    payment description, e.g. 'Salma Abdelkarem ELR/M $205 each' or
    'Drake Dela Pena M $195 ELR $205' (two subjects, two prices).

    Returns (amount, note). This is used to WEIGHT how a transaction's
    actual total is split among students; it is not the source of truth
    for the total itself. If no subject code is present, falls back to
    the stated dollar amount(s) summed."""
    # All dollar figures in the fragment.
    prices = [float(p.replace(',', '')) for p in re.findall(r'\$\s*([\d,]+(?:\.\d{2})?)', description)]
    codes = _SUBJECT_CODE_RE.findall(description)

    if codes and len(prices) >= 2 and len(prices) == len(codes):
        # Each subject code has its own explicit price (e.g. "M $195 ELR $205")
        total = sum((p + add_fee) for p in prices)
        note = f"{'/'.join(codes)}: " + " + ".join(f"${p:.2f}" for p in prices)
        if add_fee:
            note += f" (+${add_fee:.0f} fee each)"
        note += f" = ${total:.2f}"
        return total, note

    if codes and prices:
        # One (shared) price; subject count from the code(s).
        n_sub = sum(count_subjects(c) for c in codes)
        price = prices[0]
        amount = (price + add_fee) * n_sub
        note = f"{'/'.join(codes)}: {n_sub} subject(s) x ${price:.2f}"
        if add_fee:
            note = f"{'/'.join(codes)}: {n_sub} subject(s) x (${price:.2f} + ${add_fee:.0f} fee)"
        note += f" = ${amount:.2f}"
        return amount, note

    if prices:
        amount = sum(p + add_fee for p in prices)
        note = f"Stated ${prices[0]:.2f}" + (f" + ${add_fee:.0f} fee" if add_fee else "")
        return amount, note

    # Look for a bare number after a subject word, e.g. "math 185"
    bare = re.search(r'(?i)(?:math|reading|rdg)\s*\$?\s*(\d{2,4})', description)
    if bare:
        amount = float(bare.group(1)) + add_fee
        return amount, f"Stated ${bare.group(1)}"

    return None, "No amount found in description"


# =====================================================================
# STUDENT-NAME EXTRACTION FROM FREE TEXT (credit / check descriptions)
# =====================================================================

_STOPWORDS = {
    'paying', 'for', 'and', 'kumon', 'tuition', 'plus', 'service', 'fee', 'per',
    'subject', 'subjects', 'each', 'math', 'reading', 'rdg', 'early', 'learner',
    'learners', 'the', 'of', 'servide', 'сredit', 'expired', 'cc', 'm', 'r',
    'elr', 'elm', 'el',
}


def _split_student_fragments(description, master_index):
    """Given a payment description covering one or more students, return
    a list of (matched_key_or_None, raw_name, fragment_text) -- one per
    student found. Strategy:

      1. Find every roster student whose last name + first-name-word both
         appear in the text (handles the common 'two first names, one
         shared last name' sibling case, and inserted middle names).
      2. If none match, fall back to a light heuristic that pulls
         capitalised name tokens before the first subject code / '$'.

    Each returned fragment is the slice of text most associated with
    that student, so compute_amount can read the right price."""
    t = norm(description)
    found = []
    for m in master_index:
        if not m['last'] or not re.search(r'\b' + re.escape(m['last']) + r'\b', t):
            continue
        if m['first_tok'] and re.search(r'\b' + re.escape(m['first_tok']) + r'\b', t):
            found.append(m)

    results = []
    if found:
        # Build a fragment per student: the window of text from this
        # student's first-name occurrence up to the next student's, so
        # each student's own price/subject code is read.
        positions = []
        for m in found:
            match = re.search(r'\b' + re.escape(m['first_tok']) + r'\b', t)
            positions.append((match.start() if match else 0, m))
        positions.sort(key=lambda x: x[0])
        for i, (start, m) in enumerate(positions):
            end = positions[i + 1][0] if i + 1 < len(positions) else len(description)
            fragment = description[start:end]
            # If this fragment has no price but the whole description has
            # a single shared price ('$195 each'), use the whole thing.
            if '$' not in fragment:
                fragment = description
            results.append(((m['first'], m['last_orig']), f"{m['first']} {m['last_orig']}", fragment))
        return results

    # No roster match: extract candidate name tokens for the not-found DB.
    # Cut the description at the first subject code, price, or tuition word
    # so trailing "M $195 ..." text isn't captured as part of the name.
    head = re.split(r'(?i)\b(?:ELR/M|ELM/R|EL[MR]|M/R|R/M|Kumon|tuition|Reading|Math|Rdg)\b|\$',
                    description, 1)[0]
    head = re.sub(r'(?i)paying for', '', head)
    tokens = [w for w in re.findall(r"[A-Za-z'\-]+", head) if w.lower() not in _STOPWORDS]
    if tokens:
        first = tokens[0]
        last = tokens[-1] if len(tokens) > 1 else ''
        results.append((None, f"{first} {last}".strip(), description))
    else:
        results.append((None, description[:40], description))
    return results


# =====================================================================
# DATABASE SETUP
# =====================================================================

def get_connection():
    return sqlite3.connect(DB_PATH)


def build_schema(conn):
    """Drop and recreate every table so each run starts clean (no
    double-counting from a previous run)."""
    cur = conn.cursor()
    for tbl in ("master", "paid_not_found", "cash_db", "credit_db", "check_db", "prepayment_db"):
        cur.execute(f"DROP TABLE IF EXISTS {tbl}")

    for tbl in ("master", "paid_not_found"):
        cur.execute(f"""
            CREATE TABLE {tbl} (
                student_id INTEGER PRIMARY KEY AUTOINCREMENT,
                first_name TEXT NOT NULL,
                last_name  TEXT NOT NULL,
                cash       REAL DEFAULT 0,
                credit     REAL DEFAULT 0,
                "check"    REAL DEFAULT 0,
                prepayment REAL DEFAULT 0,
                notes      TEXT DEFAULT '',
                UNIQUE (first_name, last_name)
            )
        """)

    for tbl in ("cash_db", "credit_db", "check_db"):
        cur.execute(f"""
            CREATE TABLE {tbl} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                first_name TEXT,
                last_name  TEXT,
                payment_amount REAL DEFAULT 0,
                notes TEXT DEFAULT ''
            )
        """)

    cur.execute("""
        CREATE TABLE prepayment_db (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT,
            last_name  TEXT,
            amount_paid REAL DEFAULT 0,
            monthly_paid REAL DEFAULT 0,
            notes TEXT DEFAULT ''
        )
    """)
    conn.commit()


def import_roster(conn, csv_path):
    """Read the student roster CSV into the master table."""
    students = []
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        # tolerate slightly different header capitalisation/spacing
        fn_key = next((k for k in reader.fieldnames if norm(k) == 'first name'), reader.fieldnames[0])
        ln_key = next((k for k in reader.fieldnames if norm(k) == 'last name'), reader.fieldnames[1])
        for row in reader:
            first = (row.get(fn_key) or '').strip()
            last = (row.get(ln_key) or '').strip()
            if first and last:
                students.append((first, last))
    conn.executemany(
        "INSERT OR IGNORE INTO master (first_name, last_name) VALUES (?, ?)", students)
    conn.commit()
    return students


# ---- helpers to add money to a student's row in master / not-found ----

def _add_note(existing, new):
    existing = existing or ''
    if not new:
        return existing
    return (existing + " | " + new).strip(" |") if existing else new


def _qcol(column):
    """Quote a column name so SQL reserved words like 'check' are safe."""
    return f'"{column}"'


def add_to_master(conn, key, column, amount, note):
    first, last = key
    qc = _qcol(column)
    row = conn.execute(
        f'SELECT {qc}, notes, cash, credit, "check", prepayment FROM master '
        f'WHERE first_name=? AND last_name=?', (first, last)).fetchone()
    if row is None:
        return
    current, notes = row[0] or 0, row[1] or ''
    # Flag when a prepayment coexists with another payment type (or vice versa)
    existing_types = {'cash': row[2], 'credit': row[3], 'check': row[4], 'prepayment': row[5]}
    flag = ''
    if column == 'prepayment' and any(existing_types[t] for t in ('cash', 'credit', 'check')):
        flag = "FLAG: has prepayment AND another payment type"
    elif column in ('cash', 'credit', 'check') and existing_types['prepayment']:
        flag = "FLAG: has prepayment AND another payment type"
    new_notes = _add_note(notes, note)
    if flag and 'FLAG:' not in new_notes:
        new_notes = _add_note(new_notes, flag)
    conn.execute(
        f'UPDATE master SET {qc}=?, notes=? WHERE first_name=? AND last_name=?',
        (current + amount, new_notes, first, last))
    conn.commit()


def add_to_not_found(conn, first, last, column, amount, note):
    qc = _qcol(column)
    row = conn.execute(
        f'SELECT student_id, {qc}, notes FROM paid_not_found '
        f'WHERE first_name=? AND last_name=?', (first, last)).fetchone()
    if row is None:
        conn.execute(
            f'INSERT INTO paid_not_found (first_name, last_name, {qc}, notes) '
            f'VALUES (?, ?, ?, ?)', (first, last, amount, note))
    else:
        conn.execute(
            f'UPDATE paid_not_found SET {qc}=?, notes=? WHERE student_id=?',
            ((row[1] or 0) + amount, _add_note(row[2], note), row[0]))
    conn.commit()


# =====================================================================
# FILE PARSERS  (each returns nothing; they write straight to the DB)
# =====================================================================

def parse_cash(conn, path, master_index):
    with open(path, encoding='utf-8', errors='replace') as f:
        lines = [l.rstrip('\r\n') for l in f]
    start = 0
    for i, l in enumerate(lines):
        if l.strip().lower().startswith('first name'):
            start = i + 1
            break
    for l in lines[start:]:
        parts = [p.strip() for p in l.split('\t') if p.strip() != '']
        if len(parts) < 3:
            continue
        first, last, amt_str = parts[0], parts[1], parts[2]
        try:
            amount = float(amt_str.replace('$', '').replace(',', ''))
        except ValueError:
            continue
        # Always record into the cash_db (found or not)
        conn.execute("INSERT INTO cash_db (first_name, last_name, payment_amount, notes) VALUES (?,?,?,?)",
                     (first, last, amount, ''))
        key = match_pair(first, last, master_index)
        if key:
            add_to_master(conn, key, 'cash', amount, f"Cash ${amount:.2f}")
        else:
            add_to_not_found(conn, first, last, 'cash', amount, f"Cash ${amount:.2f}")
    conn.commit()


def _parse_text_payment_file(conn, path, master_index, method, db_table, column, add_fee):
    """Shared logic for the credit and check files: both are RTF whose
    rows contain a 'Paying for/For:' description that may cover one or
    more students, plus the ACTUAL transaction amount (the deposited /
    charged figure). Writes to the per-method db_table (every student)
    and to master / paid_not_found.

    Key principle: the transaction's stated amount is the source of
    truth so the per-method grand total always matches the statement.
    Subject-code math is used only to SPLIT that exact amount across
    multiple students in a transaction (e.g. siblings). If the split
    math is unclear, the amount is divided evenly and the row is noted
    for manual review -- but the sum for the transaction always equals
    the stated amount, so no dollars are gained or lost.

    'Not charged' / declined credit rows are recorded to paid_not_found
    (never to the paid totals) with a note that they were not charged."""
    text = read_rtf_file(path)

    rows = []            # (description, txn_amount) for charged/deposited txns
    not_charged = []     # (description, amount, status) for declined credit rows

    if method == 'Check':
        cutoff = text.find('Total Deposits')       # trim duplicated table
        if cutoff != -1:
            text = text[:cutoff]
        for m in re.finditer(r'Paying For:\s*(.*?)\s*\n', text, re.IGNORECASE):
            window = text[m.end():m.end() + 400]
            amt = re.search(r'([\d,]+\.\d{2,4})', window)
            if amt:
                rows.append((m.group(1), float(amt.group(1).replace(',', ''))))
    else:  # Credit
        for chunk in text.split('\n\n'):
            if 'paying for' not in chunk.lower():
                continue
            fields = [f.strip('\t').strip() for f in chunk.split('\t\n') if f.strip(' \t') != '']
            # A real transaction row has 5 fields: name, description, txn id,
            # status, amount. Shorter rows (e.g. the "cards expiring soon"
            # notice at the bottom, which only has name/desc/date) are not
            # payments -- skip them so they don't drop real data.
            if len(fields) < 5:
                continue
            desc = fields[1].split('Paying for', 1)[-1].strip() if 'Paying for' in fields[1] else fields[1]
            status = fields[3].strip().lower()
            try:
                amount = float(fields[4].replace(',', '').replace('$', ''))
            except ValueError:
                continue
            if status == 'approved':
                rows.append((desc, amount))
            else:
                # not charged / declined / failed -> not a payment
                not_charged.append((desc, amount, fields[3].strip()))

    # ---- charged / deposited transactions ----
    for desc, txn_total in rows:
        fragments = _split_student_fragments(desc, master_index)
        n = len(fragments)

        # Weight each student by their computed subject amount, then
        # distribute the ACTUAL txn_total by those weights. This keeps
        # sibling breakdowns meaningful while guaranteeing the parts sum
        # to the real charge.
        weights = []
        for key, raw_name, fragment in fragments:
            amt, _note = compute_amount(fragment, add_fee=add_fee)
            weights.append(amt if (amt and amt > 0) else None)

        known = [w for w in weights if w is not None]
        if n == 1:
            shares = [txn_total]
            split_note = ""
        elif known and abs(sum(known) - txn_total) <= max(1.0, 0.02 * txn_total) and len(known) == n:
            # computed parts reconcile to the charge -> use them as-is
            shares = weights
            split_note = ""
        elif known and len(known) == n and sum(known) > 0:
            # parts don't reconcile but we have a weight for everyone ->
            # scale them proportionally so they sum to the actual total
            scale = txn_total / sum(known)
            shares = [w * scale for w in weights]
            split_note = f"split of ${txn_total:.2f} scaled from subject estimate"
        else:
            # missing weights -> even split, flag for manual review
            shares = [txn_total / n] * n
            split_note = f"CHECK: ${txn_total:.2f} split evenly across {n} student(s)"

        for (key, raw_name, fragment), share in zip(fragments, shares):
            _amt, subj_note = compute_amount(fragment, add_fee=add_fee)
            note = subj_note
            if split_note:
                note = _add_note(note, split_note)
            if key:
                fn, ln = key
            else:
                bits = raw_name.split()
                fn = bits[0] if bits else raw_name
                ln = ' '.join(bits[1:]) if len(bits) > 1 else ''
            conn.execute(
                f"INSERT INTO {db_table} (first_name, last_name, payment_amount, notes) VALUES (?,?,?,?)",
                (fn, ln, share, note))
            if key:
                add_to_master(conn, key, column, share, f"{method}: {note}")
            else:
                add_to_not_found(conn, fn, ln, column, share, f"{method}: {note}")

    # ---- not-charged / declined credit rows -> paid_not_found only ----
    for desc, amount, status in not_charged:
        fragments = _split_student_fragments(desc, master_index)
        for key, raw_name, fragment in fragments:
            if key:
                fn, ln = key
            else:
                bits = raw_name.split()
                fn = bits[0] if bits else raw_name
                ln = ' '.join(bits[1:]) if len(bits) > 1 else ''
            note = f"{method} {status.upper()} (${amount:.2f}) - NOT counted as paid"
            # amount 0 so it never adds to any paid total; note explains why
            add_to_not_found(conn, fn, ln, column, 0.0, note)

    conn.commit()


def parse_credit(conn, path, master_index):
    _parse_text_payment_file(conn, path, master_index, 'Credit', 'credit_db', 'credit', CREDIT_SERVICE_FEE)


def parse_check(conn, path, master_index):
    _parse_text_payment_file(conn, path, master_index, 'Check', 'check_db', 'check', 0.0)


def parse_prepayment(conn, path, master_index):
    with open(path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        cols = {norm(c): c for c in reader.fieldnames}
        fn_key = cols.get('first name')
        ln_key = cols.get('last name')
        exp_key = next((cols[c] for c in cols if 'expir' in c or 'prepaid' in c), None)
        total_key = next((cols[c] for c in cols if 'total paid' in c or c == 'total'), None)
        monthly_key = next((cols[c] for c in cols if 'monthly' in c), None)
        for row in reader:
            first = (row.get(fn_key) or '').strip()
            last = (row.get(ln_key) or '').strip()
            if not first or not last:
                continue
            exp_text = (row.get(exp_key) or '').strip() if exp_key else ''
            total_str = (row.get(total_key) or '').strip() if total_key else ''
            monthly_str = (row.get(monthly_key) or '').strip() if monthly_key else ''

            def _num(s):
                try:
                    return float(s.replace('$', '').replace(',', '').strip())
                except (ValueError, AttributeError):
                    return 0.0
            total_paid = _num(total_str)
            monthly_paid = _num(monthly_str)

            conn.execute(
                "INSERT INTO prepayment_db (first_name, last_name, amount_paid, monthly_paid, notes) "
                "VALUES (?,?,?,?,?)", (first, last, total_paid, monthly_paid, exp_text))

            note = f"Prepaid; {exp_text}" if exp_text else "Prepaid"
            key = match_pair(first, last, master_index)
            if key:
                # master 'prepayment' column holds the MONTHLY amount (the
                # amount attributable to this month); total is kept in the note.
                add_to_master(conn, key, 'prepayment', monthly_paid,
                              f"{note}; total paid ${total_paid:.2f}, monthly ${monthly_paid:.2f}")
            else:
                add_to_not_found(conn, first, last, 'prepayment', monthly_paid,
                                 f"{note}; total paid ${total_paid:.2f}, monthly ${monthly_paid:.2f}")
    conn.commit()


# ---- routing by filename keyword ----

def route_file(path):
    """Return the parser function for a file based on a keyword in its
    name, or raise ValueError explaining why it can't be routed."""
    name = os.path.basename(path).lower()
    if 'cash' in name:
        return parse_cash
    if 'credit' in name:
        return parse_credit
    if 'deposit' in name:
        return parse_check
    if 'prepayment' in name or 'prepaid' in name:
        return parse_prepayment
    raise ValueError(
        f"file name '{os.path.basename(path)}' contains none of the expected keywords "
        f"(cash / credit / deposit / prepayment), so its type can't be determined.")


def run_pipeline(conn, roster_path, payment_paths, log=print):
    """Full Step-3 procedure. Returns a list of {'file','reason'} for any
    files that could not be processed (each skipped, run continues)."""
    build_schema(conn)
    import_roster(conn, roster_path)
    master_rows = conn.execute("SELECT first_name, last_name FROM master").fetchall()
    master_index = build_master_index(master_rows)

    errors = []
    for path in payment_paths:
        fname = os.path.basename(path)
        try:
            parser = route_file(path)
            log(f"Processing {fname} ...")
            parser(conn, path, master_index)
        except ValueError as e:
            log(f"SKIPPED {fname}: {e}")
            errors.append({'file': fname, 'reason': str(e)})
        except Exception as e:
            log(f"SKIPPED {fname}: unexpected error -- {e}")
            errors.append({'file': fname, 'reason': f"unexpected error: {e}"})
    return errors


# =====================================================================
# XLSX EXPORT
# =====================================================================

HEADER_FONT = Font(name="Arial", bold=True)
BODY_FONT = Font(name="Arial")
HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
BAND_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
MONEY_FMT = '$#,##0.00'


def _write_sheet(ws, headers, rows, money_cols=()):
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    # header style
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    # body font + money format + banding
    for ridx in range(2, ws.max_row + 1):
        fill = BAND_FILL if ridx % 2 == 0 else WHITE_FILL
        for cidx in range(1, len(headers) + 1):
            cell = ws.cell(row=ridx, column=cidx)
            cell.font = BODY_FONT
            cell.fill = fill
            if cidx in money_cols:
                cell.number_format = MONEY_FMT
    # column widths
    for cidx in range(1, len(headers) + 1):
        longest = len(str(headers[cidx - 1]))
        for r in rows:
            if cidx - 1 < len(r) and r[cidx - 1] is not None:
                longest = max(longest, len(str(r[cidx - 1])))
        ws.column_dimensions[get_column_letter(cidx)].width = min(max(longest + 2, 10), 60)


def _make_table(ws, name):
    last_col = get_column_letter(ws.max_column)
    ref = f"A1:{last_col}{max(ws.max_row, 1)}"
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False,
                                         showLastColumn=False, showRowStripes=False,
                                         showColumnStripes=False)
    ws.add_table(tbl)


def export_xlsx(conn, out_path):
    wb = Workbook()
    q = conn.execute

    master_headers = ["First Name", "Last Name", "Cash", "Credit", "Check", "Prepayment", "Notes"]
    master_money = (3, 4, 5, 6)

    # Tab 1: full master
    ws = wb.active
    ws.title = "Master"
    rows = q('SELECT first_name, last_name, cash, credit, "check", prepayment, notes '
             'FROM master ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, master_headers, rows, master_money)
    _make_table(ws, "MasterTable")

    # Tab 2: Paid (master students found in any payment file)
    ws = wb.create_sheet("Paid")
    rows = q('SELECT first_name, last_name, cash, credit, "check", prepayment, notes FROM master '
             'WHERE cash>0 OR credit>0 OR "check">0 OR prepayment>0 '
             'ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, master_headers, rows, master_money)
    _make_table(ws, "PaidTable")

    # Tab 3: Paid, not found in master
    ws = wb.create_sheet("Paid, Not Found in Master")
    rows = q('SELECT first_name, last_name, cash, credit, "check", prepayment, notes '
             'FROM paid_not_found ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, master_headers, rows, master_money)
    _make_table(ws, "PaidNotFoundTable")

    # Tab 4: Not paid (in master, no payment anywhere)
    ws = wb.create_sheet("Not Paid")
    rows = q('SELECT first_name, last_name FROM master '
             'WHERE cash=0 AND credit=0 AND "check"=0 AND prepayment=0 '
             'ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, ["First Name", "Last Name"], rows)
    _make_table(ws, "NotPaidTable")

    # Tab 5: Cash DB
    ws = wb.create_sheet("Cash")
    rows = q('SELECT first_name, last_name, payment_amount, notes FROM cash_db '
             'ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, ["First Name", "Last Name", "Payment Amount", "Notes"], rows, (3,))
    _make_table(ws, "CashTable")

    # Tab 6: Check DB
    ws = wb.create_sheet("Check")
    rows = q('SELECT first_name, last_name, payment_amount, notes FROM check_db '
             'ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, ["First Name", "Last Name", "Payment Amount", "Notes"], rows, (3,))
    _make_table(ws, "CheckTable")

    # Tab 7: Credit DB
    ws = wb.create_sheet("Credit")
    rows = q('SELECT first_name, last_name, payment_amount, notes FROM credit_db '
             'ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, ["First Name", "Last Name", "Payment Amount", "Notes"], rows, (3,))
    _make_table(ws, "CreditTable")

    # Tab 8: Prepayment DB
    ws = wb.create_sheet("Prepayment")
    rows = q('SELECT first_name, last_name, amount_paid, monthly_paid, notes FROM prepayment_db '
             'ORDER BY last_name, first_name').fetchall()
    _write_sheet(ws, ["First Name", "Last Name", "Amount Paid", "Monthly Paid", "Notes"], rows, (3, 4))
    _make_table(ws, "PrepaymentTable")

    # Tab 9: Totals
    cash_total = q("SELECT COALESCE(SUM(payment_amount),0) FROM cash_db").fetchone()[0]
    credit_total = q("SELECT COALESCE(SUM(payment_amount),0) FROM credit_db").fetchone()[0]
    check_total = q("SELECT COALESCE(SUM(payment_amount),0) FROM check_db").fetchone()[0]
    prepay_total_paid = q("SELECT COALESCE(SUM(amount_paid),0) FROM prepayment_db").fetchone()[0]
    prepay_monthly = q("SELECT COALESCE(SUM(monthly_paid),0) FROM prepayment_db").fetchone()[0]
    grand_monthly = cash_total + credit_total + check_total + prepay_monthly
    grand_total_paid = cash_total + credit_total + check_total + prepay_total_paid

    ws = wb.create_sheet("Totals")
    total_rows = [
        ["Cash Total", cash_total],
        ["Credit Total", credit_total],
        ["Check Total", check_total],
        ["Prepayment Total Paid", prepay_total_paid],
        ["Prepayment Monthly Total", prepay_monthly],
        ["Grand Total (using prepayment MONTHLY)", grand_monthly],
        ["Grand Total (using prepayment TOTAL PAID)", grand_total_paid],
    ]
    _write_sheet(ws, ["Category", "Total"], total_rows, (2,))
    _make_table(ws, "TotalsTable")

    wb.save(out_path)


# =====================================================================
# GUI  (Tkinter, titled "Kumon Project")
# =====================================================================

def main():
    root = tk.Tk()
    root.title("Kumon Project")
    root.geometry("600x430")
    root.resizable(False, False)

    state = {'roster_path': None, 'payment_paths': []}

    tk.Label(root, text="Kumon Payment Reconciliation", font=("Arial", 16, "bold")).pack(pady=(20, 10))

    # Step 1
    tk.Label(root, text="Step 1: Select the student database CSV", font=("Arial", 11, "bold")).pack(pady=(10, 4))
    roster_label = tk.Label(root, text="No file selected", fg="gray")
    roster_label.pack()

    def select_roster():
        path = filedialog.askopenfilename(title="Select Student Database CSV",
                                           filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        state['roster_path'] = path
        roster_label.config(text=os.path.basename(path), fg="black")
        payments_button.config(state="normal")

    tk.Button(root, text="Browse...", width=20, command=select_roster).pack(pady=4)

    # Step 2
    tk.Label(root, text="Step 2: Select the payment files\n(select multiple together; cash / credit / deposit / prepayment)",
             font=("Arial", 11, "bold"), justify="center").pack(pady=(18, 4))
    payments_label = tk.Label(root, text="No files selected", fg="gray")
    payments_label.pack()

    def select_payment_files():
        paths = filedialog.askopenfilenames(
            title="Select Payment Files",
            filetypes=[("Payment files", "*.csv *.txt *.rtf"), ("All files", "*.*")])
        if not paths:
            return
        state['payment_paths'] = list(paths)
        payments_label.config(text=f"{len(paths)} file(s) selected", fg="black")
        run_button.config(state="normal")

    payments_button = tk.Button(root, text="Browse...", width=20, command=select_payment_files, state="disabled")
    payments_button.pack(pady=4)

    # Step 3
    def run_program():
        if not state['roster_path'] or not state['payment_paths']:
            messagebox.showwarning("Missing Input", "Please select the roster and at least one payment file.")
            return
        try:
            conn = get_connection()
            errors = run_pipeline(conn, state['roster_path'], state['payment_paths'], log=lambda *_: None)
            out_path = filedialog.asksaveasfilename(
                title="Save Output File As", defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")], initialfile="kumon_reconciliation.xlsx")
            if not out_path:
                conn.close()
                return
            export_xlsx(conn, out_path)
            conn.close()
            if errors:
                detail = "\n\n".join(f"{e['file']}:\n  {e['reason']}" for e in errors)
                messagebox.showerror("Some Files Were Skipped",
                                      f"{len(errors)} file(s) could not be processed:\n\n{detail}")
            messagebox.showinfo("Done", "Done")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    run_button = tk.Button(root, text="Step 3: Run Program && Create Output File",
                            command=run_program, state="disabled",
                            bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=35, height=2)
    run_button.pack(pady=24)

    root.mainloop()


if __name__ == "__main__":
    main()
