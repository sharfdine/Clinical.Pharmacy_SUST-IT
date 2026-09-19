"""
LabPilot — database layer.

Everything persists in a single SQLite file (data/labpilot.db). All pages
import this module rather than touching sqlite3 directly, so the schema
lives in exactly one place.
"""

import sqlite3
import csv
import os
from datetime import datetime, date
from contextlib import contextmanager

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "labpilot.db")
SEED_INVENTORY_CSV = os.path.join(os.path.dirname(__file__), "data", "seed_inventory.csv")

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ROOMS = ["135B", "138B"]

# The two standard lab blocks. Kept as constants so every page agrees on
# what "a slot" means; if your institution's block times ever change,
# update them here once.
SLOTS = [
    ("11:00", "14:00", "11:00 AM – 2:00 PM"),
    ("14:00", "17:00", "2:00 PM – 5:00 PM"),
]

# Three entry types share one table — a "follow-up" is just a note or
# correspondence with a due date and a status attached. This is what lets
# the Notes page show all three together, color-coded, in one feed.
ENTRY_TYPES = ["note", "correspondence", "followup"]

# The current, corrected official FA26 timetable for 135B / 138B, per the
# FCCU registrar export. Only used to populate an empty database on first
# run — after that, the Weekly Schedule page is the source of truth.
SEED_SCHEDULE = [
    # day, room, slot_start, course, section, instructor
    ("Monday", "135B", "11:00", "Pharmaceutics-IA: Phys. Pharmacy-I", "121C", "Mr. M. Razzaq"),
    ("Monday", "135B", "14:00", "Pharmaceutics IIA", "210A", "Dr. Abdul Haleem Khan"),
    ("Monday", "138B", "11:00", "Pharmacognosy-IIA", "313A", "Mr. S. Peerzada"),
    ("Monday", "138B", "14:00", "Pharma. Chem. IIIB", "316A", "Mr. Nasir Ali"),

    ("Tuesday", "135B", "11:00", None, None, None),
    ("Tuesday", "135B", "14:00", "Pharmaceutics IIIA", "211A", "Ms. Saman Ali"),
    ("Tuesday", "138B", "11:00", None, None, None),
    ("Tuesday", "138B", "14:00", "Pharmacognosy", "213B", "Mr. S. Peerzada"),

    ("Wednesday", "135B", "11:00", None, None, None),
    ("Wednesday", "135B", "14:00", "Pharmaceutics IIA", "210B", "Dr. Abdul Haleem Khan"),
    ("Wednesday", "138B", "11:00", "Pharmacognosy-IIB", "318A", "Mr. S. U. Rehman"),
    ("Wednesday", "138B", "14:00", "Pharmacognosy", "213C", "Mr. S. Peerzada"),

    ("Thursday", "135B", "11:00", "PhRM Practice-IIA", "310A", "Mr. Sufyan U."),
    ("Thursday", "135B", "14:00", "Pharmaceutics-IA: Phys. Pharmacy-I", "121B", "Mr. M. Razzaq"),
    ("Thursday", "138B", "11:00", None, None, None),
    ("Thursday", "138B", "14:00", "Pharmaceutics IIIA", "211C", "Ms. Saman Ali"),

    ("Friday", "135B", "11:00", "Pharmacy Practice-VB", "356B", "Mr. Omaid Khan"),
    ("Friday", "135B", "14:00", "Pharmaceutics IIA", "210C", "Dr. Abdul Haleem Khan"),
    ("Friday", "138B", "11:00", "Pharmacognosy-IIB", "318B", "Mr. S. U. Rehman"),
    ("Friday", "138B", "14:00", "Pharmacognosy", "213A", "Mr. S. U. Rehman"),
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS schedule (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    day TEXT NOT NULL,
    room TEXT NOT NULL,
    slot_start TEXT NOT NULL,
    slot_end TEXT NOT NULL,
    course TEXT,
    section TEXT,
    instructor TEXT
);

CREATE TABLE IF NOT EXISTS inventory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_name TEXT NOT NULL,
    item_type TEXT NOT NULL DEFAULT 'chemical',
    unit TEXT,
    location TEXT,
    room TEXT,
    stock_qty REAL,
    UNIQUE(item_name, room, location, item_type)
);

CREATE TABLE IF NOT EXISTS experiments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    course TEXT NOT NULL,
    instructor TEXT NOT NULL,
    room TEXT,
    experiment_name TEXT,
    set_date TEXT,
    UNIQUE(course, instructor)
);

CREATE TABLE IF NOT EXISTS experiment_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    experiment_id INTEGER REFERENCES experiments(id) ON DELETE CASCADE,
    experiment_name TEXT,
    changed_date TEXT
);

CREATE TABLE IF NOT EXISTS experiment_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    experiment_id INTEGER REFERENCES experiments(id) ON DELETE CASCADE,
    item_name TEXT,
    item_type TEXT,
    unit TEXT,
    location TEXT,
    room TEXT,
    planned_qty REAL
);

CREATE TABLE IF NOT EXISTS tracker_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    log_date TEXT NOT NULL,
    day TEXT,
    room TEXT,
    slot_start TEXT,
    slot_end TEXT,
    course TEXT,
    section TEXT,
    instructor TEXT,
    experiment_name TEXT,
    notes TEXT,
    created_at TEXT
);

CREATE TABLE IF NOT EXISTS tracker_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tracker_log_id INTEGER REFERENCES tracker_log(id) ON DELETE CASCADE,
    item_name TEXT,
    item_type TEXT,
    unit TEXT,
    location TEXT,
    qty_used REAL
);

-- Unified table for notes, correspondence, and follow-ups. A follow-up is
-- simply an entry with entry_type='followup' plus a due_date and status;
-- notes/correspondence leave those two columns NULL.
CREATE TABLE IF NOT EXISTS entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,
    entry_type TEXT NOT NULL DEFAULT 'note',
    title TEXT,
    body TEXT,
    due_date TEXT,
    status TEXT,
    created_at TEXT,
    last_touched TEXT
);
"""


@contextmanager
def get_conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    """Create tables if missing, and seed schedule/inventory on first run only."""
    with get_conn() as conn:
        conn.executescript(SCHEMA)

        row = conn.execute("SELECT COUNT(*) AS c FROM schedule").fetchone()
        if row["c"] == 0:
            for day, room, start, course, section, instr in SEED_SCHEDULE:
                end = "14:00" if start == "11:00" else "17:00"
                conn.execute(
                    "INSERT INTO schedule (day, room, slot_start, slot_end, course, section, instructor) "
                    "VALUES (?,?,?,?,?,?,?)",
                    (day, room, start, end, course, section, instr),
                )

        row = conn.execute("SELECT COUNT(*) AS c FROM inventory").fetchone()
        if row["c"] == 0 and os.path.exists(SEED_INVENTORY_CSV):
            with open(SEED_INVENTORY_CSV, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    qty = r["stock_qty"]
                    try:
                        qty = float(qty) if qty not in (None, "") else None
                    except ValueError:
                        qty = None
                    conn.execute(
                        "INSERT OR IGNORE INTO inventory (item_name, item_type, unit, location, room, stock_qty) "
                        "VALUES (?,?,?,?,?,?)",
                        (r["item_name"], r["item_type"], r["unit"], r["location"], r["room"], qty),
                    )


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def today_iso():
    return date.today().isoformat()


def fmt_time(t):
    """'14:00' -> '2:00 PM'"""
    h, m = map(int, t.split(":"))
    suffix = "AM" if h < 12 else "PM"
    h12 = h % 12
    if h12 == 0:
        h12 = 12
    return f"{h12}:{m:02d} {suffix}"


def slot_label(start, end):
    return f"{fmt_time(start)} – {fmt_time(end)}"
