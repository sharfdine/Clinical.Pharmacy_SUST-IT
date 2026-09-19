"""
LabPilot — shared logic used by app.py and every page.

Keeping this separate from db.py: db.py only knows about SQL, this module
knows about *lab* concepts (what's happening now, what's overdue, what
still needs to be confirmed).
"""

from datetime import datetime, date, timedelta
import pandas as pd
from db import get_conn, DAYS, ROOMS, SLOTS, slot_label, now_iso, today_iso


# ---------- Schedule ----------

def get_schedule_df():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM schedule ORDER BY "
                             "CASE day WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 "
                             "WHEN 'Wednesday' THEN 3 WHEN 'Thursday' THEN 4 ELSE 5 END, "
                             "room, slot_start").fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def get_slot(day, room, slot_start):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM schedule WHERE day=? AND room=? AND slot_start=?",
            (day, room, slot_start),
        ).fetchone()
    return dict(row) if row else None


def upsert_slot(day, room, slot_start, slot_end, course, section, instructor):
    with get_conn() as conn:
        existing = conn.execute(
            "SELECT id FROM schedule WHERE day=? AND room=? AND slot_start=?",
            (day, room, slot_start),
        ).fetchone()
        if existing:
            conn.execute(
                "UPDATE schedule SET course=?, section=?, instructor=?, slot_end=? WHERE id=?",
                (course, section, instructor, slot_end, existing["id"]),
            )
        else:
            conn.execute(
                "INSERT INTO schedule (day, room, slot_start, slot_end, course, section, instructor) "
                "VALUES (?,?,?,?,?,?,?)",
                (day, room, slot_start, slot_end, course, section, instructor),
            )


def course_groups():
    """Every distinct (course, instructor) pairing currently on the schedule,
    with every day/room/time it recurs on. This is what the Experiments page
    lists — one card per *course*, not per calendar slot."""
    df = get_schedule_df()
    df = df[df["course"].notna()]
    groups = {}
    for _, r in df.iterrows():
        key = (r["course"], r["instructor"])
        groups.setdefault(key, {"course": r["course"], "instructor": r["instructor"],
                                 "room": r["room"], "occurrences": []})
        groups[key]["occurrences"].append(
            {"day": r["day"], "room": r["room"], "slot_start": r["slot_start"], "slot_end": r["slot_end"]}
        )
    return groups


# ---------- Current / upcoming ----------

def current_sessions(now=None):
    """What's happening in each room right now, or None if free."""
    now = now or datetime.now()
    day_idx = now.weekday()  # Monday=0
    if day_idx >= 5:
        return {room: None for room in ROOMS}
    day = DAYS[day_idx]
    t = now.strftime("%H:%M")
    result = {}
    with get_conn() as conn:
        for room in ROOMS:
            rows = conn.execute(
                "SELECT * FROM schedule WHERE day=? AND room=? AND course IS NOT NULL",
                (day, room),
            ).fetchall()
            hit = None
            for r in rows:
                if r["slot_start"] <= t < r["slot_end"]:
                    hit = dict(r)
                    break
            result[room] = hit
    return result


def most_recent_date_for_weekday(weekday_idx, ref=None):
    """weekday_idx: Monday=0 ... Friday=4. Returns the most recent date
    (today or earlier) that falls on that weekday."""
    ref = ref or date.today()
    diff = (ref.weekday() - weekday_idx) % 7
    return ref - timedelta(days=diff)


# ---------- Experiments (planning) ----------

def get_experiment(course, instructor):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM experiments WHERE course=? AND instructor=?",
            (course, instructor),
        ).fetchone()
        if not row:
            return None
        exp = dict(row)
        items = conn.execute(
            "SELECT * FROM experiment_items WHERE experiment_id=? ORDER BY id",
            (exp["id"],),
        ).fetchall()
        history = conn.execute(
            "SELECT * FROM experiment_history WHERE experiment_id=? ORDER BY changed_date DESC LIMIT 15",
            (exp["id"],),
        ).fetchall()
        exp["items"] = [dict(i) for i in items]
        exp["history"] = [dict(h) for h in history]
        return exp


def set_experiment_name(course, instructor, room, new_name):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM experiments WHERE course=? AND instructor=?",
            (course, instructor),
        ).fetchone()
        if row:
            if row["experiment_name"] and row["experiment_name"] != new_name:
                conn.execute(
                    "INSERT INTO experiment_history (experiment_id, experiment_name, changed_date) VALUES (?,?,?)",
                    (row["id"], row["experiment_name"], now_iso()),
                )
            conn.execute(
                "UPDATE experiments SET experiment_name=?, set_date=? WHERE id=?",
                (new_name, today_iso(), row["id"]),
            )
        else:
            conn.execute(
                "INSERT INTO experiments (course, instructor, room, experiment_name, set_date) VALUES (?,?,?,?,?)",
                (course, instructor, room, new_name, today_iso()),
            )


def add_experiment_item(course, instructor, room, item_name, item_type, unit, location, item_room, planned_qty):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id FROM experiments WHERE course=? AND instructor=?", (course, instructor)
        ).fetchone()
        if not row:
            conn.execute(
                "INSERT INTO experiments (course, instructor, room, experiment_name, set_date) VALUES (?,?,?,?,?)",
                (course, instructor, room, "", today_iso()),
            )
            exp_id = conn.execute(
                "SELECT id FROM experiments WHERE course=? AND instructor=?", (course, instructor)
            ).fetchone()["id"]
        else:
            exp_id = row["id"]
        conn.execute(
            "INSERT INTO experiment_items (experiment_id, item_name, item_type, unit, location, room, planned_qty) "
            "VALUES (?,?,?,?,?,?,?)",
            (exp_id, item_name, item_type, unit, location, item_room, planned_qty),
        )


def remove_experiment_item(item_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM experiment_items WHERE id=?", (item_id,))


# ---------- Inventory (cross-lab lookup) ----------

def search_inventory(query, item_type=None):
    q = f"%{query.lower()}%"
    with get_conn() as conn:
        if item_type:
            rows = conn.execute(
                "SELECT * FROM inventory WHERE lower(item_name) LIKE ? AND item_type=? ORDER BY item_name LIMIT 30",
                (q, item_type),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM inventory WHERE lower(item_name) LIKE ? ORDER BY item_name LIMIT 30",
                (q,),
            ).fetchall()
    return [dict(r) for r in rows]


def all_inventory_names(item_type=None):
    with get_conn() as conn:
        if item_type:
            rows = conn.execute(
                "SELECT DISTINCT item_name, room, location, unit, item_type FROM inventory "
                "WHERE item_type=? ORDER BY item_name", (item_type,)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT DISTINCT item_name, room, location, unit, item_type FROM inventory ORDER BY item_name"
            ).fetchall()
    return [dict(r) for r in rows]


# ---------- Tracker (what actually happened) ----------

def pending_confirmations(days_back=14):
    """Every schedule slot, in the last `days_back` days up to today, whose
    end time has passed and that has no matching tracker_log entry yet.
    This is what makes end-of-week batch logging possible: nothing is lost
    just because you didn't log it same-day."""
    pending = []
    today = date.today()
    now = datetime.now()
    with get_conn() as conn:
        logged = conn.execute(
            "SELECT log_date, room, slot_start FROM tracker_log"
        ).fetchall()
        logged_keys = {(r["log_date"], r["room"], r["slot_start"]) for r in logged}

        for offset in range(days_back):
            d = today - timedelta(days=offset)
            if d.weekday() >= 5:
                continue
            day_name = DAYS[d.weekday()]
            rows = conn.execute(
                "SELECT * FROM schedule WHERE day=? AND course IS NOT NULL", (day_name,)
            ).fetchall()
            for r in rows:
                slot_end_dt = datetime.combine(d, datetime.strptime(r["slot_end"], "%H:%M").time())
                if slot_end_dt > now:
                    continue  # hasn't finished yet
                key = (d.isoformat(), r["room"], r["slot_start"])
                if key in logged_keys:
                    continue
                exp = get_experiment(r["course"], r["instructor"])
                pending.append({
                    "date": d.isoformat(),
                    "day": day_name,
                    "room": r["room"],
                    "slot_start": r["slot_start"],
                    "slot_end": r["slot_end"],
                    "slot_label": slot_label(r["slot_start"], r["slot_end"]),
                    "course": r["course"],
                    "section": r["section"],
                    "instructor": r["instructor"],
                    "experiment_name": exp["experiment_name"] if exp else "",
                    "planned_items": exp["items"] if exp else [],
                })
    pending.sort(key=lambda p: p["date"], reverse=True)
    return pending


def save_tracker_entry(log_date, day, room, slot_start, slot_end, course, section,
                        instructor, experiment_name, notes, items):
    """items: list of dicts with item_name, item_type, unit, location, qty_used"""
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO tracker_log (log_date, day, room, slot_start, slot_end, course, section, "
            "instructor, experiment_name, notes, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (log_date, day, room, slot_start, slot_end, course, section, instructor,
             experiment_name, notes, now_iso()),
        )
        log_id = cur.lastrowid
        for it in items:
            conn.execute(
                "INSERT INTO tracker_items (tracker_log_id, item_name, item_type, unit, location, qty_used) "
                "VALUES (?,?,?,?,?,?)",
                (log_id, it["item_name"], it.get("item_type", "chemical"),
                 it.get("unit", ""), it.get("location", ""), it.get("qty_used")),
            )
        return log_id


def get_tracker_log(limit=200, search=None, date_from=None, date_to=None):
    query = "SELECT * FROM tracker_log WHERE 1=1"
    params = []
    if search:
        query += " AND (course LIKE ? OR instructor LIKE ? OR experiment_name LIKE ? OR notes LIKE ?)"
        like = f"%{search}%"
        params += [like, like, like, like]
    if date_from:
        query += " AND log_date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND log_date <= ?"
        params.append(date_to)
    query += " ORDER BY log_date DESC LIMIT ?"
    params.append(limit)
    with get_conn() as conn:
        rows = conn.execute(query, params).fetchall()
        entries_out = []
        for r in rows:
            e = dict(r)
            items = conn.execute(
                "SELECT * FROM tracker_items WHERE tracker_log_id=?", (e["id"],)
            ).fetchall()
            e["items"] = [dict(i) for i in items]
            entries_out.append(e)
    return entries_out


def delete_tracker_entry(log_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM tracker_log WHERE id=?", (log_id,))


# ---------- Unified entries: notes, correspondence, follow-ups ----------
#
# One table, three "flavors". A follow-up carries a due_date and a status
# ('open'/'done'); notes and correspondence leave both NULL. This is what
# lets the Notes page show a single color-coded weekly feed instead of
# three separate places to check.

def add_entry(entry_date, entry_type, title, body, due_date=None):
    status = "open" if entry_type == "followup" else None
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO entries (entry_date, entry_type, title, body, due_date, status, "
            "created_at, last_touched) VALUES (?,?,?,?,?,?,?,?)",
            (entry_date, entry_type, title, body, due_date, status, now_iso(), today_iso()),
        )
        return cur.lastrowid


def get_entries(search=None, entry_types=None, date_from=None, date_to=None,
                status=None, limit=1000):
    query = "SELECT * FROM entries WHERE 1=1"
    params = []
    if search:
        query += " AND (title LIKE ? OR body LIKE ?)"
        params += [f"%{search}%", f"%{search}%"]
    if entry_types:
        placeholders = ",".join("?" for _ in entry_types)
        query += f" AND entry_type IN ({placeholders})"
        params += list(entry_types)
    if date_from:
        query += " AND entry_date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND entry_date <= ?"
        params.append(date_to)
    if status:
        query += " AND status = ?"
        params.append(status)
    query += " ORDER BY entry_date DESC, id DESC LIMIT ?"
    params.append(limit)
    with get_conn() as conn:
        rows = conn.execute(query, params).fetchall()
    return [dict(r) for r in rows]


def delete_entry(entry_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM entries WHERE id=?", (entry_id,))


def mark_followup(entry_id, status):
    with get_conn() as conn:
        conn.execute(
            "UPDATE entries SET status=?, last_touched=? WHERE id=?",
            (status, today_iso(), entry_id),
        )


def snooze_followup(entry_id, new_due_date):
    with get_conn() as conn:
        conn.execute(
            "UPDATE entries SET due_date=?, last_touched=? WHERE id=?",
            (new_due_date, today_iso(), entry_id),
        )


def overdue_followups():
    today = today_iso()
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM entries WHERE entry_type='followup' AND status='open' "
            "AND due_date IS NOT NULL AND due_date < ? ORDER BY due_date", (today,),
        ).fetchall()
    return [dict(r) for r in rows]


def open_followups():
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM entries WHERE entry_type='followup' AND status='open' "
            "ORDER BY due_date IS NULL, due_date, id",
        ).fetchall()
    return [dict(r) for r in rows]


def entries_calendar_days(year, month):
    """{iso_date: {'note': n, 'correspondence': n, 'followup': n}} for the month,
    used to color-code the calendar dots."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT entry_date, entry_type, COUNT(*) as c FROM entries "
            "WHERE entry_date LIKE ? GROUP BY entry_date, entry_type",
            (f"{year:04d}-{month:02d}-%",),
        ).fetchall()
    result = {}
    for r in rows:
        result.setdefault(r["entry_date"], {})[r["entry_type"]] = r["c"]
    return result


def week_range(ref=None):
    """Monday..Sunday ISO dates for the week containing `ref` (default today)."""
    ref = ref or date.today()
    monday = ref - timedelta(days=ref.weekday())
    sunday = monday + timedelta(days=6)
    return monday.isoformat(), sunday.isoformat()
