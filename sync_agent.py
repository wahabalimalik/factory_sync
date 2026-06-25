"""
Factory Access DB -> Odoo Sync Agent
=====================================
Reads new rows from a password-protected MS Access (.accdb) table and
pushes them to an Odoo 15 JSON endpoint.

Designed to run as a standalone .exe (built with PyInstaller) on a
factory PC that has no Python installed. Just copy sync_agent.exe and
config.json into the same folder and run the exe.

How new rows are detected
--------------------------
Access has no built-in "row was just inserted" notification that can be
read from outside the application that owns the file. The reliable
approach used here is ID-based polling:

    1. Remember the highest row ID we have already sent (in sync_state.json).
    2. Every `poll_interval_seconds`, query for rows with a higher ID.
    3. Send only the new rows to Odoo.
    4. Only update the remembered ID after Odoo confirms success.

This is crash-safe (if the script restarts, it resumes from the last
confirmed ID), network-safe (failed sends are retried next cycle without
duplicating already-sent rows), and doesn't require modifying whatever
software currently writes into pro.accdb.

Author: Growise Tech
"""

import json
import logging
import logging.handlers
import os
import signal
import sys
import time
from datetime import datetime, date

import pyodbc
import requests

# ---------------------------------------------------------------------------
# Paths (works both as a .py script and as a frozen PyInstaller .exe)
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(
            f"config.json not found next to the executable: {CONFIG_PATH}"
        )
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


CONFIG = load_config()

STATE_PATH = os.path.join(BASE_DIR, CONFIG.get("state_file", "sync_state.json"))
LOG_PATH = os.path.join(BASE_DIR, CONFIG.get("log_file", "sync_agent.log"))


# ---------------------------------------------------------------------------
# Logging - rotates at 5MB, keeps 5 backups, also prints to console if
# you run the .py directly (frozen .exe stays silent / log-file only).
# ---------------------------------------------------------------------------
logger = logging.getLogger("sync_agent")
logger.setLevel(logging.INFO)

file_handler = logging.handlers.RotatingFileHandler(
    LOG_PATH, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
)
file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(file_handler)

if not getattr(sys, "frozen", False):
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(console_handler)


# ---------------------------------------------------------------------------
# State (remembers the last synced row so we never resend old data)
# ---------------------------------------------------------------------------
def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_id": 0}


def save_state(state):
    # Write to a temp file then rename -> atomic, can't end up half-written
    tmp_path = STATE_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(state, f)
    os.replace(tmp_path, STATE_PATH)


# ---------------------------------------------------------------------------
# Access DB connection
# ---------------------------------------------------------------------------
def get_connection():
    db_path = CONFIG["access_db_path"]
    password = CONFIG.get("access_db_password", "")
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
        f"PWD={password};"
    )
    # ReadOnly connection avoids fighting for a lock with whatever
    # software is already writing into pro.accdb.
    return pyodbc.connect(conn_str, autocommit=True, readonly=True)


def json_safe(value):
    """Convert Access/ODBC types that json.dumps can't handle natively."""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def fetch_new_rows(last_id):
    """
    Pull every row whose ID is greater than the last one we already sent.
    Assumes the table has an autonumber / integer key column
    (configurable via "id_column" in config.json - default "ID").
    """
    id_column = CONFIG.get("id_column", "ID")
    table_name = CONFIG["table_name"]

    conn = get_connection()
    try:
        cursor = conn.cursor()
        query = (
            f"SELECT * FROM [{table_name}] "
            f"WHERE [{id_column}] > ? ORDER BY [{id_column}] ASC"
        )
        cursor.execute(query, last_id)
        columns = [col[0] for col in cursor.description]
        rows = []
        for row in cursor.fetchall():
            rows.append({col: json_safe(val) for col, val in zip(columns, row)})
        return rows, id_column
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Push to Odoo
# ---------------------------------------------------------------------------
def push_to_odoo(rows, id_column):
    url = CONFIG["odoo_url"]
    api_key = CONFIG.get("odoo_api_key", "")
    headers = {
        "Content-Type": "application/json",
        "X-API-KEY": api_key,
    }
    # id_column is sent alongside the rows so Odoo knows which field to
    # use for de-duplication, regardless of what it's actually named in
    # your Access table (ID, RowID, RecordNo, whatever it is).
    payload = {"id_column": id_column, "records": rows}
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    return response.json()


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------
RUNNING = True


def handle_stop(signum, frame):
    global RUNNING
    logger.info("Stop signal received - shutting down after current cycle.")
    RUNNING = False


signal.signal(signal.SIGINT, handle_stop)
signal.signal(signal.SIGTERM, handle_stop)


def run_cycle():
    state = load_state()
    last_id = state.get("last_id", 0)

    rows, id_column = fetch_new_rows(last_id)
    if not rows:
        logger.info("No new rows since ID %s.", last_id)
        return

    logger.info("Found %d new row(s). Sending to Odoo...", len(rows))
    try:
        result = push_to_odoo(rows, id_column)
        logger.info("Odoo response: %s", result)
    except Exception:
        logger.exception("Failed to push rows to Odoo - will retry next cycle.")
        return  # do NOT advance last_id - retry the same rows next time

    new_last_id = max(row[id_column] for row in rows)
    save_state({"last_id": new_last_id})
    logger.info("State updated. last_id is now %s.", new_last_id)


def main():
    interval = CONFIG.get("poll_interval_seconds", 60)
    logger.info("Sync agent started. Polling every %s seconds.", interval)
    while RUNNING:
        try:
            run_cycle()
        except Exception:
            logger.exception("Unexpected error in sync cycle.")
        time.sleep(interval)
    logger.info("Sync agent stopped.")


if __name__ == "__main__":
    main()
