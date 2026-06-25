"""
Factory Access DB -> Odoo Sync Agent
=====================================
Reads new rows from a password-protected MS Access (.accdb) table and
pushes them to an Odoo 15 JSON endpoint.

Designed to run as a standalone .exe (built with PyInstaller) on a
factory PC that has no Python installed. Just copy sync_agent.exe and
config.json into the same folder and run the exe.

Crash safety
------------
This script is wrapped so that NOTHING can fail silently:
  - If imports fail (e.g. a missing bundled DLL), it writes the error
    to "startup_error.log" next to the exe AND keeps the console window
    open until you press Enter.
  - If config.json is missing/broken, same thing.
  - Any other uncaught error during the run also gets written there.
If you ever see an empty-looking window that closes fast, check for a
"startup_error.log" file in the same folder as the exe - it will now
always be created when something goes wrong this early.

How new rows are detected
--------------------------
MS Access has no built-in "row was just inserted" event you can listen
to from outside the application that owns the file. This agent polls
by ID instead: remembers the highest row ID already sent
(sync_state.json), checks for higher IDs every poll_interval_seconds,
sends only the new ones, and only advances the checkpoint after Odoo
confirms success. Crash-safe, network-failure-safe, never duplicates.

Author: Growise Tech
"""

import sys
import os
import traceback


# ---------------------------------------------------------------------------
# Paths - figured out before anything else, frozen-exe-safe
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CRASH_LOG_PATH = os.path.join(BASE_DIR, "startup_error.log")


def emergency_log(message):
    """Last-resort logging that works even if normal logging was never
    set up yet. Never raises, even if the disk write itself fails."""
    try:
        with open(CRASH_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(message + "\n" + ("-" * 60) + "\n")
    except Exception:
        pass


def fail_loudly(context, exc_text):
    """Print the error, save it to disk, and keep the window open so
    it can actually be read before the console closes."""
    full_message = "%s\n%s" % (context, exc_text)
    emergency_log(full_message)
    print(full_message)
    try:
        input("\nSomething went wrong - see startup_error.log. Press Enter to close...")
    except Exception:
        pass
    sys.exit(1)


# ---------------------------------------------------------------------------
# Imports that might fail if PyInstaller didn't bundle something right
# ---------------------------------------------------------------------------
try:
    import json
    import logging
    import logging.handlers
    import signal
    import time
    from datetime import datetime, date

    import pyodbc
    import requests
except Exception:
    fail_loudly("STARTUP FAILED while importing required libraries:", traceback.format_exc())


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(
            "config.json not found next to the executable.\nExpected at: %s\n"
            "Make sure config.json sits in the SAME folder as sync_agent.exe."
            % CONFIG_PATH
        )
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(
                "config.json exists but is not valid JSON (check for a missing "
                "comma or quote). Details: %s" % e
            )


try:
    CONFIG = load_config()
except Exception:
    fail_loudly("STARTUP FAILED while loading config.json:", traceback.format_exc())

STATE_PATH = os.path.join(BASE_DIR, CONFIG.get("state_file", "sync_state.json"))
LOG_PATH = os.path.join(BASE_DIR, CONFIG.get("log_file", "sync_agent.log"))


# ---------------------------------------------------------------------------
# Logging - rotates at 5MB, keeps 5 backups. Also prints to console.
# ---------------------------------------------------------------------------
try:
    logger = logging.getLogger("sync_agent")
    logger.setLevel(logging.INFO)

    file_handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(console_handler)
except Exception:
    fail_loudly("STARTUP FAILED while setting up logging (check folder write permissions):",
                traceback.format_exc())


# ---------------------------------------------------------------------------
# State (remembers the last synced row so we never resend old data)
# ---------------------------------------------------------------------------
def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_id": 0}


def save_state(state):
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
    return pyodbc.connect(conn_str, autocommit=True, readonly=True)


def json_safe(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def fetch_new_rows(last_id):
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
        if result.get("status") != "ok" or "created_ids" not in result:
            logger.error(
                "Odoo response did not confirm success - NOT advancing checkpoint. "
                "Will retry these same rows next cycle."
            )
            return
    except Exception:
        logger.exception("Failed to push rows to Odoo - will retry next cycle.")
        return

    new_last_id = max(row[id_column] for row in rows)
    save_state({"last_id": new_last_id})
    logger.info("State updated. last_id is now %s.", new_last_id)


def main():
    interval = CONFIG.get("poll_interval_seconds", 60)
    logger.info("Sync agent started. Polling every %s seconds.", interval)
    logger.info("Watching table '%s' in %s", CONFIG.get("table_name"), CONFIG.get("access_db_path"))
    while RUNNING:
        try:
            run_cycle()
        except Exception:
            logger.exception("Unexpected error in sync cycle.")
        time.sleep(interval)
    logger.info("Sync agent stopped.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        fail_loudly("UNCAUGHT ERROR WHILE RUNNING:", traceback.format_exc())
