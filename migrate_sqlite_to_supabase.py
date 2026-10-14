"""One-time migration from data/costing_quotes.db to Supabase PostgreSQL.

Run locally from this folder after setting SUPABASE_DB_URL in your environment.
The script refuses to overwrite existing PostgreSQL quotations.
"""
from __future__ import annotations
import os
import sqlite3
import sys
from pathlib import Path

import psycopg
from psycopg.rows import dict_row

BASE = Path(__file__).resolve().parent
SQLITE_PATH = Path(os.environ.get("COSTING_APP_DB", BASE / "data" / "costing_quotes.db"))
DB_URL = os.environ.get("SUPABASE_DB_URL", "").strip()
TABLES = ["settings", "catalog_items", "datecpay_tiers", "price_history", "quotes"]

if not DB_URL:
    sys.exit("Set SUPABASE_DB_URL in your local environment first. Do not commit the secret to GitHub.")
if not SQLITE_PATH.exists():
    sys.exit(f"SQLite database not found: {SQLITE_PATH}")

# Import the app's schema initializer and point it at PostgreSQL.
from pricing_engine import init_db
init_db()

source = sqlite3.connect(SQLITE_PATH)
source.row_factory = sqlite3.Row

with psycopg.connect(DB_URL, row_factory=dict_row, connect_timeout=20) as target:
    # Never overwrite a live PostgreSQL database that already has quotes.
    quote_count = target.execute("SELECT COUNT(*) AS n FROM quotes").fetchone()["n"]
    if quote_count:
        sys.exit(f"Migration stopped: Supabase already contains {quote_count} quotation(s). No data was changed.")

    # Remove only the app's default seed rows. The target quotations table must be empty.
    target.execute("DELETE FROM price_history")
    target.execute("DELETE FROM catalog_items")
    target.execute("DELETE FROM datecpay_tiers")
    target.execute("DELETE FROM settings")

    counts = {}
    for table in TABLES:
        rows = source.execute(f'SELECT * FROM "{table}"').fetchall()
        if not rows:
            counts[table] = 0
            continue
        cols = rows[0].keys()
        col_sql = ",".join('"' + c + '"' for c in cols)
        placeholders = ",".join(["%s"] * len(cols))
        sql = f'INSERT INTO "{table}" ({col_sql}) VALUES ({placeholders})'
        with target.cursor() as cursor:
            cursor.executemany(sql, [tuple(row[c] for c in cols) for row in rows])
        counts[table] = len(rows)

    # Restore serial/identity counters to the highest imported ID.
    for table in ["catalog_items", "datecpay_tiers", "price_history", "quotes"]:
        target.execute(
            "SELECT setval(pg_get_serial_sequence(%s, 'id'), GREATEST(COALESCE((SELECT MAX(id) FROM " + table + "), 1), 1), EXISTS (SELECT 1 FROM " + table + "))",
            (table,),
        )

source.close()
print("Migration completed. Rows imported:")
for table, count in counts.items():
    print(f"  {table}: {count}")
print("Verify these counts in Supabase before using the live app.")
