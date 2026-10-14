from __future__ import annotations

import json
import math
import os
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("COSTING_APP_DB", BASE_DIR / "data" / "costing_quotes.db"))

DEFAULT_TERMS = """Price: All quoted prices are VAT inclusive unless expressly stated otherwise.
Payment terms for customers with an approved Datec account:
- Software licences: Purchase Order with 100% advance payment.
- Services: 50% on commencement and 50% on final sign-off.
Customers without an approved Datec account are required to pay 100% in advance.
Quote validity: 14 days from the quotation date or the end of the month, whichever comes first.
Price volatility: Prices are based on current supplier pricing and may change due to currency fluctuations, supplier increases, or product availability. Datec reserves the right to review pricing where material changes occur beyond its control.
Electronic software delivery is generally 48-72 hours after receipt of payment, subject to supplier confirmation.
Renewals: Renewal orders should be placed at least 3 weeks before licence or contract expiry to avoid interruption.
No order cancellation: An order cannot be cancelled once the supplier purchase order has been processed.
Warranty and support are subject to the applicable product and supplier terms. Any work outside the agreed scope will be quoted separately."""

DEFAULT_SETTINGS: dict[str, Any] = {
    "company_name": "Datec (Fiji) PTE Ltd",
    "company_address": "Level 1 Garden City Complex, PO Box 12577, Suva, Fiji Islands",
    "company_phone": "+679 331 4411",
    "company_email": "info@datec.com.fj",
    "branch_address": "1st Floor Sunlight Building, Deo Street, Namaka, Nadi, Fiji Islands",
    "branch_phone": "+679 672 0181",
    "prepared_by_default": "",
    "prepared_by_phone_default": "",
    "prepared_by_email_default": "",
    "default_vat_pct": 12.5,
    "default_margin_pct": 20.0,
    "default_epicor_margin_pct": 40.0,
    "default_duty_pct": 18.0,
    "fx_rate_USD": 0.435,  # From the uploaded HFC costing workbook: foreign currency units per FJD.
    "fx_rate_AUD": 1.0,    # Placeholder; confirm before quoting in AUD.
    "fx_rate_NZD": 0.745,  # From the uploaded PayGlobal cloud costing workbook.
    "fx_rate_FJD": 1.0,
    "default_datecpay_assurance_pct": 30.0,
    "default_payglobal_assurance_pct": 50.0,
    "default_epicor_assurance_pct": 0.0,
    "default_quote_validity_days": 14,
    "amount_display": "whole",  # Customer-facing FJD amounts: whole numbers or two decimal places.
    "default_service_rate": 165.0,
    "default_datecpay_setup": 1500.0,  # Source calculator unit input; customer-quote adjustments are quote-only.
    "default_terms": DEFAULT_TERMS,
}

# Fixed sell-price presets are seeded from the submitted quotation examples / costing workbook.
CATALOG_SEED: list[dict[str, Any]] = [
    # PayGlobal Cloud: raw unit buy prices only from the supplied costing sheet.
    # Pre-calculated unit VIPs, extended totals and totals from customer quotes are deliberately not seeded.
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Payroll / HR - per employee / month", "item_type":"cost", "default_qty":90, "buy_price":2.75, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 2.75 per unit. Unit buy price only; downstream columns are recalculated."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Employee Self Service - per employee / month", "item_type":"cost", "default_qty":0, "buy_price":0.46, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 0.46 per unit."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Time and Attendance / Rostering - per employee / month", "item_type":"cost", "default_qty":0, "buy_price":1.84, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 1.84 per unit."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"MYOB PGOS Online Support - per employee / month", "item_type":"cost", "default_qty":90, "buy_price":0.21, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 0.21 per unit."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Named PayGlobal User - per user / month", "item_type":"cost", "default_qty":1, "buy_price":65.27, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 65.27 per unit. Quantity is editable."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Storage (25 GB unit) - per month", "item_type":"cost", "default_qty":2, "buy_price":43.86, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 43.86 per unit."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Cloud Configuration and Setup - one-time", "item_type":"cost", "default_qty":1, "buy_price":2500, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"Source: PayGlobal Cloud Costing Sheet, Datec Unit Buy (NZD) = 2500.00. This input price is recalculated using the selected FX, duty, margin and VAT."},
    {"product":"PayGlobal", "variant":"Cloud", "item_name":"Implementation / Training Services - hourly sell rate", "item_type":"fixed", "default_qty":49, "buy_price":0, "buy_currency":"FJD", "fixed_unit_vip":165, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"Source: ServicesCost sheet unit rate = FJD 165 per hour. Qty is hours; total is calculated as hours x unit rate."},

    # PayGlobal On-Premise: raw unit buy prices from the supplied costing sheet.
    {"product":"PayGlobal", "variant":"On-Premise", "item_name":"Payroll - per employee", "item_type":"cost", "default_qty":90, "buy_price":79, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal On-Prem Costing Sheet, Datec Unit Buy (FJD) = 79 per employee. Totals are recalculated."},
    {"product":"PayGlobal", "variant":"On-Premise", "item_name":"Human Resources - per employee", "item_type":"cost", "default_qty":0, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal On-Prem Costing Sheet, Datec Unit Buy (FJD) = 37 per employee."},
    {"product":"PayGlobal", "variant":"On-Premise", "item_name":"Time and Attendance - per employee", "item_type":"cost", "default_qty":0, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal On-Prem Costing Sheet, Datec Unit Buy (FJD) = 37 per employee."},
    {"product":"PayGlobal", "variant":"On-Premise", "item_name":"Rostering - per employee", "item_type":"cost", "default_qty":0, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal On-Prem Costing Sheet, Datec Unit Buy (FJD) = 37 per employee."},
    {"product":"PayGlobal", "variant":"On-Premise", "item_name":"Employee Self Service - per employee", "item_type":"cost", "default_qty":0, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal On-Prem Costing Sheet, Datec Unit Buy (FJD) = 37 per employee."},
    {"product":"PayGlobal", "variant":"On-Premise", "item_name":"Implementation / Training Services - hourly sell rate", "item_type":"fixed", "default_qty":49, "buy_price":0, "buy_currency":"FJD", "fixed_unit_vip":165, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"Source: ServicesCost sheet unit rate = FJD 165 per hour. Qty is hours; total is calculated as hours x unit rate."},

    # PayGlobal Hybrid: the hybrid worksheet has local licence unit inputs plus PaaS supplier unit buy inputs.
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Payroll Licence - per employee", "item_type":"cost", "default_qty":160, "buy_price":79, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing, local licence unit buy price = FJD 79. Totals are recalculated."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Human Resources Licence - per employee", "item_type":"cost", "default_qty":0, "list_price":36, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing, local licence unit buy price = FJD 37."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Time and Attendance Licence - per employee", "item_type":"cost", "default_qty":0, "list_price":36, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing, local licence unit buy price = FJD 37."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Rostering Licence - per employee", "item_type":"cost", "default_qty":0, "list_price":36, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing, local licence unit buy price = FJD 37."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Employee Self Service Licence - per employee", "item_type":"cost", "default_qty":0, "list_price":35, "buy_price":37, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing, local licence unit buy price = FJD 37."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Payroll / HR - PaaS per employee / month", "item_type":"cost", "default_qty":160, "buy_price":0.66, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing PaaS section, unit buy price = NZD 0.66."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"MYOB PGOS Support - per employee / month", "item_type":"cost", "default_qty":160, "buy_price":0.21, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing PaaS section, unit buy price = NZD 0.21."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Named PayGlobal User - per user / month", "item_type":"cost", "default_qty":2, "buy_price":65.27, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing PaaS section, unit buy price = NZD 65.27."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Storage (25 GB unit) - per month", "item_type":"cost", "default_qty":1, "buy_price":43.86, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"Monthly", "assurance_eligible":1, "notes":"Source: PayGlobal Hybrid Costing PaaS section, unit buy price = NZD 43.86."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Cloud Configuration and Setup - one-time", "item_type":"cost", "default_qty":1, "buy_price":2500, "buy_currency":"NZD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"Source: PayGlobal Hybrid Costing PaaS section, unit buy price = NZD 2500.00."},
    {"product":"PayGlobal", "variant":"Hybrid", "item_name":"Implementation / Training Services - hourly sell rate", "item_type":"fixed", "default_qty":49, "buy_price":0, "buy_currency":"FJD", "fixed_unit_vip":165, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"Source: ServicesCost sheet unit rate = FJD 165 per hour. Qty is hours; total is calculated as hours x unit rate."},

    # Epicor: raw unit buy price only from the HFC NZ costing sheet. Other project services remain quote-entry prices.
    {"product":"Epicor", "variant":"Kinetic Upgrade", "item_name":"ATS Upgrade existing Kinetic installation to latest release - technical unit cost", "item_type":"cost", "default_qty":1, "list_price":0, "buy_price":0, "buy_currency":"USD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"Epicor supplier cost changes for each project. Enter the current unit buy price in every quote; default cost is intentionally zero and quantity is one."},
    {"product":"Epicor", "variant":"Kinetic Upgrade", "item_name":"Project Management / UAT / Go-Live Services (enter approved quote-only unit price if required)", "item_type":"cost", "default_qty":0, "buy_price":0, "buy_currency":"FJD", "discount_pct":0, "freight_pct":0, "billing_frequency":"One-time", "assurance_eligible":0, "notes":"The costing sheet does not provide a supplier unit buy price for this line. Do not use the previous calculated balancing amount; enter an approved quote-only unit VIP override or leave excluded."},
]
DATECPAY_TIERS = [
    (1, 1, 10, 20.0, 0.0),
    (2, 11, 20, 18.0, 10.0),
    (3, 21, 30, 16.0, 20.0),
    (4, 31, 40, 14.0, 30.0),
    (5, 41, 50, 12.0, 40.0),
    (6, 51, 60, 6.0, 50.0),
    (7, 61, 70, 8.0, 60.0),
    (8, 71, 80, 6.0, 70.0),
    (9, 81, 90, 4.0, 80.0),
    (10, 91, 100, 2.0, 90.0),
]


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none", "<na>"} else text


def _parse_id(value: Any) -> int | None:
    text = _clean_text(value)
    if not text:
        return None
    try:
        number = float(text)
        if number > 0 and number.is_integer():
            return int(number)
    except (ValueError, TypeError):
        return None
    return None


def _number(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or _clean_text(value) == "":
            return default
        number = float(value)
        return number if number == number and abs(number) != float("inf") else default
    except (ValueError, TypeError):
        return default


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _migrate_v1_1_unit_price_catalog(conn: sqlite3.Connection) -> None:
    """Remove v1.0 quote-derived seed prices while preserving editable source unit costs."""
    marker = conn.execute("SELECT value FROM settings WHERE key=?", ("catalog_unit_price_migration_v1_1",)).fetchone()
    if marker:
        try:
            if json.loads(marker["value"]) is True:
                return
        except (json.JSONDecodeError, TypeError):
            pass

    seed_by_key = {(str(x["product"]), str(x["variant"]), str(x["item_name"])): dict(x) for x in CATALOG_SEED}
    aliases = {
        ("PayGlobal", "Cloud", "Payroll / HR"): "Payroll / HR - per employee / month",
        ("PayGlobal", "Cloud", "Employee Self Service"): "Employee Self Service - per employee / month",
        ("PayGlobal", "Cloud", "Time and Attendance / Rostering"): "Time and Attendance / Rostering - per employee / month",
        ("PayGlobal", "Cloud", "MYOB PGOS Online Support"): "MYOB PGOS Online Support - per employee / month",
        ("PayGlobal", "Cloud", "Named PayGlobal Users"): "Named PayGlobal User - per user / month",
        ("PayGlobal", "Cloud", "Storage (25 GB units)"): "Storage (25 GB unit) - per month",
        ("PayGlobal", "Cloud", "Cloud Configuration and Setup"): "Cloud Configuration and Setup - one-time",
        ("PayGlobal", "On-Premise", "Payroll"): "Payroll - per employee",
        ("PayGlobal", "On-Premise", "Human Resources"): "Human Resources - per employee",
        ("PayGlobal", "On-Premise", "Time and Attendance"): "Time and Attendance - per employee",
        ("PayGlobal", "On-Premise", "Rostering"): "Rostering - per employee",
        ("PayGlobal", "On-Premise", "Employee Self Service"): "Employee Self Service - per employee",
        ("PayGlobal", "Hybrid", "Payroll Licence"): "Payroll Licence - per employee",
        ("PayGlobal", "Hybrid", "Human Resources Licence"): "Human Resources Licence - per employee",
        ("PayGlobal", "Hybrid", "Time and Attendance Licence"): "Time and Attendance Licence - per employee",
        ("PayGlobal", "Hybrid", "Rostering Licence"): "Rostering Licence - per employee",
        ("PayGlobal", "Hybrid", "Employee Self Service Licence"): "Employee Self Service Licence - per employee",
        ("PayGlobal", "Hybrid", "Payroll / HR - PaaS"): "Payroll / HR - PaaS per employee / month",
        ("PayGlobal", "Hybrid", "MYOB PGOS Support"): "MYOB PGOS Support - per employee / month",
        ("PayGlobal", "Hybrid", "Named PayGlobal Users"): "Named PayGlobal User - per user / month",
        ("PayGlobal", "Hybrid", "Storage (25 GB units)"): "Storage (25 GB unit) - per month",
        ("PayGlobal", "Hybrid", "Cloud Configuration and Setup"): "Cloud Configuration and Setup - one-time",
        ("Epicor", "Kinetic Upgrade", "Epicor Kinetic Upgrade - Technical Services"): "ATS Upgrade existing Kinetic installation to latest release - technical unit cost",
    }
    raw_rows = conn.execute("SELECT * FROM catalog_items ORDER BY id").fetchall()
    preserved: dict[tuple[str, str, str], dict[str, Any]] = {}
    legacy_patterns = (
        "sample quotation fixed price", "costing workbook default", "standard service sell-price preset",
        "standard vip sell-price preset", "fixed vip amount inferred", "hfc costing workbook input",
        "hfc proposal sample price", "source costing sheet uses", "source workbook defaults", "monthly buy price per employee; source workbook defaults",
        "source costing sheet default", "imported from payglobal hybrid costing", "monthly cost per employee; source workbook defaults",
        "standard service sell-price preset from the supplied customer quotation",
    )
    delete_ids: list[int] = []
    for old in raw_rows:
        row = dict(old)
        product, variant, name = str(row.get("product", "")), str(row.get("variant", "")), str(row.get("item_name", ""))
        note = str(row.get("notes", "")).lower()
        is_legacy_seed = any(pattern in note for pattern in legacy_patterns)
        # v1.0 used quote-derived VIP amounts inside cost rows as well as fixed rows.
        if str(row.get("item_type", "")) == "cost" and float(row.get("fixed_unit_vip") or 0) > 0 and (
            "preset" in note or "inferred" in note or "balance" in name.lower()
        ):
            is_legacy_seed = True
        # Preserve original buy-unit price edits, using explicit aliases for the v1.0 row labels.
        if str(row.get("item_type", "")) == "cost" and float(row.get("buy_price") or 0) > 0 and float(row.get("fixed_unit_vip") or 0) == 0:
            target_name = aliases.get((product, variant, name), name)
            target_key = (product, variant, target_name)
            if target_key in seed_by_key and target_key not in preserved:
                preserved[target_key] = {
                    "list_price": float(row.get("list_price") or row.get("buy_price") or 0),
                    "buy_price": float(row.get("buy_price") or 0), "buy_currency": row.get("buy_currency") or "FJD",
                    "discount_pct": float(row.get("discount_pct") or 0), "freight_pct": float(row.get("freight_pct") or 0),
                    "default_qty": float(row.get("default_qty") or 0), "billing_frequency": row.get("billing_frequency") or seed_by_key[target_key].get("billing_frequency", "One-time"),
                    "assurance_eligible": int(row.get("assurance_eligible") or 0),
                }
        if is_legacy_seed:
            delete_ids.append(int(row["id"]))

    if delete_ids:
        conn.executemany("DELETE FROM catalog_items WHERE id=?", [(item_id,) for item_id in delete_ids])

    now = datetime.now().isoformat(timespec="seconds")
    for key, seed in seed_by_key.items():
        existing = conn.execute("SELECT id FROM catalog_items WHERE product=? AND variant=? AND item_name=?", key).fetchone()
        if existing:
            continue
        row = {"default_qty": 1, "list_price": 0, "buy_price": 0, "buy_currency": "FJD", "discount_pct": 0,
               "freight_pct": 0, "fixed_unit_vip": 0, "billing_frequency": "One-time",
               "assurance_eligible": 0, "notes": "", **seed}
        row.update(preserved.get(key, {}))
        conn.execute("""INSERT INTO catalog_items
            (product, variant, item_name, item_type, default_qty, list_price, buy_price, buy_currency,
             discount_pct, freight_pct, fixed_unit_vip, billing_frequency, assurance_eligible, notes, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (row["product"], row["variant"], row["item_name"], row["item_type"], row["default_qty"],
             (row.get("list_price") or row["buy_price"]), row["buy_price"], row["buy_currency"], row["discount_pct"], row["freight_pct"],
             row["fixed_unit_vip"], row["billing_frequency"], int(row["assurance_eligible"]), row["notes"], now))

    # The previous app's setup default was the altered sample quote amount. Use the costing-sheet unit input by default.
    setup_row = conn.execute("SELECT value FROM settings WHERE key='default_datecpay_setup'").fetchone()
    if setup_row:
        try:
            setup_value = json.loads(setup_row["value"])
        except (json.JSONDecodeError, TypeError):
            setup_value = setup_row["value"]
        if isinstance(setup_value, (int, float)) and abs(float(setup_value) - 2500.0) < 1e-9:
            conn.execute("UPDATE settings SET value=? WHERE key='default_datecpay_setup'", (json.dumps(1500.0),))
    conn.execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                 ("catalog_unit_price_migration_v1_1", json.dumps(True)))


def _migrate_payglobal_unit_prices_v1_2(conn: sqlite3.Connection) -> None:
    """One-time correction of known PayGlobal unit prices affected by whole-number editor steps.

    This migration only changes the old seeded/rounded values once. Later supplier
    updates in the Price Book are left untouched.
    """
    marker_key = "payglobal_unit_price_precision_migration_v1_2"
    marker = conn.execute("SELECT value FROM settings WHERE key=?", (marker_key,)).fetchone()
    if marker:
        try:
            if json.loads(marker["value"]) is True:
                return
        except (json.JSONDecodeError, TypeError):
            pass

    corrections = [
        # Variant, item name, old rounded value(s), source unit-buy price, currency
        ("Cloud", "Payroll / HR - per employee / month", {2.0}, 2.75, "NZD"),
        ("Cloud", "Employee Self Service - per employee / month", {0.0}, 0.46, "NZD"),
        ("Cloud", "Time and Attendance / Rostering - per employee / month", {1.0}, 1.84, "NZD"),
        ("Cloud", "MYOB PGOS Online Support - per employee / month", {0.0}, 0.21, "NZD"),
        ("Cloud", "Named PayGlobal User - per user / month", {65.0}, 65.27, "NZD"),
        ("Cloud", "Storage (25 GB unit) - per month", {43.0}, 43.86, "NZD"),
        ("On-Premise", "Payroll - per employee", {65.0}, 79.0, "FJD"),
        ("On-Premise", "Human Resources - per employee", {35.0}, 37.0, "FJD"),
        ("On-Premise", "Time and Attendance - per employee", {35.0}, 37.0, "FJD"),
        ("On-Premise", "Rostering - per employee", {35.0}, 37.0, "FJD"),
        ("On-Premise", "Employee Self Service - per employee", {35.0}, 37.0, "FJD"),
        ("Hybrid", "Payroll Licence - per employee", {65.0}, 79.0, "FJD"),
        ("Hybrid", "Human Resources Licence - per employee", {35.0}, 37.0, "FJD"),
        ("Hybrid", "Time and Attendance Licence - per employee", {35.0}, 37.0, "FJD"),
        ("Hybrid", "Rostering Licence - per employee", {35.0}, 37.0, "FJD"),
        ("Hybrid", "Employee Self Service Licence - per employee", {35.0}, 37.0, "FJD"),
    ]
    now = datetime.now().isoformat(timespec="seconds")
    for variant, item_name, old_values, correct_price, currency in corrections:
        item = conn.execute(
            "SELECT id,buy_price,buy_currency FROM catalog_items WHERE product='PayGlobal' AND variant=? AND item_name=?",
            (variant, item_name),
        ).fetchone()
        if not item:
            continue
        current = float(item["buy_price"] or 0)
        if abs(current - correct_price) < 1e-9:
            continue
        if current in old_values and str(item["buy_currency"] or "").upper() == currency:
            conn.execute(
                "INSERT INTO price_history(catalog_item_id,item_name,field_name,old_value,new_value,changed_at) VALUES (?,?,?,?,?,?)",
                (int(item["id"]), item_name, "buy_price (precision/source correction)", str(current), str(correct_price), now),
            )
            conn.execute("UPDATE catalog_items SET buy_price=?,updated_at=? WHERE id=?", (correct_price, now, int(item["id"])))

    # Repair List Price from the source costing sheets once. Do not use previous quote totals.
    # Cloud list prices mirror source unit prices; On-Prem uses FJD list-price inputs;
    # Hybrid has distinct local licence list prices plus the PaaS unit list prices.
    source_list_prices = {
        ("Cloud", "Payroll / HR - per employee / month"): 2.75,
        ("Cloud", "Employee Self Service - per employee / month"): 0.46,
        ("Cloud", "Time and Attendance / Rostering - per employee / month"): 1.84,
        ("Cloud", "MYOB PGOS Online Support - per employee / month"): 0.21,
        ("Cloud", "Named PayGlobal User - per user / month"): 65.27,
        ("Cloud", "Storage (25 GB unit) - per month"): 43.86,
        ("Cloud", "Cloud Configuration and Setup - one-time"): 2500.00,
        ("On-Premise", "Payroll - per employee"): 79.00,
        ("On-Premise", "Human Resources - per employee"): 37.00,
        ("On-Premise", "Time and Attendance - per employee"): 37.00,
        ("On-Premise", "Rostering - per employee"): 37.00,
        ("On-Premise", "Employee Self Service - per employee"): 37.00,
        ("Hybrid", "Payroll Licence - per employee"): 79.00,
        ("Hybrid", "Human Resources Licence - per employee"): 36.00,
        ("Hybrid", "Time and Attendance Licence - per employee"): 36.00,
        ("Hybrid", "Rostering Licence - per employee"): 36.00,
        ("Hybrid", "Employee Self Service Licence - per employee"): 35.00,
        ("Hybrid", "Payroll / HR - PaaS per employee / month"): 0.66,
        ("Hybrid", "MYOB PGOS Support - per employee / month"): 0.21,
        ("Hybrid", "Named PayGlobal User - per user / month"): 65.27,
        ("Hybrid", "Storage (25 GB unit) - per month"): 43.86,
        ("Hybrid", "Cloud Configuration and Setup - one-time"): 2500.00,
    }
    for (variant, item_name), list_price in source_list_prices.items():
        item = conn.execute(
            "SELECT id,list_price FROM catalog_items WHERE product='PayGlobal' AND variant=? AND item_name=?",
            (variant, item_name),
        ).fetchone()
        if not item:
            continue
        current_list = float(item["list_price"] or 0)
        if abs(current_list - list_price) > 1e-9:
            conn.execute(
                "INSERT INTO price_history(catalog_item_id,item_name,field_name,old_value,new_value,changed_at) VALUES (?,?,?,?,?,?)",
                (int(item["id"]), item_name, "list_price (source correction)", str(current_list), str(list_price), now),
            )
            conn.execute("UPDATE catalog_items SET list_price=?,updated_at=? WHERE id=?",
                         (list_price, now, int(item["id"])))
    # Other legacy/new items still get a sensible reference list price when none is configured.
    conn.execute("UPDATE catalog_items SET list_price=buy_price WHERE list_price=0 AND buy_price>0")

    # Standard assurance percentage from the uploaded On-Prem costing sheet is 50%.
    setting = conn.execute("SELECT value FROM settings WHERE key='default_payglobal_assurance_pct'").fetchone()
    if setting:
        try:
            current_assurance = float(json.loads(setting["value"]))
        except (json.JSONDecodeError, TypeError, ValueError):
            current_assurance = None
        if current_assurance == 40.0:
            conn.execute("UPDATE settings SET value=? WHERE key='default_payglobal_assurance_pct'", (json.dumps(50.0),))

    conn.execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                 (marker_key, json.dumps(True)))


def _migrate_epicor_dynamic_cost_defaults_v1_6(conn: sqlite3.Connection) -> None:
    """Reset the prior seeded Epicor cost to zero once; each project cost is quote-specific."""
    marker_key = "epicor_dynamic_cost_defaults_v1_6"
    marker = conn.execute("SELECT value FROM settings WHERE key=?", (marker_key,)).fetchone()
    if marker:
        try:
            if json.loads(marker["value"]) is True:
                return
        except (json.JSONDecodeError, TypeError):
            pass

    now = datetime.now().isoformat(timespec="seconds")
    row = conn.execute(
        "SELECT id, item_name, buy_price, list_price, default_qty FROM catalog_items "
        "WHERE product='Epicor' AND variant='Kinetic Upgrade' "
        "AND lower(item_name) LIKE '%technical unit cost%' LIMIT 1"
    ).fetchone()
    if row:
        if float(row["buy_price"] or 0) != 0 or float(row["list_price"] or 0) != 0:
            conn.execute(
                "INSERT INTO price_history(catalog_item_id,item_name,field_name,old_value,new_value,changed_at) "
                "VALUES (?,?,?,?,?,?)",
                (int(row["id"]), row["item_name"], "buy_price (Epicor dynamic-cost reset)",
                 str(row["buy_price"]), "0", now),
            )
        conn.execute(
            "UPDATE catalog_items SET buy_price=0, list_price=0, default_qty=1, "
            "notes=?, updated_at=? WHERE id=?",
            ("Epicor supplier cost changes for each project. Enter the current unit buy price in every quote; default cost is intentionally zero and quantity is one.", now, int(row["id"])),
        )
    # Do not store a prior project-specific price in the optional service cost line either.
    conn.execute(
        "UPDATE catalog_items SET buy_price=0, list_price=0, updated_at=? "
        "WHERE product='Epicor' AND variant='Kinetic Upgrade' AND item_name LIKE 'Project Management / UAT / Go-Live Services%'",
        (now,),
    )
    conn.execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                 (marker_key, json.dumps(True)))


def init_db() -> None:
    with connect() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS catalog_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product TEXT NOT NULL,
            variant TEXT NOT NULL,
            item_name TEXT NOT NULL,
            item_type TEXT NOT NULL DEFAULT 'fixed',
            default_qty REAL NOT NULL DEFAULT 1,
            list_price REAL NOT NULL DEFAULT 0,
            buy_price REAL NOT NULL DEFAULT 0,
            buy_currency TEXT NOT NULL DEFAULT 'FJD',
            discount_pct REAL NOT NULL DEFAULT 0,
            freight_pct REAL NOT NULL DEFAULT 0,
            fixed_unit_vip REAL NOT NULL DEFAULT 0,
            billing_frequency TEXT NOT NULL DEFAULT 'One-time',
            assurance_eligible INTEGER NOT NULL DEFAULT 0,
            notes TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS datecpay_tiers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tier_no INTEGER NOT NULL,
            min_employees INTEGER NOT NULL,
            max_employees INTEGER NOT NULL,
            monthly_price REAL NOT NULL,
            discount_pct REAL NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS price_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            catalog_item_id INTEGER,
            item_name TEXT NOT NULL,
            field_name TEXT NOT NULL,
            old_value TEXT,
            new_value TEXT,
            changed_at TEXT NOT NULL,
            FOREIGN KEY(catalog_item_id) REFERENCES catalog_items(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS quotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quote_no TEXT UNIQUE NOT NULL,
            quote_date TEXT NOT NULL,
            customer TEXT NOT NULL,
            organisation TEXT NOT NULL,
            email TEXT NOT NULL,
            phone TEXT NOT NULL,
            subject TEXT NOT NULL,
            product TEXT NOT NULL,
            variant TEXT NOT NULL,
            pricing_mode TEXT NOT NULL,
            subtotal REAL NOT NULL,
            assurance_total REAL NOT NULL DEFAULT 0,
            grand_total REAL NOT NULL,
            quote_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Draft',
            loss_reason TEXT NOT NULL DEFAULT ''
        );
        """)
        quote_columns = {r[1] for r in conn.execute("PRAGMA table_info(quotes)").fetchall()}
        if "loss_reason" not in quote_columns:
            conn.execute("ALTER TABLE quotes ADD COLUMN loss_reason TEXT NOT NULL DEFAULT ''")
        catalog_columns = {r[1] for r in conn.execute("PRAGMA table_info(catalog_items)").fetchall()}
        if "list_price" not in catalog_columns:
            conn.execute("ALTER TABLE catalog_items ADD COLUMN list_price REAL NOT NULL DEFAULT 0")
        existing_settings = {r["key"] for r in conn.execute("SELECT key FROM settings")}
        for key, value in DEFAULT_SETTINGS.items():
            if key not in existing_settings:
                conn.execute("INSERT INTO settings(key, value) VALUES (?, ?)", (key, json.dumps(value)))
        if conn.execute("SELECT COUNT(*) AS n FROM catalog_items").fetchone()["n"] == 0:
            now = datetime.now().isoformat(timespec="seconds")
            for item in CATALOG_SEED:
                row = {"default_qty": 1, "list_price": 0, "buy_price": 0, "buy_currency": "FJD", "discount_pct": 0,
                       "freight_pct": 0, "fixed_unit_vip": 0, "billing_frequency": "One-time",
                       "assurance_eligible": 0, "notes": "", **item}
                conn.execute("""INSERT INTO catalog_items
                    (product, variant, item_name, item_type, default_qty, list_price, buy_price, buy_currency,
                     discount_pct, freight_pct, fixed_unit_vip, billing_frequency, assurance_eligible, notes, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (row["product"], row["variant"], row["item_name"], row["item_type"], row["default_qty"],
                     (row.get("list_price") or row["buy_price"]), row["buy_price"], row["buy_currency"], row["discount_pct"], row["freight_pct"],
                     row["fixed_unit_vip"], row["billing_frequency"], int(row["assurance_eligible"]), row["notes"], now))
        if conn.execute("SELECT COUNT(*) AS n FROM datecpay_tiers").fetchone()["n"] == 0:
            conn.executemany("INSERT INTO datecpay_tiers(tier_no,min_employees,max_employees,monthly_price,discount_pct) VALUES (?,?,?,?,?)", DATECPAY_TIERS)
        _migrate_v1_1_unit_price_catalog(conn)
        _migrate_payglobal_unit_prices_v1_2(conn)
        _migrate_epicor_dynamic_cost_defaults_v1_6(conn)
        conn.commit()


def get_settings() -> dict[str, Any]:
    init_db()
    with connect() as conn:
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
    values = dict(DEFAULT_SETTINGS)
    for row in rows:
        try:
            values[row["key"]] = json.loads(row["value"])
        except (json.JSONDecodeError, TypeError):
            values[row["key"]] = row["value"]
    return values


def save_settings(values: dict[str, Any]) -> None:
    with connect() as conn:
        for key, value in values.items():
            conn.execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                         (key, json.dumps(value)))
        conn.commit()


def list_catalog(product: str | None = None, variant: str | None = None, item_type: str | None = None) -> list[dict[str, Any]]:
    init_db()
    clauses, args = [], []
    if product:
        clauses.append("product = ?"); args.append(product)
    if variant:
        clauses.append("variant = ?"); args.append(variant)
    if item_type:
        clauses.append("item_type = ?"); args.append(item_type)
    sql = "SELECT * FROM catalog_items" + (" WHERE " + " AND ".join(clauses) if clauses else "") + " ORDER BY product, variant, item_type, id"
    with connect() as conn:
        return [dict(r) for r in conn.execute(sql, args).fetchall()]


def list_tiers() -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        return [dict(r) for r in conn.execute("SELECT * FROM datecpay_tiers ORDER BY min_employees").fetchall()]


def save_tiers(rows: list[dict[str, Any]]) -> None:
    with connect() as conn:
        now = datetime.now().isoformat(timespec="seconds")
        current = conn.execute("SELECT * FROM datecpay_tiers").fetchall()
        incoming_ids = {parsed for parsed in (_parse_id(r.get("id")) for r in rows) if parsed is not None}
        for old in current:
            if old["id"] not in incoming_ids:
                conn.execute("DELETE FROM datecpay_tiers WHERE id=?", (old["id"],))
        for r in rows:
            if _clean_text(r.get("min_employees")) == "" or _clean_text(r.get("max_employees")) == "":
                continue
            vals = (int(_number(r.get("tier_no"), 0)), int(_number(r.get("min_employees"), 0)), int(_number(r.get("max_employees"), 0)),
                    _number(r.get("monthly_price"), 0), _number(r.get("discount_pct"), 0))
            tier_id = _parse_id(r.get("id"))
            if tier_id is not None:
                old = conn.execute("SELECT * FROM datecpay_tiers WHERE id=?", (tier_id,)).fetchone()
                if old:
                    fields = ["tier_no", "min_employees", "max_employees", "monthly_price", "discount_pct"]
                    for field, new_value in zip(fields, vals):
                        if float(old[field]) != float(new_value):
                            field_label = field
                            item_label = f"DatecPay Tier {vals[0]} ({vals[1]}-{vals[2]} employees)"
                            conn.execute("INSERT INTO price_history(catalog_item_id,item_name,field_name,old_value,new_value,changed_at) VALUES (NULL,?,?,?,?,?)",
                                         (item_label, field_label, str(old[field]), str(new_value), now))
                    conn.execute("UPDATE datecpay_tiers SET tier_no=?,min_employees=?,max_employees=?,monthly_price=?,discount_pct=? WHERE id=?", (*vals, tier_id))
                else:
                    conn.execute("INSERT INTO datecpay_tiers(tier_no,min_employees,max_employees,monthly_price,discount_pct) VALUES (?,?,?,?,?)", vals)
            else:
                conn.execute("INSERT INTO datecpay_tiers(tier_no,min_employees,max_employees,monthly_price,discount_pct) VALUES (?,?,?,?,?)", vals)
        conn.commit()


def save_catalog(rows: list[dict[str, Any]]) -> int:
    """Update price book and log changes to supplier prices / assumptions."""
    changed = 0
    now = datetime.now().isoformat(timespec="seconds")
    fields = ["product", "variant", "item_name", "item_type", "default_qty", "list_price", "buy_price", "buy_currency",
              "discount_pct", "freight_pct", "fixed_unit_vip", "billing_frequency", "assurance_eligible", "notes"]
    numeric_fields = {"default_qty", "list_price", "buy_price", "discount_pct", "freight_pct", "fixed_unit_vip"}
    with connect() as conn:
        for raw in rows:
            name = _clean_text(raw.get("item_name"))
            product = _clean_text(raw.get("product"))
            variant = _clean_text(raw.get("variant"))
            if not name or not product or not variant:
                continue
            data: dict[str, Any] = {}
            for field in fields:
                val = raw.get(field)
                if field in numeric_fields:
                    try:
                        parsed = float(val) if val is not None and str(val).strip() not in ("", "nan", "None") else 0.0
                        val = parsed if parsed == parsed and abs(parsed) != float("inf") else 0.0
                    except (ValueError, TypeError):
                        val = 0.0
                elif field == "assurance_eligible":
                    val = int(val is True or str(val).strip().lower() in {"1", "true", "yes", "y"})
                else:
                    val = _clean_text(val)
                data[field] = val
            item_id = _parse_id(raw.get("id"))
            if item_id is not None:
                old_row = conn.execute("SELECT * FROM catalog_items WHERE id=?", (item_id,)).fetchone()
                if old_row:
                    for f in ["list_price", "buy_price", "buy_currency", "discount_pct", "freight_pct", "fixed_unit_vip", "default_qty"]:
                        old_value, new_value = old_row[f], data[f]
                        if str(old_value) != str(new_value):
                            conn.execute("INSERT INTO price_history(catalog_item_id,item_name,field_name,old_value,new_value,changed_at) VALUES (?,?,?,?,?,?)",
                                         (item_id, name, f, str(old_value), str(new_value), now))
                            changed += 1
                    assignments = ",".join(f"{f}=?" for f in fields) + ",updated_at=?"
                    conn.execute(f"UPDATE catalog_items SET {assignments} WHERE id=?", (*[data[f] for f in fields], now, item_id))
                    continue
            conn.execute(f"INSERT INTO catalog_items ({','.join(fields)},updated_at) VALUES ({','.join('?' for _ in fields)},?)",
                         (*[data[f] for f in fields], now))
        conn.commit()
    return changed


def price_history(limit: int = 50) -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        return [dict(r) for r in conn.execute("SELECT * FROM price_history ORDER BY id DESC LIMIT ?", (limit,)).fetchall()]


def fx_rates(settings: dict[str, Any] | None = None) -> dict[str, float]:
    s = settings or get_settings()
    out = {}
    for currency in ["USD", "AUD", "NZD", "FJD"]:
        try: out[currency] = float(s.get(f"fx_rate_{currency}", 1) or 1)
        except (ValueError, TypeError): out[currency] = 1.0
    return out


def money(value: float) -> str:
    return f"FJD {float(value or 0):,.2f}"


def _float(row: dict[str, Any], key: str, default: float = 0.0) -> float:
    try:
        v = row.get(key, default)
        if v is None or v == "": return default
        return float(v)
    except (ValueError, TypeError):
        return default


def _round_up_2dp(value: float) -> float:
    """Match PayGlobal workbook: ROUNDUP calculated unit VIP FJD to 2 decimals."""
    return float(math.ceil(max(0.0, float(value or 0.0)) * 100 - 1e-9) / 100.0)


def calculate_cost_quote(
    cost_rows: list[dict[str, Any]], *,
    selected_currency: str,
    selected_fx_rate: float,
    all_fx_rates: dict[str, float],
    vat_pct: float,
    margin_pct: float,
    duty_pct: float,
    assurance_enabled: bool = False,
    assurance_pct: float = 0.0,
    assurance_basis: str = "Annual licence VIP",
    round_unit_vip_2dp: bool = False,
) -> dict[str, Any]:
    """Cost-plus model. FX convention is foreign-currency units per 1 FJD, matching the submitted workbooks.

    When round_unit_vip_2dp is enabled, calculated PayGlobal unit VIP is rounded up
    to 2 decimal places before extending by quantity and billing periods, matching Excel.
    Extended totals retain calculation precision and are displayed as whole FJD in views.
    """
    quote_lines: list[dict[str, Any]] = []
    eligible_vip_annual = 0.0
    eligible_landed_annual = 0.0
    subtotal = 0.0
    for row in cost_rows:
        name = str(row.get("item_name") or row.get("description") or "Custom item").strip()
        qty = max(0.0, _float(row, "default_qty", _float(row, "qty", 1)))
        buy_price = max(0.0, _float(row, "buy_price", 0))
        discount_pct = min(100.0, max(0.0, _float(row, "discount_pct", 0)))
        freight_pct = max(-99.0, _float(row, "freight_pct", 0))
        currency = str(row.get("buy_currency") or "FJD").upper()
        frequency = str(row.get("billing_frequency") or "One-time")
        fixed_price = max(0.0, _float(row, "fixed_unit_vip", 0))
        raw_quote_override = row.get("quote_unit_override")
        quote_override = None
        if raw_quote_override is not None and str(raw_quote_override).strip().lower() not in ("", "nan", "none"):
            try:
                quote_override = max(0.0, float(raw_quote_override))
            except (ValueError, TypeError):
                quote_override = None
        if currency == selected_currency:
            fx = float(selected_fx_rate or 1.0)
        else:
            fx = float(all_fx_rates.get(currency, 1.0) or 1.0)
        if currency == "FJD":
            fx = 1.0
        if fx <= 0:
            fx = 1.0

        # Supplier unit buy price -> landed cost -> FJD -> margin -> VAT.
        # Some costing sheets use a different margin factor by row (e.g., Hybrid local
        # licence vs. PaaS subscription), so allow a per-row override while preserving
        # the shared margin default for all other rows.
        row_margin_pct = _float(row, "margin_pct_override", margin_pct)
        unit_landed_fjd = 0.0
        if quote_override is not None:
            unit_vip = quote_override
            # Keep landed cost independently calculated for internal costing / assurance basis.
            net_buy = buy_price * (1.0 - discount_pct / 100.0)
            landed_foreign = net_buy * (1.0 + freight_pct / 100.0) * (1.0 + duty_pct / 100.0)
            unit_landed_fjd = landed_foreign / fx
        elif fixed_price > 0:
            unit_vip = fixed_price
        else:
            net_buy = buy_price * (1.0 - discount_pct / 100.0)
            landed_foreign = net_buy * (1.0 + freight_pct / 100.0) * (1.0 + duty_pct / 100.0)
            unit_landed_fjd = landed_foreign / fx
            unit_vip = unit_landed_fjd * (1.0 + row_margin_pct / 100.0) * (1.0 + vat_pct / 100.0)
        if round_unit_vip_2dp and quote_override is None and fixed_price <= 0:
            unit_vip = _round_up_2dp(unit_vip)
        if qty <= 0:
            continue
        amount = unit_vip * qty
        landed_total = unit_landed_fjd * qty
        eligible = bool(row.get("assurance_eligible", False))

        if frequency.lower().startswith("month"):
            out_qty = 12
            out_unit = amount
            out_total = amount * 12
            out_desc = f"{name} ({qty:g} units; recurring annually)"
            billing = "Annual recurring"
        elif frequency.lower().startswith("quarter"):
            out_qty = 4
            out_unit = amount
            out_total = amount * 4
            out_desc = f"{name} ({qty:g} units; recurring quarterly)"
            billing = "Quarterly recurring"
        elif frequency.lower().startswith("annual"):
            out_qty = qty
            out_unit = unit_vip
            out_total = amount
            out_desc = name
            billing = "Annual recurring"
        else:
            out_qty = qty
            out_unit = unit_vip
            out_total = amount
            out_desc = name
            billing = "One-time"

        if eligible:
            eligible_vip_annual += out_total
            if frequency.lower().startswith("month"):
                eligible_landed_annual += landed_total * 12
            elif frequency.lower().startswith("quarter"):
                eligible_landed_annual += landed_total * 4
            else:
                # One-time licence costs can still be the base for recurring Assurance.
                eligible_landed_annual += landed_total

        quote_lines.append({"description": out_desc, "qty": out_qty, "unit_vip": out_unit,
                            "total_vip": out_total, "billing": billing, "assurance_eligible": eligible,
                            "buy_currency": currency, "buy_unit_price": buy_price,
                            "landed_cost_fjd": landed_total})
        subtotal += out_total

    assurance = 0.0
    if assurance_enabled and assurance_pct > 0:
        base = eligible_landed_annual if assurance_basis == "Annual landed cost" else eligible_vip_annual
        assurance = base * assurance_pct / 100.0
        if assurance > 0:
            quote_lines.append({"description": "Assurance", "qty": 1,
                                "unit_vip": assurance, "total_vip": assurance, "billing": "Annual recurring",
                                "assurance_eligible": False, "buy_currency": "FJD", "buy_unit_price": 0,
                                "landed_cost_fjd": 0})
            subtotal += assurance
    return {"lines": quote_lines, "subtotal": subtotal, "assurance_total": assurance,
            "grand_total": subtotal, "assurance_basis": assurance_basis,
            "eligible_annual_vip": eligible_vip_annual, "eligible_annual_landed": eligible_landed_annual}


def calculate_datecpay(employees: int, tiers: list[dict[str, Any]], vat_pct: float, setup_fee: float,
                        assurance_enabled: bool, assurance_pct: float,
                        custom_license_vip: float | None = None) -> dict[str, Any]:
    employee_count = max(1, int(employees))
    tier = next((t for t in tiers if int(t["min_employees"]) <= employee_count <= int(t["max_employees"])), None)
    if custom_license_vip is None:
        if not tier:
            monthly_price = 0.0
            tier_desc = "No matching tier - update DatecPay price tiers"
        else:
            monthly_price = float(tier["monthly_price"])
            tier_desc = f"Tier {tier['tier_no']} ({tier['min_employees']}-{tier['max_employees']} employees)"
        annual_vip = employee_count * monthly_price * 12 * (1 + vat_pct / 100.0)
    else:
        monthly_price = 0.0
        tier_desc = "Custom annual licence price"
        annual_vip = max(0.0, float(custom_license_vip))
    lines = [{"description": f"DatecPay Payroll - Annual Licence ({employee_count} employees; {tier_desc})",
              "qty": 1, "unit_vip": annual_vip, "total_vip": annual_vip, "billing": "Annual recurring",
              "assurance_eligible": True}]
    total = annual_vip
    assurance = annual_vip * assurance_pct / 100.0 if assurance_enabled else 0.0
    if assurance > 0:
        lines.append({"description": "Assurance", "qty": 1,
                      "unit_vip": assurance, "total_vip": assurance, "billing": "Annual recurring",
                      "assurance_eligible": False})
        total += assurance
    if setup_fee > 0:
        lines.append({"description": "Implementation and Setup", "qty": 1,
                      "unit_vip": float(setup_fee), "total_vip": float(setup_fee),
                      "billing": "One-time", "assurance_eligible": False})
        total += float(setup_fee)
    return {"lines": lines, "subtotal": total, "assurance_total": assurance,
            "grand_total": total, "tier": tier, "monthly_price": monthly_price,
            "annual_license_vip": annual_vip, "employees": employee_count}


def quote_number(for_date: Any = None) -> str:
    init_db()
    month_source = for_date if for_date is not None else datetime.now()
    month = month_source.strftime("%Y%m")
    with connect() as conn:
        count = conn.execute("SELECT COUNT(*) AS n FROM quotes WHERE quote_no LIKE ?", (f"QT-{month}-%",)).fetchone()["n"]
    return f"QT-{month}-{count + 1:04d}"


def save_quote_record(quote: dict[str, Any]) -> None:
    with connect() as conn:
        conn.execute("""INSERT INTO quotes
            (quote_no, quote_date, customer, organisation, email, phone, subject, product, variant,
             pricing_mode, subtotal, assurance_total, grand_total, quote_json, created_at, status)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (quote["quote_no"], quote["quote_date"], quote.get("customer", ""), quote.get("organisation", ""),
             quote.get("email", ""), quote.get("phone", ""), quote.get("subject", ""), quote.get("product", ""),
             quote.get("variant", ""), quote.get("pricing_mode", ""), float(quote.get("subtotal", 0)),
             float(quote.get("assurance_total", 0)), float(quote.get("grand_total", 0)),
             json.dumps(quote, ensure_ascii=False), datetime.now().isoformat(timespec="seconds"), "Draft"))
        conn.commit()


def list_quotes(limit: int = 100) -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        return [dict(r) for r in conn.execute("SELECT id,quote_no,quote_date,customer,organisation,product,variant,pricing_mode,grand_total,status,loss_reason,created_at FROM quotes ORDER BY id DESC LIMIT ?", (limit,)).fetchall()]


def update_quote_status(quote_no: str, status: str, loss_reason: str = "") -> bool:
    """Update workflow status and record a reason for declined quotations."""
    allowed = {"Draft", "Quote Sent", "Quote Accepted", "Quote Declined", "Expired", "Cancelled"}
    if status not in allowed:
        raise ValueError(f"Unsupported quote status: {status}")
    if status == "Quote Declined" and not str(loss_reason).strip():
        raise ValueError("Please record a reason for the declined quotation.")
    with connect() as conn:
        row = conn.execute("SELECT quote_json FROM quotes WHERE quote_no=?", (quote_no,)).fetchone()
        if not row:
            return False
        try:
            quote_data = json.loads(row["quote_json"] or "{}")
        except (json.JSONDecodeError, TypeError):
            quote_data = {}
        quote_data["status"] = status
        quote_data["loss_reason"] = str(loss_reason or "")
        conn.execute("UPDATE quotes SET status=?, loss_reason=?, quote_json=? WHERE quote_no=?",
                     (status, str(loss_reason or ""), json.dumps(quote_data, ensure_ascii=False), quote_no))
        conn.commit()
    return True


def update_saved_quote(quote_no: str, quote_data: dict[str, Any]) -> bool:
    """Replace an existing quote snapshot while preserving its internal reference and status."""
    with connect() as conn:
        row = conn.execute("SELECT status, loss_reason FROM quotes WHERE quote_no=?", (quote_no,)).fetchone()
        if not row:
            return False
        quote_data["quote_no"] = quote_no
        quote_data["status"] = row["status"]
        quote_data["loss_reason"] = row["loss_reason"]
        options = quote_data.get("options") or []
        if options:
            quote_data["grand_total"] = float(options[0].get("grand_total", quote_data.get("grand_total", 0)) or 0)
            quote_data["subtotal"] = float(options[0].get("subtotal", quote_data.get("subtotal", 0)) or 0)
        conn.execute("""UPDATE quotes SET quote_date=?, customer=?, organisation=?, email=?, phone=?, subject=?,
            product=?, variant=?, pricing_mode=?, subtotal=?, assurance_total=?, grand_total=?, quote_json=? WHERE quote_no=?""",
            (quote_data.get("quote_date", ""), quote_data.get("customer", ""), quote_data.get("organisation", ""),
             quote_data.get("email", ""), quote_data.get("phone", ""), quote_data.get("subject", ""),
             quote_data.get("product", ""), quote_data.get("variant", ""), quote_data.get("pricing_mode", ""),
             float(quote_data.get("subtotal", 0) or 0), float(quote_data.get("assurance_total", 0) or 0),
             float(quote_data.get("grand_total", 0) or 0), json.dumps(quote_data, ensure_ascii=False), quote_no))
        conn.commit()
    return True

def delete_quote_record(quote_no: str) -> bool:
    with connect() as conn:
        cur = conn.execute("DELETE FROM quotes WHERE quote_no=?", (quote_no,))
        conn.commit()
        return cur.rowcount > 0


def get_quote(quote_id: int) -> dict[str, Any] | None:
    with connect() as conn:
        row = conn.execute("SELECT * FROM quotes WHERE id=?", (quote_id,)).fetchone()
    if not row: return None
    out = dict(row)
    out["quote_data"] = json.loads(out.pop("quote_json"))
    out["quote_data"]["status"] = out.get("status", out["quote_data"].get("status", "Draft"))
    return out


def get_quote_by_no(quote_no: str) -> dict[str, Any] | None:
    with connect() as conn:
        row = conn.execute("SELECT * FROM quotes WHERE quote_no=?", (quote_no,)).fetchone()
    if not row: return None
    out = dict(row)
    out["quote_data"] = json.loads(out.pop("quote_json"))
    out["quote_data"]["status"] = out.get("status", out["quote_data"].get("status", "Draft"))
    return out
