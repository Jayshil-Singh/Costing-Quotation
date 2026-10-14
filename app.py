from __future__ import annotations

import io
import json
import math
import sqlite3
import tempfile
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from pricing_engine import (
    DEFAULT_TERMS,
    calculate_cost_quote,
    calculate_datecpay,
    fx_rates,
    get_quote_by_no,
    get_settings,
    init_db,
    list_catalog,
    list_quotes,
    list_tiers,
    money,
    price_history,
    quote_number,
    save_catalog,
    save_quote_record,
    save_settings,
    save_tiers,
    update_quote_status,
    update_saved_quote,
    delete_quote_record,
    DB_PATH,
)
from quote_pdf import build_quote_pdf
from costing_xlsx import build_costing_xlsx

BASE_DIR = Path(__file__).resolve().parent
CURRENCIES = ["USD", "AUD", "NZD", "FJD"]
PRODUCTS = ["PayGlobal", "DatecPay 2.0", "Epicor"]

st.set_page_config(
    page_title="Datec | Costing & Quotations",
    page_icon=str(BASE_DIR / "assets" / "datec_logo.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      :root {
        --datec-navy: #173453;
        --datec-blue: #315b8a;
        --datec-blue-light: #e8eff7;
        --datec-edge: #c5d0de;
        --datec-ink: #202d3e;
        --datec-muted: #65768a;
        --datec-shadow: rgba(24, 45, 73, .13);
      }
      .stApp {
        background: radial-gradient(ellipse at 12% 0%, #f9fbfd 0%, #edf2f7 48%, #e7edf4 100%);
        color: var(--datec-ink);
      }
      .main .block-container { max-width: 1580px; padding-top: 1.45rem; padding-bottom: 3rem; }
      [data-testid="stSidebar"] {
        background: linear-gradient(165deg, #102641 0%, #203f63 52%, #152d4b 100%);
        border-right: 1px solid #9dacc0;
        box-shadow: inset -2px 0 5px rgba(255,255,255,.08), 3px 0 12px rgba(17,38,63,.12);
      }
      [data-testid="stSidebar"] * { color: #f4f7fb !important; }
      [data-testid="stSidebar"] [data-testid="stRadio"] label { padding: 7px 9px; border-radius: 8px; }
      [data-testid="stSidebar"] [data-testid="stRadio"] label:hover { background: rgba(255,255,255,.09); }
      [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: #d9e4f0 !important; }
      .hero-panel {
        position: relative; overflow: hidden;
        padding: 22px 25px 19px; margin: 0 0 1.2rem;
        background: linear-gradient(180deg, #ffffff 0%, #f4f7fb 54%, #e9eff6 100%);
        border: 1px solid #c6d1df; border-top-color: #ffffff; border-radius: 15px;
        box-shadow: 0 5px 12px rgba(27,49,77,.09), inset 0 1px 0 #ffffff, inset 0 -1px 0 rgba(170,186,204,.45);
      }
      .hero-panel:before { content: ""; position:absolute; left:0; top:0; right:0; height:4px; background:linear-gradient(90deg,#182f4d,#4272a9 55%,#b4c8dc); }
      .brand-kicker { color:#54718f; font-size:.69rem; letter-spacing:.16em; font-weight:800; margin-bottom:.35rem; }
      .brand-title { color: #1b3554; font-size: 1.85rem; line-height: 1.15; font-weight: 800; margin-bottom: .28rem; text-shadow: 0 1px 0 #fff; }
      .brand-sub { color: #65768a; font-size: .94rem; margin-bottom: 0; }
      .section-label { display:inline-block; color:#294a70; font-size:.74rem; font-weight:800; letter-spacing:.085em; text-transform:uppercase; margin:.15rem 0 .55rem; padding:5px 9px; border:1px solid #c9d4e2; border-radius:7px; background:linear-gradient(180deg,#fbfdff,#e9eff6); box-shadow:inset 0 1px 0 #fff, 0 1px 2px rgba(25,47,74,.04); }
      .hint, .small-note { color: #65768a; font-size: .84rem; }
      div[data-testid="stMetric"] {
        background: linear-gradient(180deg,#fff 0%,#f4f7fb 100%); border: 1px solid #cbd5e1;
        padding: 15px 16px; border-radius: 12px;
        box-shadow: 0 3px 6px rgba(21,43,70,.08), inset 0 1px 0 white, inset 0 -1px 0 #e0e6ee;
      }
      div[data-testid="stMetricLabel"] { color: #6c7c90; }
      div[data-testid="stMetricValue"] { color: #173453; text-shadow: 0 1px 0 #fff; }
      div[data-testid="stForm"], div[data-testid="stVerticalBlockBorderWrapper"] {
        background: linear-gradient(180deg, rgba(255,255,255,.98), rgba(243,247,251,.98));
        border: 1px solid #cad5e2; border-top-color:#fff; border-radius: 13px;
        box-shadow: 0 3px 9px rgba(25,47,75,.055), inset 0 1px 0 rgba(255,255,255,.95);
      }
      div[data-testid="stForm"] { padding: 15px 17px; }
      div[data-testid="stExpander"] { background:linear-gradient(180deg,#fff,#f2f6fa); border-radius:10px; border:1px solid #cbd6e3; box-shadow:0 2px 5px rgba(25,47,75,.05),inset 0 1px 0 white; }
      div[data-testid="stDataFrame"], div[data-testid="stTable"] { border:1px solid #cbd5e1; border-radius:9px; box-shadow:inset 0 1px 2px rgba(21,43,70,.04); overflow:hidden; }
      hr { border-color: #d3dce7; margin-top: 1.1rem; margin-bottom: 1.1rem; }
      [data-baseweb="input"] > div, [data-baseweb="textarea"] > div, [data-baseweb="select"] > div {
        border-radius:8px !important; border-color:#bac7d7 !important;
        background:linear-gradient(180deg,#fdfefe,#f4f7fa) !important;
        box-shadow:inset 0 1px 3px rgba(25,47,75,.09), 0 1px 0 rgba(255,255,255,.9) !important;
      }
      [data-baseweb="input"] input, [data-baseweb="textarea"] textarea { color:#1e334c !important; }
      [data-baseweb="input"] > div:focus-within, [data-baseweb="textarea"] > div:focus-within, [data-baseweb="select"] > div:focus-within {
        border-color:#557ca7 !important; box-shadow:inset 0 1px 2px rgba(25,47,75,.06),0 0 0 2px rgba(69,111,158,.14) !important;
      }
      .stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {
        border-radius:9px; border:1px solid #aebdd0; color:#213c5a;
        background:linear-gradient(180deg,#ffffff 0%,#edf2f7 48%,#dfe7f0 100%);
        font-weight:700; text-shadow:0 1px 0 #fff;
        box-shadow:0 2px 3px rgba(24,45,73,.12), inset 0 1px 0 #fff, inset 0 -1px 0 rgba(163,180,199,.45);
        transition:transform .08s ease, box-shadow .08s ease, filter .08s ease;
      }
      .stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
        border-color:#8da4bf; filter:brightness(1.015); box-shadow:0 3px 6px rgba(24,45,73,.16),inset 0 1px 0 #fff;
      }
      .stButton > button:active, .stDownloadButton > button:active { transform:translateY(1px); box-shadow:inset 0 2px 4px rgba(24,45,73,.18); }
      .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"], [data-testid="stFormSubmitButton"] > button[kind="primary"] {
        color:#fff; text-shadow:0 -1px 0 rgba(0,0,0,.2); border:1px solid #193c63;
        background:linear-gradient(180deg,#426d9c 0%,#2b527d 48%,#1d3e64 100%);
        box-shadow:0 3px 5px rgba(21,43,70,.22),inset 0 1px 0 rgba(255,255,255,.23),inset 0 -1px 0 rgba(0,0,0,.16);
      }
      [data-testid="stAlert"] { border-radius:10px; box-shadow:0 2px 5px rgba(24,45,73,.04); }
      .stCaption, [data-testid="stCaptionContainer"] { color:#6a7a8e; }
      @media (max-width: 760px) { .brand-title { font-size:1.45rem; } .hero-panel { padding:18px 16px 15px; } }
    </style>
    """,
    unsafe_allow_html=True,
)

init_db()
settings = get_settings()

# Clear quote-specific editor and assumption widget state after a quote is saved so
# the next quote starts from the master defaults rather than inheriting manual edits.
if st.session_state.pop("_reset_quote_inputs", False):
    _quote_widget_prefixes = ("datecpay_", "assumptions_", "assurance_", "fixed_editor_", "cost_editor_", "custom_editor_")
    st.session_state.pop("q_amount_display", None)
    for _key in list(st.session_state.keys()):
        if _key.startswith(_quote_widget_prefixes) or (_key.startswith("option_") and _key != "quote_option_count"):
            del st.session_state[_key]


def nfloat(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or pd.isna(value) or value == "":
            return default
        out = float(value)
        return out if math.isfinite(out) else default
    except (ValueError, TypeError):
        return default


def fmt_amount(value: Any, display: str | None = None) -> str:
    amount = nfloat(value)
    mode = display or str(settings.get("amount_display", "whole"))
    if mode == "whole":
        amount = float(Decimal(str(amount)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        return f"{amount:,.0f}"
    return f"{amount:,.2f}"


def fmt_fjd(value: Any) -> str:
    return f"FJD {fmt_amount(value)}"


def text_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none", "<na>"} else text


def database_backup_bytes() -> bytes:
    """Create a consistent SQLite backup, including committed settings and quote data."""
    with tempfile.TemporaryDirectory(prefix="datec_quote_backup_") as temp_dir:
        backup_path = Path(temp_dir) / "costing_quotes_backup.db"
        source = sqlite3.connect(str(DB_PATH))
        target = sqlite3.connect(str(backup_path))
        try:
            source.backup(target)
        finally:
            target.close()
            source.close()
        return backup_path.read_bytes()


def title_block(title: str, subtitle: str = "") -> None:
    subtitle_html = f'<div class="brand-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="hero-panel"><div class="brand-kicker">DATEC  /  COSTING &amp; QUOTATIONS</div>'
        f'<div class="brand-title">{title}</div>{subtitle_html}</div>',
        unsafe_allow_html=True,
    )


def quote_meta_inputs() -> dict[str, Any]:
    today = date.today()
    left, right = st.columns([1.25, 1])
    with left:
        st.markdown('<div class="section-label">Customer contact</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        customer = c1.text_input("TO — Contact person *", key="q_customer", placeholder="e.g. Ajesh Dutt")
        organisation = c2.text_input("Organisation *", key="q_org", placeholder="Customer company name")
        c3, c4 = st.columns(2)
        email = c3.text_input("Customer email", key="q_email", placeholder="name@customer.com")
        phone = c4.text_input("Customer phone", key="q_phone", placeholder="+679 ...")
    with right:
        st.markdown('<div class="section-label">Quotation &amp; consultant details</div>', unsafe_allow_html=True)
        quote_date = st.date_input("Quotation date", value=today, key="q_date")
        amount_display = st.selectbox(
            "Amount display for this quotation", ["whole", "decimal"],
            index=0 if str(settings.get("amount_display", "whole")) == "whole" else 1,
            format_func=lambda value: "Whole FJD (rounded)" if value == "whole" else "Two decimal places",
            help="Choose whole FJD amounts or retain cents on this quotation. This is a display choice; costing formulas keep their precision.",
            key="q_amount_display",
        )
        prepared_by = st.text_input("FROM — Prepared by", value=str(settings.get("prepared_by_default") or ""), key="q_prepared")
        subject = st.text_input("Subject", key="q_subject", placeholder="Product / project quotation")
        c3, c4 = st.columns(2)
        prepared_phone = c3.text_input(
            "Consultant phone", value=str(settings.get("prepared_by_phone_default") or ""),
            key="q_prepared_phone", placeholder="Direct number for preparer",
        )
        prepared_email = c4.text_input(
            "Consultant email", value=str(settings.get("prepared_by_email_default") or ""),
            key="q_prepared_email", placeholder="preparer@datec.com.fj",
        )
    return {
        "customer": customer.strip(), "organisation": organisation.strip(),
        "email": email.strip(), "phone": phone.strip(), "quote_date_obj": quote_date,
        "prepared_by": prepared_by.strip(), "prepared_phone": prepared_phone.strip(),
        "prepared_email": prepared_email.strip(), "subject": subject.strip(),
        "amount_display": amount_display,
    }


def costing_assumption_inputs(product: str, variant: str, key_prefix: str) -> dict[str, Any]:
    """Per-quote assumptions. PayGlobal inputs mirror the factors used in the source Excel model."""
    base_rates = fx_rates(settings)
    is_payglobal = product == "PayGlobal"
    if is_payglobal:
        vat_factor_default = 1.125
        duty_factor_default = 1.0
        margin_factor_default = 1.2
        platform_margin_default = 1.2
        if variant == "Hybrid":
            margin_factor_default = 1.0  # Local licence margin factor in the supplied Hybrid workbook.
        rate_default_currency = "FJD" if variant == "On-Premise" else "NZD"
        st.caption("PayGlobal uses the same calculation factors as the supplied workbook: landed cost = unit buy × discount/freight/duty factors ÷ exchange rate; VEP = landed cost × margin factor; VIP = VEP × VAT factor.")
        a, b, c = st.columns(3)
        vat_factor = a.number_input("VAT factor", min_value=1.0, max_value=3.0, value=vat_factor_default, step=0.005, format="%.3f", key=f"{key_prefix}_vat_factor")
        duty_factor = c.number_input("Duty factor", min_value=1.0, max_value=5.0, value=duty_factor_default, step=0.01, format="%.3f", key=f"{key_prefix}_duty_factor")
        if variant == "Hybrid":
            margin_factor = b.number_input("Licence margin factor", min_value=1.0, max_value=10.0, value=margin_factor_default, step=0.01, format="%.3f", key=f"{key_prefix}_margin_factor_license")
            platform_margin_factor = st.number_input("Subscription / platform margin factor", min_value=1.0, max_value=10.0, value=platform_margin_default, step=0.01, format="%.3f", key=f"{key_prefix}_margin_factor_platform")
        else:
            margin_factor = b.number_input("Margin factor", min_value=1.0, max_value=10.0, value=margin_factor_default, step=0.01, format="%.3f", key=f"{key_prefix}_margin_factor")
            platform_margin_factor = margin_factor
        st.caption("Currency rates are foreign-currency units per FJD. For example, NZD 0.745 means FJD 1 = NZD 0.745. Hybrid applies FJD 1.000000 to local licence costs and the selected foreign rate to supplier subscription / platform costs.")
        f1, f2, f3, f4 = st.columns(4)
        fx_usd = f1.number_input("USD rate", min_value=0.0001, max_value=1000.0, value=float(base_rates.get("USD", 0.435)), step=0.0001, format="%.6f", key=f"{key_prefix}_fx_usd")
        fx_aud = f2.number_input("AUD rate", min_value=0.0001, max_value=1000.0, value=float(base_rates.get("AUD", 1.0)), step=0.0001, format="%.6f", key=f"{key_prefix}_fx_aud")
        fx_nzd = f3.number_input("NZD rate", min_value=0.0001, max_value=1000.0, value=float(base_rates.get("NZD", 0.745)), step=0.0001, format="%.6f", key=f"{key_prefix}_fx_nzd")
        pricing_currency = f4.selectbox("Default buy currency", CURRENCIES, index=CURRENCIES.index(rate_default_currency), key=f"{key_prefix}_buy_currency")
        quote_fx = {"USD": float(fx_usd), "AUD": float(fx_aud), "NZD": float(fx_nzd), "FJD": 1.0}
        if abs(float(fx_aud) - 1.0) < 1e-9:
            st.warning("The AUD rate is still at the initial placeholder value of 1.000000. Confirm the current AUD/FJD rate before quoting any AUD-denominated buy price.")
        return {
            "vat_pct": (float(vat_factor) - 1.0) * 100.0,
            "margin_pct": (float(platform_margin_factor if variant == "Hybrid" else margin_factor) - 1.0) * 100.0,
            "duty_pct": (float(duty_factor) - 1.0) * 100.0,
            "vat_factor": float(vat_factor), "duty_factor": float(duty_factor),
            "margin_factor": float(margin_factor), "platform_margin_factor": float(platform_margin_factor),
            "fx_rates": quote_fx, "selected_currency": pricing_currency,
            "selected_fx_rate": quote_fx.get(pricing_currency, 1.0),
        }

    if product == "Epicor":
        default_margin = float(settings.get("default_epicor_margin_pct", 40.0))
    else:
        default_margin = float(settings.get("default_margin_pct", 20.0))
    default_duty = float(settings.get("default_duty_pct", 18.0))
    default_vat = float(settings.get("default_vat_pct", 12.5))
    p1, p2, p3 = st.columns(3)
    vat_pct = p1.number_input("VAT (%)", min_value=0.0, max_value=100.0, value=default_vat, step=0.5, key=f"{key_prefix}_vat")
    margin_pct = p2.number_input("Margin (%)", min_value=0.0, max_value=500.0, value=default_margin, step=1.0, key=f"{key_prefix}_margin")
    duty_pct = p3.number_input("Duty (%)", min_value=0.0, max_value=300.0, value=default_duty, step=1.0, key=f"{key_prefix}_duty")
    st.caption("Customer quotations are generated in FJD VIP. Currency rates are foreign-currency units per FJD. These rates apply to this quote only.")
    f1, f2, f3, f4 = st.columns(4)
    fx_usd = f1.number_input("USD rate", min_value=0.0001, max_value=1000.0, value=float(base_rates.get("USD", 0.435)), format="%.6f", key=f"{key_prefix}_fx_usd")
    fx_aud = f2.number_input("AUD rate", min_value=0.0001, max_value=1000.0, value=float(base_rates.get("AUD", 1.0)), format="%.6f", key=f"{key_prefix}_fx_aud")
    fx_nzd = f3.number_input("NZD rate", min_value=0.0001, max_value=1000.0, value=float(base_rates.get("NZD", 0.745)), format="%.6f", key=f"{key_prefix}_fx_nzd")
    default_currency = "USD" if product == "Epicor" else "FJD"
    pricing_currency = f4.selectbox("Default buy currency", CURRENCIES, index=CURRENCIES.index(default_currency), key=f"{key_prefix}_buy_currency")
    if abs(float(fx_aud) - 1.0) < 1e-9:
        st.warning("The AUD rate is still at the initial placeholder value of 1.000000. Confirm the current AUD/FJD rate before quoting any AUD-denominated buy price.")
    quote_fx = {"USD": float(fx_usd), "AUD": float(fx_aud), "NZD": float(fx_nzd), "FJD": 1.0}
    return {"vat_pct": float(vat_pct), "margin_pct": float(margin_pct), "duty_pct": float(duty_pct),
            "fx_rates": quote_fx, "selected_currency": pricing_currency,
            "selected_fx_rate": quote_fx.get(pricing_currency, 1.0)}


def _billing_periods(billing: str) -> int:
    text = str(billing or "One-time").lower()
    if text.startswith("month"):
        return 12
    if text.startswith("quarter"):
        return 4
    return 1


def _to_quote_line(description: str, qty: float, unit_price: float, billing: str,
                   pricing_method: str) -> dict[str, Any]:
    period_count = _billing_periods(billing)
    clean_billing = str(billing or "One-time")
    if clean_billing.lower().startswith("month"):
        return {"description": description, "qty": 12, "unit_vip": qty * unit_price,
                "total_vip": qty * unit_price * 12, "billing": "Annual recurring",
                "source_qty": qty, "source_unit_vip": unit_price, "pricing_method": pricing_method}
    if clean_billing.lower().startswith("quarter"):
        return {"description": description, "qty": 4, "unit_vip": qty * unit_price,
                "total_vip": qty * unit_price * 4, "billing": "Quarterly recurring",
                "source_qty": qty, "source_unit_vip": unit_price, "pricing_method": pricing_method}
    output_billing = "Annual recurring" if clean_billing.lower().startswith("annual") else "One-time"
    return {"description": description, "qty": qty, "unit_vip": unit_price,
            "total_vip": qty * unit_price, "billing": output_billing,
            "source_qty": qty, "source_unit_vip": unit_price, "pricing_method": pricing_method}


def data_editor_fixed(catalog_rows: list[dict[str, Any]], key: str) -> pd.DataFrame:
    rows = [{
        "Include": nfloat(row.get("default_qty"), 1) > 0,
        "Description": row.get("item_name", ""),
        "Quantity": nfloat(row.get("default_qty"), 1),
        "Unit VIP FJD": nfloat(row.get("fixed_unit_vip")),
        "Billing": row.get("billing_frequency") or "One-time",
    } for row in catalog_rows]
    if not rows:
        rows = [{"Include": True, "Description": "", "Quantity": 1.0, "Unit VIP FJD": 0.0, "Billing": "One-time"}]
        st.info("No standard fixed unit prices have been configured for this option. Enter an approved fixed unit price below, or create a master fixed price in Price book.")
    st.caption("These are unit selling prices, not extended totals. Monthly rates are multiplied by the entered quantity and 12 periods. Quote edits do not change the master price book.")
    return st.data_editor(
        pd.DataFrame(rows), key=key, use_container_width=True, hide_index=True, num_rows="dynamic",
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", default=True, width="small"),
            "Description": st.column_config.TextColumn("Description", required=True, width="large"),
            "Quantity": st.column_config.NumberColumn("Qty / units", min_value=0.0, step=1.0, format="%.2f"),
            "Unit VIP FJD": st.column_config.NumberColumn("Unit VIP FJD", min_value=0.0, step=0.01, format="%.2f"),
            "Billing": st.column_config.SelectboxColumn("Billing", options=["One-time", "Annual recurring", "Monthly", "Quarterly recurring"], required=True),
        },
    )


def data_editor_cost(catalog_rows: list[dict[str, Any]], default_currency: str, key: str) -> pd.DataFrame:
    rows = []
    for row in catalog_rows:
        rows.append({
            "Include": nfloat(row.get("default_qty"), 1) > 0,
            "Description": row.get("item_name", ""),
            "Quantity": nfloat(row.get("default_qty"), 1),
            "List Price": nfloat(row.get("list_price") or row.get("buy_price")),
            "Buy price / unit": nfloat(row.get("buy_price")),
            "Buy currency": row.get("buy_currency") or default_currency,
            "Discount %": nfloat(row.get("discount_pct")),
            "Freight %": nfloat(row.get("freight_pct")),
            "Quote unit VIP override": None,
            "Billing": row.get("billing_frequency") or "One-time",
            "Assurance eligible": bool(row.get("assurance_eligible", False)),
            "Source notes": row.get("notes", ""),
        })
    if not rows:
        rows = [{"Include": True, "Description": "Custom item", "Quantity": 1.0, "List Price": 0.0, "Buy price / unit": 0.0,
                 "Buy currency": default_currency, "Discount %": 0.0, "Freight %": 0.0,
                 "Quote unit VIP override": None, "Billing": "One-time", "Assurance eligible": False,
                 "Source notes": ""}]
    st.caption("Cost-based lines start from the unit buy price only. VAT, margin, duty, currency and extended totals are recalculated. A quote unit VIP override is optional and applies to this quote only.")
    return st.data_editor(
        pd.DataFrame(rows), key=key, use_container_width=True, hide_index=True, num_rows="dynamic",
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", default=True, width="small"),
            "Description": st.column_config.TextColumn("Description", required=True, width="large"),
            "Quantity": st.column_config.NumberColumn("Qty / units", min_value=0.0, step=1.0, format="%.2f"),
            "List Price": st.column_config.NumberColumn("List Price", min_value=0.0, step=0.01, format="%.2f", help="Source workbook list price; display/reference value. Cent-level editing is enabled."),
            "Buy price / unit": st.column_config.NumberColumn("Datec Unit Buy / unit", min_value=0.0, step=0.01, format="%.2f", help="Decimal supplier unit buy price. Cent-level editing is enabled; calculations use this unit price, not previous extended totals."),
            "Buy currency": st.column_config.SelectboxColumn("Buy currency", options=CURRENCIES, required=True, width="small"),
            "Discount %": st.column_config.NumberColumn("Discount %", min_value=0.0, max_value=100.0, step=1.0, format="%.2f"),
            "Freight %": st.column_config.NumberColumn("Freight %", min_value=0.0, max_value=300.0, step=1.0, format="%.2f"),
            "Quote unit VIP override": st.column_config.NumberColumn("Quote unit VIP override", min_value=0.0, step=0.01, format="%.2f", help="Approved final VIP price per input unit. It is only saved to this quote."),
            "Billing": st.column_config.SelectboxColumn("Billing", options=["One-time", "Monthly", "Annual recurring", "Quarterly recurring"], required=True),
            "Assurance eligible": st.column_config.CheckboxColumn("Assurance eligible", default=False),
            "Source notes": st.column_config.TextColumn("Source notes", disabled=True, width="large"),
        },
    )



PAYGLOBAL_SERVICE_MODULES = ["Payroll", "ESS", "HR", "T&A", "Rostering"]
PAYGLOBAL_SERVICE_ACTIVITIES = [
    "Setup / Configuration", "Data Migration", "Integration / Technical",
    "Testing & UAT", "Training", "Go-Live / Implementation",
]


def payglobal_services_editor(key: str, default_rate: float = 165.0) -> tuple[pd.DataFrame, pd.DataFrame, float]:
    """Separate implementation-services worksheet inputs, mirroring ServicesCost in Excel."""
    st.markdown("#### Implementation services costing (separate calculation)")
    st.caption("Enter hours by module and activity. Services are calculated separately at the selected FJD hourly rate and added to the overall quotation. The default source rate is FJD 165.00 per hour.")
    rows = []
    for module in PAYGLOBAL_SERVICE_MODULES:
        row = {"Module": module}
        row.update({f"{act} hrs": 0.0 for act in PAYGLOBAL_SERVICE_ACTIVITIES})
        row["Service rate FJD/hr"] = float(default_rate)
        row["Discount FJD"] = 0.0
        rows.append(row)
    editable_cols = ["Module"] + [f"{act} hrs" for act in PAYGLOBAL_SERVICE_ACTIVITIES] + ["Service rate FJD/hr", "Discount FJD"]
    services_df = st.data_editor(
        pd.DataFrame(rows, columns=editable_cols), key=key, use_container_width=True,
        hide_index=True, num_rows="fixed",
        column_config={
            "Module": st.column_config.TextColumn("Module", disabled=True, width="small"),
            **{f"{act} hrs": st.column_config.NumberColumn(f"{act} hrs", min_value=0.0, step=0.5, format="%.2f", width="medium") for act in PAYGLOBAL_SERVICE_ACTIVITIES},
            "Service rate FJD/hr": st.column_config.NumberColumn("Service rate FJD/hr", min_value=0.0, step=0.01, format="%.2f"),
            "Discount FJD": st.column_config.NumberColumn("Discount FJD", min_value=0.0, step=0.01, format="%.2f"),
        },
    )
    summary_rows = []
    grand_total = 0.0
    for _, row in services_df.iterrows():
        hours = sum(max(0.0, nfloat(row.get(f"{act} hrs"))) for act in PAYGLOBAL_SERVICE_ACTIVITIES)
        rate = max(0.0, nfloat(row.get("Service rate FJD/hr"), default_rate))
        gross = hours * rate
        discount = min(gross, max(0.0, nfloat(row.get("Discount FJD"))))
        total = max(0.0, gross - discount)
        grand_total += total
        summary_rows.append({"Module": text_value(row.get("Module")), "Total Hours": hours,
                             "Service Rate (FJD)": rate, "Gross Services (FJD)": gross,
                             "Discount (FJD)": discount, "Total Services Cost (FJD)": total})
    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True,
                 column_config={c: st.column_config.NumberColumn(c, format="%.0f" if c == "Total Services Cost (FJD)" else "%.2f") for c in summary_df.columns if c != "Module"})
    st.metric("Total Services Cost", f"FJD {grand_total:,.0f}")
    return services_df, summary_df, grand_total


def payglobal_costing_preview(df: pd.DataFrame, assumptions: dict[str, Any], variant: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build the live full and simplified PayGlobal costing views from input unit prices only."""
    rows = cost_editor_to_rows(df)
    out = []
    all_rates = assumptions.get("fx_rates") or {"USD": 0.435, "AUD": 1.0, "NZD": 0.745, "FJD": 1.0}
    margin_factor = nfloat(assumptions.get("margin_factor"), 1.2)
    platform_margin_factor = nfloat(assumptions.get("platform_margin_factor"), margin_factor)
    vat_factor = nfloat(assumptions.get("vat_factor"), 1.125)
    duty_factor = nfloat(assumptions.get("duty_factor"), 1.0)
    for idx, row in enumerate(rows, start=1):
        qty = max(0.0, nfloat(row.get("qty")))
        if qty <= 0:
            continue
        currency = str(row.get("buy_currency") or "FJD").upper()
        fx = 1.0 if currency == "FJD" else nfloat(all_rates.get(currency), 1.0)
        if fx <= 0:
            fx = 1.0
        buy_price = max(0.0, nfloat(row.get("buy_price")))
        discount_pct = min(100.0, max(0.0, nfloat(row.get("discount_pct"))))
        freight_pct = max(0.0, nfloat(row.get("freight_pct")))
        net_unit = buy_price * (1 - discount_pct / 100.0)
        buy_total = net_unit * qty
        landed_ex = net_unit * (1 + freight_pct / 100.0) * duty_factor
        unit_landed_fjd = landed_ex / fx
        total_landed_fjd = unit_landed_fjd * qty
        if variant == "Hybrid":
            row_margin_factor = margin_factor if currency == "FJD" else platform_margin_factor
        else:
            row_margin_factor = margin_factor
        unit_margin = unit_landed_fjd * (row_margin_factor - 1.0)
        total_margin = unit_margin * qty
        unit_vep = unit_landed_fjd * row_margin_factor
        total_vep = unit_vep * qty
        unit_vip = math.ceil(max(0.0, unit_vep * vat_factor) * 100 - 1e-9) / 100.0
        total_vip = unit_vip * qty
        billing = str(row.get("billing_frequency") or "One-time")
        monthly_total = total_vip if billing.lower().startswith("month") else 0.0
        if billing.lower().startswith("month"):
            yearly_total = monthly_total * 12
        else:
            # Year-1 total includes both recurring subscription lines and one-time setup lines.
            # Monthly column remains zero for anything not billed monthly.
            yearly_total = total_vip
        override = row.get("quote_unit_override")
        if override is not None and not pd.isna(override) and str(override).strip() != "":
            unit_vip = max(0.0, nfloat(override))
            total_vip = unit_vip * qty
            if billing.lower().startswith("month"):
                monthly_total = total_vip
                yearly_total = total_vip * 12
            else:
                yearly_total = total_vip
        out.append({
            "No.": idx, "Description Of Items": row.get("description", ""), "Vendor": currency,
            "Quantity": qty, "List Price": nfloat(row.get("list_price") or buy_price), "Buy Currency": currency,
            "Discount %": discount_pct, "Datec Total Buy": buy_total,
            "Freight %": freight_pct, "Duty factor": duty_factor, "L/C Ex": landed_ex,
            "Rate": fx, "Unit L/C - FJD": unit_landed_fjd, "Total L/C - FJD": total_landed_fjd,
            "Margin factor": row_margin_factor, "Unit Margin Value": unit_margin,
            "Total Margin Value": total_margin, "Unit VEP FJD": unit_vep,
            "Total VEP FJD": total_vep, "VAT factor": vat_factor,
            "Unit VIP FJD": unit_vip, "Total VIP FJD (Monthly)": monthly_total,
            "Total VIP FJD (Yearly)": yearly_total, "Billing": billing,
        })
    full_df = pd.DataFrame(out)
    if full_df.empty:
        return full_df, full_df
    if variant == "Cloud":
        summary_cols = ["No.", "Description Of Items", "Quantity", "List Price", "Datec Total Buy", "Total L/C - FJD", "Total VIP FJD (Monthly)", "Total VIP FJD (Yearly)"]
    elif variant == "On-Premise":
        summary_cols = ["No.", "Description Of Items", "Quantity", "List Price", "Datec Total Buy", "Total L/C - FJD", "Total VIP FJD (Yearly)"]
    else:
        summary_cols = ["No.", "Description Of Items", "Quantity", "List Price", "Datec Total Buy", "Total L/C - FJD", "Total VIP FJD (Yearly)"]
    return full_df, full_df[summary_cols]

def data_editor_custom(key: str) -> pd.DataFrame:
    rows = [{"Include": True, "Description": "", "Quantity": 1.0, "Unit VIP FJD": 0.0, "Billing": "One-time"}]
    st.caption("Enter final VAT-inclusive VIP unit prices. Add as many lines as needed; monthly rates are extended to 12 months in the quote.")
    return st.data_editor(
        pd.DataFrame(rows), key=key, use_container_width=True, hide_index=True, num_rows="dynamic",
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", default=True, width="small"),
            "Description": st.column_config.TextColumn("Description", required=True, width="large"),
            "Quantity": st.column_config.NumberColumn("Qty / units", min_value=0.0, step=1.0, format="%.2f"),
            "Unit VIP FJD": st.column_config.NumberColumn("Unit VIP FJD", min_value=0.0, step=0.01, format="%.2f"),
            "Billing": st.column_config.SelectboxColumn("Billing", options=["One-time", "Annual recurring", "Monthly", "Quarterly recurring"], required=True),
        },
    )


def fixed_editor_to_lines(df: pd.DataFrame) -> list[dict[str, Any]]:
    lines = []
    for _, row in df.iterrows():
        if not bool(row.get("Include", False)):
            continue
        desc = text_value(row.get("Description"))
        qty = max(0.0, nfloat(row.get("Quantity")))
        unit = max(0.0, nfloat(row.get("Unit VIP FJD")))
        if not desc or qty <= 0 or unit <= 0:
            continue
        lines.append(_to_quote_line(desc, qty, unit, str(row.get("Billing") or "One-time"), "Fixed unit price / quote-only override"))
    return lines


def custom_editor_to_lines(df: pd.DataFrame) -> list[dict[str, Any]]:
    lines = []
    for _, row in df.iterrows():
        if not bool(row.get("Include", False)):
            continue
        desc = text_value(row.get("Description"))
        qty = max(0.0, nfloat(row.get("Quantity")))
        unit = max(0.0, nfloat(row.get("Unit VIP FJD")))
        if not desc or qty <= 0 or unit <= 0:
            continue
        lines.append(_to_quote_line(desc, qty, unit, str(row.get("Billing") or "One-time"), "Custom unit VIP price"))
    return lines


def fixed_editor_to_costing_rows(df: pd.DataFrame, *, row_type: str = "fixed") -> list[dict[str, Any]]:
    out = []
    for _, row in df.iterrows():
        if not bool(row.get("Include", False)):
            continue
        desc = text_value(row.get("Description"))
        qty = max(0.0, nfloat(row.get("Quantity")))
        unit = max(0.0, nfloat(row.get("Unit VIP FJD")))
        if not desc or qty <= 0 or unit <= 0:
            continue
        out.append({"description": desc, "qty": qty, "buy_price": 0.0, "buy_currency": "FJD",
                    "discount_pct": 0.0, "freight_pct": 0.0, "quote_unit_override": unit,
                    "billing_frequency": str(row.get("Billing") or "One-time"), "assurance_eligible": False,
                    "row_type": row_type, "price_basis": "Fixed / quote-only VIP unit price"})
    return out


def custom_editor_to_costing_rows(df: pd.DataFrame) -> list[dict[str, Any]]:
    rows = fixed_editor_to_costing_rows(df, row_type="custom")
    for row in rows:
        row["price_basis"] = "Custom quote-only VIP unit price"
    return rows


def cost_editor_to_rows(df: pd.DataFrame) -> list[dict[str, Any]]:
    rows = []
    for _, row in df.iterrows():
        if not bool(row.get("Include", False)):
            continue
        name = text_value(row.get("Description"))
        qty = max(0.0, nfloat(row.get("Quantity")))
        if not name or qty <= 0:
            continue
        override = row.get("Quote unit VIP override")
        quote_override = None if override is None or pd.isna(override) or str(override).strip() == "" else max(0.0, nfloat(override))
        rows.append({
            "item_name": name, "description": name, "default_qty": qty, "qty": qty,
            "list_price": max(0.0, nfloat(row.get("List Price"), nfloat(row.get("Buy price / unit")))),
            "buy_price": max(0.0, nfloat(row.get("Buy price / unit"))),
            "buy_currency": str(row.get("Buy currency") or "FJD").upper(),
            "discount_pct": min(100.0, max(0.0, nfloat(row.get("Discount %")))),
            "freight_pct": max(0.0, nfloat(row.get("Freight %"))),
            "quote_unit_override": quote_override,
            "fixed_unit_vip": 0.0,
            "billing_frequency": str(row.get("Billing") or "One-time"),
            "assurance_eligible": bool(row.get("Assurance eligible", False)),
            "row_type": "cost",
            "price_basis": text_value(row.get("Source notes")) or "Supplier unit buy price; recalculated by this app",
        })
    return rows


def _option_filename_text(value: str) -> str:
    import re
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", str(value or "").strip()).strip("_")
    return cleaned[:60] or "Customer"


def add_metadata_and_save(quote: dict[str, Any], meta: dict[str, Any], validity_days: int = 0) -> tuple[dict[str, Any], bytes]:
    quote_date = meta["quote_date_obj"]
    options = quote.get("options") or []
    if options:
        products = list(dict.fromkeys(str(opt.get("product", "")) for opt in options if opt.get("product")))
        quote["product"] = " / ".join(products) if products else "Software quotation"
        quote["variant"] = "Multiple options" if len(options) > 1 else str(options[0].get("variant", ""))
        quote["pricing_mode"] = "Multiple options" if len(options) > 1 else str(options[0].get("pricing_mode", ""))
        # The database headline value uses Option 1 only; alternatives must not be added together.
        quote["subtotal"] = float(options[0].get("subtotal", 0) or 0)
        quote["grand_total"] = float(options[0].get("grand_total", 0) or 0)
        quote["assurance_total"] = float(options[0].get("assurance_total", 0) or 0)
        quote["lines"] = options[0].get("lines", [])
    quote.update({
        "quote_no": quote_number(quote_date),  # Internal app reference only; never printed on the customer PDF.
        "quote_date": quote_date.strftime("%d %B %Y"),
        "customer": meta["customer"], "organisation": meta["organisation"],
        "email": meta["email"], "phone": meta["phone"],
        "prepared_by": meta["prepared_by"] or settings.get("prepared_by_default", ""),
        "prepared_phone": meta.get("prepared_phone", "") or settings.get("prepared_by_phone_default", ""),
        "prepared_email": meta.get("prepared_email", "") or settings.get("prepared_by_email_default", ""),
        "greeting": f"Dear {meta['customer'].strip()}," if meta.get("customer", "").strip() else "Dear Valued Customer,",
        "subject": meta["subject"] or f"{quote.get('product', '')} quotation",
        "company_name": settings.get("company_name", "Datec (Fiji) PTE Ltd"),
        "company_address": settings.get("company_address", ""),
        "company_phone": settings.get("company_phone", ""),
        "company_email": settings.get("company_email", ""),
        "branch_address": settings.get("branch_address", ""),
        "branch_phone": settings.get("branch_phone", ""),
        "terms": settings.get("default_terms", DEFAULT_TERMS),
        "amount_display": meta.get("amount_display") or settings.get("amount_display", "whole"),
        "status": "Draft",
    })
    for option in quote.get("options") or []:
        option["amount_display"] = quote["amount_display"]
    quote["pricing_snapshot"] = quote.get("pricing_snapshot", {})
    save_quote_record(quote)
    return quote, build_quote_pdf(quote)


def _preview_pdf(pdf_bytes: bytes, height: int = 780) -> None:
    """Render PDF pages as images instead of embedding a data: PDF iframe.

    Chrome and some managed-browser policies block embedded data:application/pdf
    URLs. Raster previews work consistently in Streamlit while the original PDF
    remains available from the adjacent download button.
    """
    try:
        import fitz  # PyMuPDF
        from PIL import Image
        with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
            if document.page_count == 0:
                st.warning("The quotation PDF contains no pages to preview.")
                return
            st.caption(f"Preview · {document.page_count} page(s). Download the PDF below for the original printable file.")
            # Moderate resolution keeps multi-page quotations responsive without
            # relying on Chrome's embedded PDF viewer or data URL support.
            matrix = fitz.Matrix(1.35, 1.35)
            for page_index, page in enumerate(document, start=1):
                pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                st.image(image, caption=f"Page {page_index} of {document.page_count}", use_container_width=True)
    except ModuleNotFoundError as exc:
        if exc.name == "fitz":
            st.error("PDF preview support is not installed in this app environment.")
            st.info("Close the app and run start_app.bat again. Version 2.0 forces a dependency check and installs PyMuPDF, which provides the fitz module.")
        else:
            st.error("The PDF was generated, but its on-page preview could not be rendered.")
            st.caption(f"You can still use the PDF download button below. Preview detail: {exc}")
    except Exception as exc:
        st.error("The PDF was generated, but its on-page preview could not be rendered.")
        st.caption(f"You can still use the PDF download button below. Preview detail: {exc}")

def _customise_costing_workbook(workbook_bytes: bytes, selected_sheets: list[str], column_choices: dict[str, list[str]], row_mode: str, row_choices: dict[str, list[int]]) -> bytes:
    """Hide unselected sheets/columns/rows while retaining source formulas and references."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    src = io.BytesIO(workbook_bytes)
    wb = load_workbook(src)
    selected = set(selected_sheets)
    for ws in wb.worksheets:
        if ws.title not in selected:
            ws.sheet_state = "hidden"
            continue
        ws.sheet_state = "visible"
        chosen = set(column_choices.get(ws.title, []))
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].hidden = letter not in chosen
        for r in range(1, ws.max_row + 1):
            populated = any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1))
            if row_mode == "Populated rows only" and not populated:
                ws.row_dimensions[r].hidden = True
            elif row_mode == "Choose rows individually" and r not in set(row_choices.get(ws.title, [])):
                ws.row_dimensions[r].hidden = True
    out = io.BytesIO(); wb.save(out); return out.getvalue()

def _costing_export_controls(quote: dict[str, Any], key_prefix: str) -> bytes:
    raw = build_costing_xlsx(quote)
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    wb = load_workbook(io.BytesIO(raw), read_only=False, data_only=False)
    sheet_names = wb.sheetnames
    selected_sheets = st.multiselect("Worksheets to include", sheet_names, default=sheet_names, key=f"{key_prefix}_sheets")
    if not selected_sheets:
        st.warning("Select at least one worksheet. The Summary sheet will be included automatically.")
        selected_sheets = ["Summary"] if "Summary" in sheet_names else [sheet_names[0]]
    column_choices: dict[str, list[str]] = {}
    with st.expander("Choose columns for each worksheet", expanded=False):
        for ws in wb.worksheets:
            if ws.title not in selected_sheets: continue
            col_opts=[]
            for c in range(1, ws.max_column + 1):
                label = ""
                for r in range(1, min(ws.max_row, 12) + 1):
                    v = ws.cell(r,c).value
                    if v is not None and str(v).strip():
                        label = str(v).replace("\n", " ").strip()[:55]; break
                letter = get_column_letter(c)
                col_opts.append((letter, f"{letter} — {label or 'Column '+letter}"))
            chosen_labels = st.multiselect(f"{ws.title} columns", [label for _,label in col_opts], default=[label for _,label in col_opts], key=f"{key_prefix}_cols_{ws.title}")
            selected_letters = [letter for letter,label in col_opts if label in chosen_labels]
            column_choices[ws.title] = selected_letters or ([col_opts[0][0]] if col_opts else [])
    row_mode = st.radio("Rows", ["All rows (recommended; preserves formula support rows)", "Populated rows only", "Choose rows individually"], horizontal=True, key=f"{key_prefix}_rows")
    row_choices: dict[str, list[int]] = {}
    if row_mode == "Choose rows individually":
        with st.expander("Choose rows to include", expanded=True):
            for ws in wb.worksheets:
                if ws.title not in selected_sheets: continue
                row_labels = []
                for r in range(1, ws.max_row + 1):
                    bits = [str(ws.cell(r,c).value).replace("\n", " ").strip()[:45] for c in range(1, min(ws.max_column, 4)+1) if ws.cell(r,c).value is not None]
                    row_labels.append(f"Row {r}: " + " | ".join(bits))
                selected_rows = st.multiselect(f"{ws.title} rows", row_labels, default=row_labels, key=f"{key_prefix}_rowpick_{ws.title}")
                row_choices[ws.title] = [i+1 for i,label in enumerate(row_labels) if label in selected_rows]
    wb.close()
    return _customise_costing_workbook(raw, selected_sheets, column_choices, row_mode, row_choices)


def show_quote_result(quote: dict[str, Any], pdf_bytes: bytes) -> None:
    st.success("Quotation generated and saved.")
    st.markdown("#### Option summary")
    options = quote.get("options") or [{"title": quote.get("product", "Quotation"), "product": quote.get("product", ""),
        "variant": quote.get("variant", ""), "pricing_mode": quote.get("pricing_mode", ""),
        "grand_total": quote.get("grand_total", 0), "lines": quote.get("lines", [])}]
    summary_df = pd.DataFrame([{
        "Option": opt.get("title", f"Option {idx}"), "Product": opt.get("product", ""),
        "Solution": opt.get("variant", ""), "Pricing method": opt.get("pricing_mode", ""),
        "Total VIP FJD": fmt_amount(opt.get("grand_total"), quote.get("amount_display", settings.get("amount_display", "whole"))), "Lines": len(opt.get("lines", [])),
    } for idx, opt in enumerate(options, 1)])
    st.dataframe(summary_df, hide_index=True, use_container_width=True)
    st.markdown("### Customer quotation PDF preview")
    _preview_pdf(pdf_bytes, 760)
    left, right = st.columns(2)
    base_name = f"Quotation_{_option_filename_text(quote.get('organisation', 'Customer'))}_{date.today().isoformat()}"
    with left:
        if st.download_button("Download customer quotation (PDF)", data=pdf_bytes, file_name=f"{base_name}.pdf", mime="application/pdf", type="primary", use_container_width=True):
            update_quote_status(str(quote.get("quote_no", "")), "Quote Sent")
    with right:
        with st.expander("Customise costing Excel download", expanded=False):
            costing_bytes = _costing_export_controls(quote, "new_quote_costing")
            st.download_button("Download selected costing workbook (Excel)", data=costing_bytes, file_name=f"Costing_{_option_filename_text(quote.get('organisation', 'Customer'))}_{date.today().isoformat()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with st.expander("Internal costing snapshot (not included in the customer PDF)"):
        st.json(quote.get("pricing_snapshot", {}))

# ── Navigation ──────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = BASE_DIR / "assets" / "datec_logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=170)
    st.markdown("## Costing & Quotes")
    st.caption("Internal software pricing workspace")
    page = st.radio("Workspace", ["Overview", "Create quotation", "Price book", "DatecPay tiers", "Quote history", "Terms & Conditions", "Settings"], label_visibility="collapsed")
    st.markdown("---")
    st.caption("Build 1.8  •  Quote lifecycle  •  PDF + Excel")


if page == "Overview":
    title_block("Costing & Quotations", "A central workspace for PayGlobal, DatecPay 2.0 and Epicor pricing.")
    quote_rows = list_quotes(500)
    now_month = datetime.now().strftime("%Y-%m")
    current_month_quotes = [q for q in quote_rows if str(q.get("created_at", "")).startswith(now_month)]
    all_total = sum(nfloat(q.get("grand_total")) for q in quote_rows)
    month_total = sum(nfloat(q.get("grand_total")) for q in current_month_quotes)
    won_count = sum(1 for q in quote_rows if str(q.get("status", "")) == "Quote Accepted")
    loss_count = sum(1 for q in quote_rows if str(q.get("status", "")) == "Quote Declined")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Saved quotations", f"{len(quote_rows):,}")
    m2.metric("This month", f"{len(current_month_quotes):,}")
    m3.metric("This month value", fmt_fjd(month_total))
    m4.metric("Saved quote value", fmt_fjd(all_total))
    m5.metric("Won", f"{won_count:,}")
    m6.metric("Lost", f"{loss_count:,}")
    pending_count = sum(1 for q in quote_rows if str(q.get("status", "Draft")) in {"Draft", "Sent"})
    st.caption(f"Workflow check: {pending_count} quotation(s) are still Draft or Sent and may need follow-up.")
    st.markdown("### Start a quotation")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**PayGlobal**")
        st.caption("Cloud, On-Premise and Hybrid; fixed-price or cost-based.")
    with c2:
        st.markdown("**DatecPay 2.0**")
        st.caption("Employee tiers, custom annual licence price, setup and optional assurance.")
    with c3:
        st.markdown("**Epicor**")
        st.caption("Kinetic upgrade preset, cost-plus pricing or a custom quote.")
    st.caption("To start, select **Create quotation** from the left-hand menu.")
    st.markdown("### Recent quotations")
    if quote_rows:
        show = pd.DataFrame(quote_rows[:10])
        show = show.drop(columns=[c for c in ["quote_no", "id", "created_at"] if c in show.columns])
        show = show.rename(columns={"quote_date": "Date", "customer": "Contact", "organisation": "Organisation", "product": "Product", "variant": "Variant", "pricing_mode": "Pricing method", "grand_total": "Option 1 / headline total (FJD)", "status": "Status"})
        st.dataframe(show, hide_index=True, use_container_width=True)
    else:
        st.info("No quotations have been saved yet. Open **Create quotation** to prepare your first quote.")
    st.markdown("### Pricing controls")
    st.write("Supplier buy prices can be updated centrally in the Price book. Any manual sell-price change made while preparing a quote is stored only in that quotation's snapshot; the master price book remains unchanged unless explicitly saved there.")

elif page == "Create quotation":
    title_block("Create quotation", "Build a professional customer quote with one or more pricing options.")
    st.info("**Source-price rule:** cost-based options start from the unit buy prices in your costing sheets. The app recalculates landed cost, margin, VAT and extended totals. It does not import old calculated totals from the workbooks or previous quotes. Fixed and custom VIP prices are maintained separately.")
    meta = quote_meta_inputs()
    st.markdown("---")

    if "quote_option_count" not in st.session_state:
        st.session_state["quote_option_count"] = 1
    option_count = int(st.session_state["quote_option_count"])
    add_col, remove_col, help_col = st.columns([1, 1, 3])
    if add_col.button("＋ Add option", type="secondary", use_container_width=True):
        st.session_state["quote_option_count"] = min(option_count + 1, 8)
        for key in ("generated_quote", "generated_quote_pdf", "generated_quote_xlsx"):
            st.session_state.pop(key, None)
        st.rerun()
    if remove_col.button("− Remove last", disabled=option_count <= 1, use_container_width=True):
        prefix = f"option_{option_count}_"
        for key in list(st.session_state.keys()):
            if str(key).startswith(prefix):
                del st.session_state[key]
        st.session_state["quote_option_count"] = max(1, option_count - 1)
        for key in ("generated_quote", "generated_quote_pdf", "generated_quote_xlsx"):
            st.session_state.pop(key, None)
        st.rerun()
    help_col.caption("Add up to 8 alternative options to the same quotation. Each option has its own product, pricing method, unit rates and assumptions.")

    option_contexts: list[dict[str, Any]] = []
    for option_idx in range(1, option_count + 1):
        prefix = f"option_{option_idx}_"
        with st.container(border=True):
            st.markdown(f"### Option {option_idx}")
            pcol, vcol, mcol = st.columns([1, 1, 1.4])
            product = pcol.selectbox("Product", PRODUCTS, key=f"{prefix}product")
            variants_by_product = {
                "PayGlobal": ["Cloud", "On-Premise", "Hybrid"],
                "DatecPay 2.0": ["Payroll"],
                "Epicor": ["Kinetic Upgrade"],
            }
            variants = variants_by_product[product]
            variant_key = f"{prefix}variant"
            if variant_key in st.session_state and st.session_state[variant_key] not in variants:
                del st.session_state[variant_key]
            variant = vcol.selectbox("Solution / variant", variants, key=variant_key)
            if product == "DatecPay 2.0":
                modes = ["Employee-tier pricing", "Custom annual licence VIP", "Fixed unit price book", "Custom line items"]
            else:
                modes = ["Cost-based (unit buy price)", "Fixed unit price book", "Custom line items"]
            mode_key = f"{prefix}mode"
            if mode_key in st.session_state and st.session_state[mode_key] not in modes:
                del st.session_state[mode_key]
            pricing_mode = mcol.selectbox("Pricing method", modes, key=mode_key)
            option_title = st.text_input("Option title (optional)", key=f"{prefix}title",
                                         placeholder=f"Option {option_idx} - e.g. PayGlobal Cloud / On-Premise")
            ctx: dict[str, Any] = {
                "option_idx": option_idx, "product": product, "variant": variant,
                "pricing_mode": pricing_mode, "option_title": option_title.strip(),
            }

            source_catalog = list_catalog(product=product, variant=variant)
            if product == "DatecPay 2.0" and pricing_mode in ("Employee-tier pricing", "Custom annual licence VIP"):
                st.markdown("#### DatecPay 2.0 unit pricing")
                tiers = list_tiers()
                a, b, c = st.columns(3)
                employees = int(a.number_input("Number of employees", min_value=1, max_value=100000, value=60, step=1, key=f"{prefix}employees"))
                setup_fee = b.number_input("Implementation & setup (FJD VIP)", min_value=0.0,
                                           value=float(settings.get("default_datecpay_setup", 1500.0)), step=100.0, key=f"{prefix}setup")
                vat_pct = c.number_input("VAT (%)", min_value=0.0, max_value=100.0,
                                         value=float(settings.get("default_vat_pct", 12.5)), step=0.5, key=f"{prefix}vat")
                assurance_enabled = st.checkbox("Include annual software assurance", value=True, key=f"{prefix}assurance_enabled")
                assurance_pct = st.number_input("Assurance (% of annual licence VIP)", min_value=0.0, max_value=300.0,
                                                value=float(settings.get("default_datecpay_assurance_pct", 30.0)), step=5.0,
                                                key=f"{prefix}assurance_pct", disabled=not assurance_enabled)
                custom_annual = None
                selected_tier = next((t for t in tiers if int(t["min_employees"]) <= employees <= int(t["max_employees"])), None)
                monthly_price = float(selected_tier["monthly_price"]) if selected_tier else 0.0
                if pricing_mode == "Custom annual licence VIP":
                    custom_annual = st.number_input("Custom annual licence unit price (VAT-inclusive VIP FJD)", min_value=0.0,
                                                    value=0.0, step=100.0, key=f"{prefix}custom_annual")
                    st.caption("Enter the approved annual unit selling price only. No previously calculated quote amount is pre-filled, and VAT is not added again to this custom VIP unit price.")
                else:
                    if selected_tier:
                        st.caption(f"Selected source tier: Tier {selected_tier['tier_no']} ({selected_tier['min_employees']}-{selected_tier['max_employees']} employees) at FJD {monthly_price:,.2f} per employee per month. The unit rate is multiplied by the employee count and 12 months, then VAT is applied.")
                    else:
                        st.warning("No DatecPay employee tier covers this employee count. Update DatecPay tiers before quoting.")
                ctx["datecpay_inputs"] = {
                    "employees": employees, "setup_fee": float(setup_fee), "vat_pct": float(vat_pct),
                    "assurance_enabled": bool(assurance_enabled), "assurance_pct": float(assurance_pct),
                    "custom_annual_license_vip": custom_annual, "monthly_price_per_employee": monthly_price,
                    "selected_tier": selected_tier, "tiers": tiers,
                }

            elif pricing_mode == "Cost-based (unit buy price)":
                with st.expander("Costing assumptions for this option", expanded=True):
                    assumptions = costing_assumption_inputs(product, variant, f"{prefix}assumptions_{product}_{variant}")
                    if product == "PayGlobal":
                        if variant == "On-Premise":
                            assurance_enabled = st.checkbox("Include annual Assurance", value=True, key=f"{prefix}assurance_enabled")
                            assurance_pct = st.number_input("Assurance Percentage (%)", min_value=0.0, max_value=300.0,
                                                            value=50.0, step=1.0, format="%.2f", key=f"{prefix}assurance_pct",
                                                            disabled=not assurance_enabled)
                            assurance_basis = st.selectbox("Assurance calculation basis", ["Annual licence VIP", "Annual landed cost"],
                                                           key=f"{prefix}assurance_basis", disabled=not assurance_enabled)
                        else:
                            assurance_enabled = False
                            assurance_pct = 0.0
                            assurance_basis = "Annual licence VIP"
                    else:
                        default_assurance_pct = float(settings.get("default_epicor_assurance_pct", 0.0))
                        assurance_enabled = st.checkbox("Calculate annual Assurance where applicable", value=(default_assurance_pct > 0), key=f"{prefix}assurance_enabled")
                        assurance_pct = st.number_input("Assurance Percentage (%)", min_value=0.0, max_value=300.0,
                                                        value=default_assurance_pct, step=1.0, format="%.2f", key=f"{prefix}assurance_pct",
                                                        disabled=not assurance_enabled)
                        assurance_basis = st.selectbox("Assurance calculation basis", ["Annual licence VIP", "Annual landed cost"],
                                                       key=f"{prefix}assurance_basis", disabled=not assurance_enabled)
                cost_catalog = [dict(r) for r in source_catalog if r.get("item_type") == "cost"]
                if product == "Epicor":
                    # Epicor buy costs are project-specific and change for every quote.
                    # Never preload a previously entered/saved cost into a new quote.
                    # Start the technical cost line at zero with quantity one so the
                    # preparer enters the current supplier cost for each proposal.
                    for row in cost_catalog:
                        if "technical unit cost" in str(row.get("item_name", "")).lower():
                            row["buy_price"] = 0.0
                            row["list_price"] = 0.0
                            row["default_qty"] = 1.0
                            row["notes"] = "Enter the current supplier unit cost for this Epicor quote; cost is intentionally not pre-filled."
                    st.info("**Epicor project costing:** supplier costs are variable, so the technical unit buy cost starts at **0.00 in the selected buy currency** and quantity **1** for every new quote. Enter the current supplier cost before generating the quotation.")
                cost_df = data_editor_cost(cost_catalog, assumptions["selected_currency"], f"{prefix}cost_editor")
                if product == "PayGlobal":
                    st.markdown("#### Costing table view")
                    cost_view = st.selectbox("Display", ["Quotation view (selected fields)", "Full costing table"],
                                             index=0, key=f"{prefix}cost_view")
                    preview_full, preview_summary = payglobal_costing_preview(cost_df, assumptions, variant)
                    preview = preview_summary if cost_view == "Quotation view (selected fields)" else preview_full
                    if not preview.empty:
                        if variant == "Hybrid":
                            st.caption("Combined view: local licence unit buys are in FJD and platform subscription buys are in NZD. Datec Total Buy remains in the source buy currency; Total L/C and VIP totals are in FJD. Use Full costing table to see each row's buy currency.")
                        elif variant == "Cloud":
                            st.caption("Datec Total Buy is in the selected source buy currency (normally NZD). Total L/C and VIP totals are in FJD.")
                        else:
                            st.caption("Datec Total Buy, Total L/C and VIP totals are in FJD for On-Premise.")
                        st.dataframe(preview, use_container_width=True, hide_index=True,
                                     column_config={c: st.column_config.NumberColumn(
                                         c, format=("%.4f" if c in ("Rate", "Margin factor", "VAT factor", "Duty factor")
                                                    else "%.0f" if c in ("Total VIP FJD (Monthly)", "Total VIP FJD (Yearly)")
                                                    else "%.2f"))
                                                    for c in preview.columns if c not in ("Description Of Items", "Vendor", "Buy Currency", "Billing")})
                    else:
                        st.info("Enter at least one positive quantity to display the costing calculations.")
                    ctx["cost_view"] = cost_view
                ctx["assumptions"] = assumptions
                ctx["assurance_enabled"] = bool(assurance_enabled)
                ctx["assurance_pct"] = float(assurance_pct)
                ctx["assurance_basis"] = assurance_basis
                ctx["cost_df"] = cost_df

            elif pricing_mode == "Fixed unit price book":
                st.markdown("#### Fixed unit-price lines")
                fixed_catalog = [r for r in source_catalog if r.get("item_type") == "fixed"]
                if product == "PayGlobal":
                    fixed_catalog = [r for r in fixed_catalog if "implementation / training services" not in str(r.get("item_name", "")).lower()]
                fixed_df = data_editor_fixed(fixed_catalog, f"{prefix}fixed_editor")
                ctx["fixed_df"] = fixed_df

            else:
                st.markdown("#### Custom unit-price lines")
                custom_df = data_editor_custom(f"{prefix}custom_editor")
                ctx["custom_df"] = custom_df

            if product == "PayGlobal":
                services_df, services_summary, services_total = payglobal_services_editor(
                    f"{prefix}services_editor", float(settings.get("default_service_rate", 165.0))
                )
                ctx["services_df"] = services_df
                ctx["services_summary"] = services_summary
                ctx["services_total"] = float(services_total)
            option_contexts.append(ctx)

    st.markdown("---")
    if st.button("Generate quotation and costing workbook", type="primary", use_container_width=True):
        if not meta["customer"] or not meta["organisation"]:
            st.error("Enter the TO contact person and organisation before generating the quotation.")
        elif not meta.get("prepared_phone") or not meta.get("prepared_email"):
            st.error("Enter the consultant phone and email in the quotation details, or set their defaults in Application settings, before generating the quotation.")
        else:
            built_options: list[dict[str, Any]] = []
            build_errors: list[str] = []
            for ctx in option_contexts:
                idx = ctx["option_idx"]
                product = ctx["product"]
                variant = ctx["variant"]
                pricing_mode = ctx["pricing_mode"]
                title = ctx["option_title"] or f"Option {idx} - {product} {variant}"
                lines: list[dict[str, Any]] = []
                costing_rows: list[dict[str, Any]] = []
                assurance_total = 0.0
                pricing_snapshot: dict[str, Any] = {}

                if product == "DatecPay 2.0" and pricing_mode in ("Employee-tier pricing", "Custom annual licence VIP"):
                    dp = ctx["datecpay_inputs"]
                    if pricing_mode == "Employee-tier pricing" and not dp.get("selected_tier"):
                        build_errors.append(f"Option {idx}: DatecPay does not have a tier for {dp['employees']} employees.")
                        continue
                    if pricing_mode == "Custom annual licence VIP" and nfloat(dp.get("custom_annual_license_vip")) <= 0:
                        build_errors.append(f"Option {idx}: enter a positive approved annual VIP unit price for the custom DatecPay licence.")
                        continue
                    calc = calculate_datecpay(dp["employees"], dp["tiers"], dp["vat_pct"], dp["setup_fee"],
                                              dp["assurance_enabled"], dp["assurance_pct"], dp["custom_annual_license_vip"])
                    lines = calc["lines"]
                    assurance_total = float(calc.get("assurance_total", 0))
                    if pricing_mode == "Employee-tier pricing":
                        costing_rows.append({
                            "description": f"DatecPay Payroll annual licence ({dp['employees']} employees; unit price per employee per month)",
                            "qty": dp["employees"], "buy_price": dp["monthly_price_per_employee"], "buy_currency": "FJD",
                            "discount_pct": 0, "freight_pct": 0, "quote_unit_override": None,
                            "billing_frequency": "Monthly", "assurance_eligible": True, "row_type": "datecpay_tier",
                            "price_basis": "Source unit price from DatecPay Pricing sheet; not a pre-calculated total",
                        })
                    else:
                        costing_rows.append({
                            "description": f"DatecPay Payroll annual licence ({dp['employees']} employees; custom VIP price)",
                            "qty": 1, "buy_price": 0, "buy_currency": "FJD", "discount_pct": 0, "freight_pct": 0,
                            "quote_unit_override": None, "billing_frequency": "Annual recurring", "assurance_eligible": True,
                            "row_type": "datecpay_custom_annual", "price_basis": "Manual custom annual licence VIP unit price",
                        })
                    if float(dp["setup_fee"]) > 0:
                        costing_rows.append({
                            "description": "Implementation and Setup", "qty": 1, "buy_price": 0, "buy_currency": "FJD",
                            "discount_pct": 0, "freight_pct": 0, "quote_unit_override": float(dp["setup_fee"]),
                            "billing_frequency": "One-time", "assurance_eligible": False, "row_type": "datecpay_setup",
                            "price_basis": "DatecPay setup unit VIP input from costing worksheet; totals are formula-linked",
                        })
                    pricing_snapshot = {
                        "datecpay": {
                            "employees": dp["employees"], "monthly_price_per_employee": dp["monthly_price_per_employee"],
                            "annual_license_vip": calc.get("annual_license_vip"), "setup_fee": dp["setup_fee"],
                            "vat_pct": dp["vat_pct"], "assurance_enabled": dp["assurance_enabled"],
                            "assurance_pct": dp["assurance_pct"], "custom_annual_license_vip": dp["custom_annual_license_vip"],
                            "tier": ({"tier_no": dp["selected_tier"]["tier_no"], "min_employees": dp["selected_tier"]["min_employees"],
                                      "max_employees": dp["selected_tier"]["max_employees"]} if dp.get("selected_tier") else None),
                        },
                        "vat_pct": dp["vat_pct"], "assurance_enabled": dp["assurance_enabled"], "assurance_pct": dp["assurance_pct"],
                        "note": "Tier pricing is calculated from the source monthly unit price per employee; setup and assurance are separate inputs.",
                    }
                    option_subtotal = float(calc["subtotal"])
                    option_total = float(calc["grand_total"])

                elif pricing_mode == "Cost-based (unit buy price)":
                    assumptions = ctx["assumptions"]
                    costing_rows = cost_editor_to_rows(ctx["cost_df"])
                    if product == "Epicor":
                        missing_epicor_costs = [
                            row.get("item_name", row.get("description", "Epicor line"))
                            for row in costing_rows
                            if nfloat(row.get("buy_price")) <= 0
                            and row.get("quote_unit_override") is None
                            and nfloat(row.get("fixed_unit_vip")) <= 0
                        ]
                        if missing_epicor_costs:
                            build_errors.append(
                                f"Option {idx}: enter the current supplier buy cost for each included Epicor line, "
                                "or uncheck any line that is not part of this quotation. Epicor costs start at zero by design."
                            )
                            continue
                    if product == "PayGlobal" and variant == "Hybrid":
                        licence_margin_pct = (nfloat(assumptions.get("margin_factor"), 1.0) - 1.0) * 100.0
                        platform_margin_pct = (nfloat(assumptions.get("platform_margin_factor"), 1.2) - 1.0) * 100.0
                        for cost_row in costing_rows:
                            cost_row["margin_pct_override"] = licence_margin_pct if str(cost_row.get("buy_currency", "FJD")).upper() == "FJD" else platform_margin_pct
                    if not costing_rows:
                        build_errors.append(f"Option {idx}: include at least one cost line with a description and positive quantity.")
                        continue
                    calc = calculate_cost_quote(
                        costing_rows,
                        selected_currency=assumptions["selected_currency"],
                        selected_fx_rate=assumptions["selected_fx_rate"],
                        all_fx_rates=assumptions["fx_rates"],
                        vat_pct=assumptions["vat_pct"], margin_pct=assumptions["margin_pct"], duty_pct=assumptions["duty_pct"],
                        assurance_enabled=ctx["assurance_enabled"], assurance_pct=ctx["assurance_pct"],
                        assurance_basis=ctx["assurance_basis"],
                        round_unit_vip_2dp=(product == "PayGlobal"),
                    )
                    lines = calc["lines"]
                    assurance_total = float(calc.get("assurance_total", 0))
                    option_subtotal = float(calc["subtotal"])
                    option_total = float(calc["grand_total"])
                    snapshot_assumptions = {
                        "vat_pct": assumptions["vat_pct"], "margin_pct": assumptions["margin_pct"], "duty_pct": assumptions["duty_pct"],
                        "vat_factor": assumptions.get("vat_factor"), "margin_factor": assumptions.get("margin_factor"),
                        "platform_margin_factor": assumptions.get("platform_margin_factor"), "duty_factor": assumptions.get("duty_factor"),
                        "fx_rates": assumptions["fx_rates"], "selected_currency": assumptions["selected_currency"],
                        "selected_fx_rate": assumptions["selected_fx_rate"], "assurance_enabled": ctx["assurance_enabled"],
                        "assurance_pct": ctx["assurance_pct"], "assurance_basis": ctx["assurance_basis"],
                    }
                    pricing_snapshot = {"assumptions": snapshot_assumptions,
                        "costing_rows": costing_rows,
                        "note": "Unit buy prices are the inputs. The app recalculates landed cost, margin, VAT and totals. Quote-only overrides never update the master price book."}

                elif pricing_mode == "Fixed unit price book":
                    lines = fixed_editor_to_lines(ctx["fixed_df"])
                    costing_rows = fixed_editor_to_costing_rows(ctx["fixed_df"], row_type="fixed")
                    if not lines:
                        build_errors.append(f"Option {idx}: enter at least one line with a description, positive quantity and unit VIP price.")
                        continue
                    option_subtotal = sum(nfloat(line.get("total_vip")) for line in lines)
                    option_total = option_subtotal
                    pricing_snapshot = {"costing_rows": costing_rows, "note": "Fixed master unit VIP prices or quote-only unit overrides are final VAT-inclusive prices; no additional VAT is applied."}

                elif pricing_mode == "Custom line items":
                    lines = custom_editor_to_lines(ctx["custom_df"])
                    costing_rows = custom_editor_to_costing_rows(ctx["custom_df"])
                    if not lines:
                        build_errors.append(f"Option {idx}: add at least one custom line with a description, positive quantity and unit VIP price.")
                        continue
                    option_subtotal = sum(nfloat(line.get("total_vip")) for line in lines)
                    option_total = option_subtotal
                    pricing_snapshot = {"costing_rows": costing_rows, "note": "Custom unit prices are final VAT-inclusive VIP values and are saved only to this quote."}

                else:
                    build_errors.append(f"Option {idx}: unsupported pricing method.")
                    continue

                services_snapshot = None
                if product == "PayGlobal":
                    service_df = ctx.get("services_df", pd.DataFrame())
                    service_lines: list[dict[str, Any]] = []
                    service_costing_rows: list[dict[str, Any]] = []
                    service_records: list[dict[str, Any]] = []
                    for _, sr in service_df.iterrows():
                        module = text_value(sr.get("Module"))
                        hours_by_activity = {act: max(0.0, nfloat(sr.get(f"{act} hrs"))) for act in PAYGLOBAL_SERVICE_ACTIVITIES}
                        hours_total = sum(hours_by_activity.values())
                        rate = max(0.0, nfloat(sr.get("Service rate FJD/hr"), 165.0))
                        gross = hours_total * rate
                        discount_amount = min(gross, max(0.0, nfloat(sr.get("Discount FJD"))))
                        service_total = max(0.0, gross - discount_amount)
                        service_records.append({"module": module, "hours_by_activity": hours_by_activity,
                                                "total_hours": hours_total, "service_rate": rate,
                                                "discount_fjd": discount_amount, "total_services_cost": service_total})
                        description = f"Implementation Services - {module}"
                        # Keep all module rows in the costing workbook so any later changes to service hours
                        # recalculate the option total. Only services with a positive total appear on the customer quote.
                        service_costing_rows.append({"description": description, "qty": 1, "buy_price": 0,
                            "list_price": 0, "buy_currency": "FJD", "discount_pct": 0, "freight_pct": 0,
                            "quote_unit_override": service_total, "billing_frequency": "One-time",
                            "assurance_eligible": False, "row_type": "service_module",
                            "price_basis": "Formula-linked to PayGlobal Services worksheet"})
                        if service_total > 0:
                            service_lines.append({"description": description, "qty": 1, "unit_vip": service_total,
                                                  "total_vip": service_total, "billing": "One-time",
                                                  "pricing_method": "PayGlobal services worksheet"})
                    services_total = sum(nfloat(x.get("total_services_cost")) for x in service_records)
                    costing_rows.extend(service_costing_rows)
                    if service_lines:
                        lines.extend(service_lines)
                        option_subtotal += services_total
                        option_total += services_total
                    services_snapshot = {"modules": service_records, "total_services_cost": services_total,
                                         "service_rate_default": float(settings.get("default_service_rate", 165.0)),
                                         "formula_note": "Total services per module = total hours × service rate − discount amount."}
                    pricing_snapshot["services"] = services_snapshot

                if not lines or option_total <= 0:
                    build_errors.append(f"Option {idx}: the option total must be greater than zero.")
                    continue
                built_options.append({
                    "option_number": idx, "title": title, "product": product, "variant": variant,
                    "pricing_mode": pricing_mode, "lines": lines, "subtotal": option_subtotal,
                    "grand_total": option_total, "assurance_total": assurance_total,
                    "costing_rows": costing_rows, "pricing_snapshot": pricing_snapshot,
                    "services_breakdown": (services_snapshot or {}).get("modules", []),
                    "services_total": float((services_snapshot or {}).get("total_services_cost", 0.0)),
                })

            if build_errors:
                for err in build_errors:
                    st.error(err)
            elif len(built_options) != len(option_contexts):
                st.error("Not all options could be priced. Review the option inputs and try again.")
            else:
                aggregate_snapshot = {"options": [
                    {"title": opt["title"], "product": opt["product"], "variant": opt["variant"],
                     "pricing_mode": opt["pricing_mode"], "pricing_snapshot": opt["pricing_snapshot"]}
                    for opt in built_options
                ], "note": "Alternative option totals are not added together. The headline database value uses Option 1 only."}
                draft_quote = {
                    "options": built_options,
                    "subtotal": float(built_options[0]["subtotal"]),
                    "grand_total": float(built_options[0]["grand_total"]),
                    "assurance_total": float(built_options[0]["assurance_total"]),
                    "pricing_snapshot": aggregate_snapshot,
                }
                saved_quote, pdf = add_metadata_and_save(draft_quote, meta)
                st.session_state["generated_quote"] = saved_quote
                st.session_state["generated_quote_pdf"] = pdf
                st.session_state["generated_quote_xlsx"] = build_costing_xlsx(saved_quote)
                st.session_state["_reset_quote_inputs"] = True
                st.rerun()

    if st.session_state.get("generated_quote"):
        show_quote_result(st.session_state["generated_quote"], st.session_state["generated_quote_pdf"])

elif page == "Price book":
    title_block("Product price book", "Maintain standard supplier buy prices and fixed sell prices separately from quote-only overrides.")
    st.info("Saving changes here updates the master price book. Every saved supplier buy-price / currency / discount / freight / fixed price change is recorded in Price history. Prices edited in Create quotation do not update this table.")
    all_items = list_catalog()
    if not all_items:
        st.warning("No price-book items are configured.")
    else:
        filter_cols = st.columns([1, 1, 1])
        prod_filter = filter_cols[0].selectbox("Filter by product", ["All"] + PRODUCTS, key="catalog_filter_product")
        var_options = sorted({r["variant"] for r in all_items if prod_filter == "All" or r["product"] == prod_filter})
        var_filter = filter_cols[1].selectbox("Filter by variant", ["All"] + var_options, key="catalog_filter_variant")
        type_filter = filter_cols[2].selectbox("Filter by pricing type", ["All", "fixed", "cost"], key="catalog_filter_type")
        filtered = [r for r in all_items if (prod_filter == "All" or r["product"] == prod_filter)
                    and (var_filter == "All" or r["variant"] == var_filter)
                    and (type_filter == "All" or r["item_type"] == type_filter)]
        editable = pd.DataFrame([{k: r.get(k) for k in ["id", "product", "variant", "item_name", "item_type", "default_qty", "list_price", "buy_price", "buy_currency", "discount_pct", "freight_pct", "fixed_unit_vip", "billing_frequency", "assurance_eligible", "notes"]} for r in filtered])
        catalog_edit = st.data_editor(editable, key="catalog_master_editor", use_container_width=True, hide_index=True, num_rows="dynamic",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True, format="%d", width="small"),
                "product": st.column_config.SelectboxColumn("Product", options=PRODUCTS, required=True),
                "variant": st.column_config.TextColumn("Variant", required=True),
                "item_name": st.column_config.TextColumn("Item name / description", required=True, width="large"),
                "item_type": st.column_config.SelectboxColumn("Price type", options=["fixed", "cost"], required=True),
                "default_qty": st.column_config.NumberColumn("Default qty", min_value=0.0, step=1.0, format="%.2f"),
                "list_price": st.column_config.NumberColumn("List Price", min_value=0.0, step=0.01, format="%.2f", help="Source workbook unit list price; edited in cents."),
                "buy_price": st.column_config.NumberColumn("Datec Unit Buy", min_value=0.0, step=0.01, format="%.2f", help="Decimal unit buy price; cent-level editing is enabled, including values such as 65.27."),
                "buy_currency": st.column_config.SelectboxColumn("Buy currency", options=CURRENCIES, required=True),
                "discount_pct": st.column_config.NumberColumn("Discount %", min_value=0.0, max_value=100.0, step=1.0),
                "freight_pct": st.column_config.NumberColumn("Freight %", min_value=0.0, max_value=300.0, step=1.0),
                "fixed_unit_vip": st.column_config.NumberColumn("Fixed VIP unit price", min_value=0.0, step=0.01, format="%.2f"),
                "billing_frequency": st.column_config.SelectboxColumn("Billing frequency", options=["One-time", "Monthly", "Annual recurring", "Quarterly recurring"], required=True),
                "assurance_eligible": st.column_config.CheckboxColumn("Assurance eligible", default=False),
                "notes": st.column_config.TextColumn("Notes", width="large"),
            })
        s1, s2 = st.columns([1, 2])
        if s1.button("Save price book changes", type="primary"):
            count = save_catalog(catalog_edit.to_dict(orient="records"))
            st.success(f"Price book saved. {count} field-level price changes logged to history.")
            st.rerun()
        with s2:
            export_bytes = catalog_edit.to_csv(index=False).encode("utf-8")
            st.download_button("Export filtered price book (CSV)", data=export_bytes, file_name="datec_price_book.csv", mime="text/csv")
        upload = st.file_uploader("Import price book CSV (include IDs to update existing items)", type=["csv"], key="catalog_csv_upload")
        if upload is not None:
            try:
                incoming = pd.read_csv(upload)
                st.caption(f"Loaded {len(incoming)} rows. Importing updates rows with existing IDs; blank IDs are treated as new items.")
                if st.button("Apply imported price book", key="apply_catalog_import"):
                    changed = save_catalog(incoming.to_dict(orient="records"))
                    st.success(f"Import complete. {changed} price fields changed.")
                    st.rerun()
            except Exception as exc:
                st.error(f"Could not read that CSV: {exc}")
    st.markdown("### Recent master price changes")
    history_rows = price_history(12)
    if history_rows:
        hist = pd.DataFrame(history_rows).rename(columns={"item_name": "Item", "field_name": "Field", "old_value": "Previous", "new_value": "New", "changed_at": "Changed at"})
        st.dataframe(hist[[c for c in ["Changed at", "Item", "Field", "Previous", "New"] if c in hist.columns]], hide_index=True, use_container_width=True)
    else:
        st.caption("No master price changes have been recorded yet.")

elif page == "DatecPay tiers":
    title_block("DatecPay 2.0 price tiers", "Edit the source employee-tier schedule used by the quote calculator.")
    st.info("The current tier rates were imported as shown in the supplied Excel workbook. This table uses a single tier price for all employees in the selected employee band.")
    st.warning("The supplied workbook has a few non-monotonic monthly rates in tiers 6–10. They are preserved exactly as provided; confirm the intended prices before changing or quoting those bands.")
    tiers = list_tiers()
    tier_df = pd.DataFrame([{k: r.get(k) for k in ["id", "tier_no", "min_employees", "max_employees", "monthly_price", "discount_pct"]} for r in tiers])
    edited_tiers = st.data_editor(tier_df, key="datecpay_tier_editor", use_container_width=True, hide_index=True, num_rows="dynamic",
        column_config={
            "id": st.column_config.NumberColumn("ID", disabled=True, format="%d", width="small"),
            "tier_no": st.column_config.NumberColumn("Tier", min_value=1, step=1),
            "min_employees": st.column_config.NumberColumn("Min employees", min_value=1, step=1),
            "max_employees": st.column_config.NumberColumn("Max employees", min_value=1, step=1),
            "monthly_price": st.column_config.NumberColumn("Monthly price / employee (FJD)", min_value=0.0, step=1.0, format="%.2f"),
            "discount_pct": st.column_config.NumberColumn("Discount %", min_value=0.0, max_value=100.0, step=5.0, format="%.2f"),
        })
    if st.button("Save DatecPay tiers", type="primary"):
        try:
            save_tiers(edited_tiers.to_dict(orient="records"))
            st.success("DatecPay tier schedule saved.")
            st.rerun()
        except Exception as exc:
            st.error(f"Unable to save tier schedule: {exc}")

elif page == "Quote history":
    title_block("Quote history", "Retrieve saved customer quotations and their internal costing worksheets.")
    quotes = list_quotes(5000)
    if not quotes:
        st.info("No quotations saved yet.")
    else:
        history_export = pd.DataFrame(quotes).copy()
        st.download_button(
            "Export quote history (CSV)", data=history_export.to_csv(index=False).encode("utf-8"),
            file_name=f"Datec_Quote_History_{date.today().isoformat()}.csv", mime="text/csv",
        )
        qdf = pd.DataFrame(quotes)
        qdf = qdf.drop(columns=[c for c in ["quote_no", "id"] if c in qdf.columns])
        qdf = qdf.rename(columns={"quote_date": "Date", "customer": "Contact", "organisation": "Organisation", "product": "Product", "variant": "Solution", "pricing_mode": "Pricing method", "grand_total": "Option 1 / headline total (FJD)", "status": "Status", "created_at": "Saved at"})
        preferred = ["Date", "Contact", "Organisation", "Product", "Solution", "Pricing method", "Option 1 / headline total (FJD)", "Status", "Saved at"]
        qdf = qdf[[col for col in preferred if col in qdf.columns]]
        history_status_filter = st.selectbox("Filter by status", ["All", "Draft", "Sent", "Accepted", "Rejected", "Expired", "Cancelled"], key="history_status_filter")
        if history_status_filter != "All" and "Status" in qdf.columns:
            qdf = qdf[qdf["Status"] == history_status_filter]
        st.dataframe(qdf, hide_index=True, use_container_width=True)

        # Use a descriptive label in the interface; the database reference remains internal only.
        option_labels: dict[str, str] = {}
        for row in quotes:
            label = f"{row.get('quote_date', '')} | {row.get('organisation', '')} | {row.get('product', '')} | {row.get('created_at', '')}"
            option_labels[label] = str(row.get("quote_no", ""))
        labels = list(option_labels.keys())
        chosen_label = st.selectbox("Select saved quotation", labels, key="history_selected_quote_label")
        chosen_quote_no = option_labels[chosen_label]
        selected_row = next((row for row in quotes if str(row.get("quote_no", "")) == chosen_quote_no), {})
        status_values = ["Draft", "Quote Sent", "Quote Accepted", "Quote Declined", "Expired", "Cancelled"]
        current_status = str(selected_row.get("status") or "Draft")
        status_index = status_values.index(current_status) if current_status in status_values else 0
        stat_col, reason_col, stat_button_col = st.columns([1, 2, 1])
        new_status = stat_col.selectbox("Quotation status", status_values, index=status_index, key=f"history_status_{chosen_quote_no}")
        loss_reason = reason_col.text_input("Reason for loss (required if declined)", value=str(selected_row.get("loss_reason", "") or ""), key=f"loss_reason_{chosen_quote_no}")
        if stat_button_col.button("Update status", key=f"update_status_{chosen_quote_no}"):
            try:
                if update_quote_status(chosen_quote_no, new_status, loss_reason):
                    st.success(f"Quotation status updated to {new_status}.")
                    st.rerun()
                else: st.error("The selected quotation could not be found.")
            except ValueError as exc: st.error(str(exc))
        action_a, action_b = st.columns(2)
        if action_a.button("Edit selected quotation", key=f"edit_quote_{chosen_quote_no}"):
            found_edit = get_quote_by_no(chosen_quote_no)
            if found_edit:
                st.session_state["editing_quote_no"] = chosen_quote_no
                st.session_state["editing_quote_data"] = found_edit["quote_data"]
        if action_b.button("Delete selected quotation", key=f"delete_quote_{chosen_quote_no}"):
            st.session_state["confirm_delete_quote_no"] = chosen_quote_no
        if st.session_state.get("confirm_delete_quote_no") == chosen_quote_no:
            st.warning("Deleting a quotation is permanent. Download its PDF/Excel first if you need an archive.")
            confirm_col, cancel_col = st.columns(2)
            if confirm_col.button("Confirm permanent deletion", type="primary", key=f"confirm_delete_{chosen_quote_no}"):
                if delete_quote_record(chosen_quote_no):
                    st.session_state.pop("confirm_delete_quote_no", None)
                    st.success("Quotation deleted."); st.rerun()
            if cancel_col.button("Cancel deletion", key=f"cancel_delete_{chosen_quote_no}"):
                st.session_state.pop("confirm_delete_quote_no", None); st.rerun()
        if st.session_state.get("editing_quote_no") == chosen_quote_no:
            st.markdown("#### Edit saved quotation")
            edit_data = st.session_state.get("editing_quote_data", {})
            st.caption("Update customer details, item descriptions, quantities and quote-only unit prices. Master price-book values are not changed.")
            ec1, ec2, ec3 = st.columns(3)
            edit_data["customer"] = ec1.text_input("TO / Customer contact", value=str(edit_data.get("customer", "")), key=f"edit_customer_{chosen_quote_no}")
            edit_data["organisation"] = ec2.text_input("Organisation", value=str(edit_data.get("organisation", "")), key=f"edit_org_{chosen_quote_no}")
            edit_data["subject"] = ec3.text_input("Subject", value=str(edit_data.get("subject", "")), key=f"edit_subject_{chosen_quote_no}")
            ec4, ec5, ec6 = st.columns(3)
            edit_data["email"] = ec4.text_input("Customer email", value=str(edit_data.get("email", "")), key=f"edit_email_{chosen_quote_no}")
            edit_data["phone"] = ec5.text_input("Customer phone", value=str(edit_data.get("phone", "")), key=f"edit_phone_{chosen_quote_no}")
            edit_data["terms"] = st.text_area("Terms & conditions for this quotation", value=str(edit_data.get("terms", "")), height=140, key=f"edit_terms_{chosen_quote_no}")
            edit_options = edit_data.get("options") or [{"title": edit_data.get("product", "Quotation"), "lines": edit_data.get("lines", []), "grand_total": edit_data.get("grand_total", 0), "subtotal": edit_data.get("subtotal", 0)}]
            for oi, opt in enumerate(edit_options, 1):
                st.markdown(f"**Option {oi}: {opt.get('title', opt.get('variant', 'Quotation'))}**")
                lines_now = opt.get("lines", [])
                if lines_now:
                    line_df = pd.DataFrame([{"Description": x.get("description", ""), "Qty": float(x.get("qty", 1) or 1), "Unit VIP FJD": float(x.get("unit_vip", 0) or 0)} for x in lines_now])
                    edited_lines = st.data_editor(line_df, key=f"edit_lines_{chosen_quote_no}_{oi}", use_container_width=True, hide_index=True, num_rows="dynamic", column_config={"Description": st.column_config.TextColumn(required=True), "Qty": st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f"), "Unit VIP FJD": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.2f")})
                    opt["lines"] = []
                    option_total = 0.0
                    for _, line_row in edited_lines.iterrows():
                        qty = float(line_row.get("Qty", 0) or 0); unit = float(line_row.get("Unit VIP FJD", 0) or 0); total = qty * unit
                        opt["lines"].append({"description": str(line_row.get("Description", "")), "qty": qty, "unit_vip": unit, "total_vip": total, "billing": "One-time"})
                        option_total += total
                    opt["subtotal"] = option_total; opt["grand_total"] = option_total
            save_edit_col, cancel_edit_col = st.columns(2)
            if save_edit_col.button("Save quotation changes", type="primary", key=f"save_edit_{chosen_quote_no}"):
                edit_data["options"] = edit_options
                if edit_options:
                    edit_data["grand_total"] = float(edit_options[0].get("grand_total", 0) or 0)
                    edit_data["subtotal"] = float(edit_options[0].get("subtotal", 0) or 0)
                    edit_data["lines"] = edit_options[0].get("lines", [])
                if update_saved_quote(chosen_quote_no, edit_data):
                    st.session_state.pop("editing_quote_no", None); st.session_state.pop("editing_quote_data", None)
                    st.session_state.pop("history_loaded_quote", None)
                    st.success("Quotation updated. Its PDF and Excel costing workbook will be regenerated when loaded."); st.rerun()
                else: st.error("Quotation not found.")
            if cancel_edit_col.button("Cancel editing", key=f"cancel_edit_{chosen_quote_no}"):
                st.session_state.pop("editing_quote_no", None); st.session_state.pop("editing_quote_data", None); st.rerun()
        if st.button("Load selected quotation", type="primary"):
            found = get_quote_by_no(option_labels[chosen_label])
            if found:
                st.session_state["history_loaded_quote"] = found["quote_data"]
                st.session_state["history_loaded_pdf"] = build_quote_pdf(found["quote_data"])
                st.session_state["history_loaded_xlsx"] = build_costing_xlsx(found["quote_data"])
                st.rerun()
        if st.session_state.get("history_loaded_quote"):
            loaded = st.session_state["history_loaded_quote"]
            st.markdown(f"### {loaded.get('organisation', '')} — {loaded.get('subject', loaded.get('product', 'Quotation'))}")
            options_loaded = loaded.get("options") or []
            if options_loaded:
                option_df = pd.DataFrame([{
                    "Option": item.get("title", f"Option {i}"), "Product": item.get("product", ""),
                    "Solution": item.get("variant", ""), "Pricing method": item.get("pricing_mode", ""),
                    "Total VIP FJD": fmt_amount(item.get("grand_total"), loaded.get("amount_display", settings.get("amount_display", "whole"))),
                } for i, item in enumerate(options_loaded, 1)])
                st.dataframe(option_df, hide_index=True, use_container_width=True)
            else:
                st.metric("Quotation total", fmt_fjd(loaded.get("grand_total")))
            st.markdown("### Customer quotation PDF preview")
            _preview_pdf(st.session_state["history_loaded_pdf"], 720)
            left, right = st.columns(2)
            base_name = f"Quotation_{_option_filename_text(loaded.get('organisation', 'Customer'))}_{date.today().isoformat()}"
            with left:
                if st.download_button("Download customer quotation (PDF)", data=st.session_state["history_loaded_pdf"], file_name=f"{base_name}.pdf", mime="application/pdf", type="primary", use_container_width=True):
                    update_quote_status(str(loaded.get("quote_no", chosen_quote_no)), "Quote Sent")
            with right:
                with st.expander("Customise costing Excel download", expanded=False):
                    workbook_bytes = _costing_export_controls(loaded, f"history_costing_{chosen_quote_no}")
                    st.download_button("Download selected costing workbook (Excel)", data=workbook_bytes, file_name=f"Costing_{_option_filename_text(loaded.get('organisation', 'Customer'))}_{date.today().isoformat()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with st.expander("Internal costing snapshot"):
                st.json(loaded.get("pricing_snapshot", {}))

elif page == "Terms & Conditions":
    title_block("Terms & Conditions", "Edit the wording that will appear on future customer quotations. Previously saved quotations keep their own terms snapshot.")
    st.info("Enter one term per line. Start a line with '- ' to format it as a bullet in the PDF. The app does not hard-code or silently add wording that is not in this editor.")
    current_terms = str(settings.get("default_terms", DEFAULT_TERMS))
    if "terms_editor" not in st.session_state:
        st.session_state["terms_editor"] = current_terms
    if st.button("Load Datec standard template", key="load_standard_terms"):
        st.session_state["terms_editor"] = DEFAULT_TERMS
        st.rerun()
    with st.form("terms_editor_form"):
        terms_editor = st.text_area("Quotation terms and conditions", key="terms_editor", height=360)
        save_terms = st.form_submit_button("Save terms & conditions", type="primary")
    if save_terms:
        save_settings({"default_terms": terms_editor.strip()})
        st.success("Terms & conditions saved. New quotations will use this wording.")
        st.rerun()
    st.markdown("### PDF preview")
    preview_lines = str(st.session_state.get("terms_editor", current_terms)).splitlines()
    for line in preview_lines:
        txt = line.strip()
        if not txt:
            st.write("")
        elif txt.startswith("- "):
            st.markdown(f"- {txt[2:]}")
        else:
            st.markdown(txt)

elif page == "Settings":
    title_block("Application settings", "Set the master defaults used to pre-fill new quotations. Quote-specific changes do not update these settings.")
    with st.form("settings_form"):
        st.markdown("### Company details")
        c1, c2 = st.columns(2)
        company_name = c1.text_input("Company name", value=str(settings.get("company_name", "")))
        company_email = c2.text_input("Company email", value=str(settings.get("company_email", "")))
        company_address = c1.text_area("Head office address", value=str(settings.get("company_address", "")), height=85)
        company_phone = c2.text_input("Head office phone", value=str(settings.get("company_phone", "")))
        branch_address = c1.text_area("Branch address", value=str(settings.get("branch_address", "")), height=85)
        branch_phone = c2.text_input("Branch phone", value=str(settings.get("branch_phone", "")))
        st.markdown("### Quotation preparer defaults")
        prep1, prep2, prep3 = st.columns(3)
        prepared_by_default = prep1.text_input("Default prepared by", value=str(settings.get("prepared_by_default", "")))
        prepared_by_phone_default = prep2.text_input("Default consultant phone", value=str(settings.get("prepared_by_phone_default", "")))
        prepared_by_email_default = prep3.text_input("Default consultant email", value=str(settings.get("prepared_by_email_default", "")))
        st.markdown("### Default costing assumptions")
        d1, d2, d3, d4 = st.columns(4)
        default_vat_pct = d1.number_input("Default VAT (%)", min_value=0.0, max_value=100.0, value=float(settings.get("default_vat_pct", 12.5)), step=0.5)
        default_margin_pct = d2.number_input("PayGlobal / general margin (%)", min_value=0.0, max_value=500.0, value=float(settings.get("default_margin_pct", 20.0)), step=1.0)
        default_epicor_margin_pct = d3.number_input("Epicor margin (%)", min_value=0.0, max_value=500.0, value=float(settings.get("default_epicor_margin_pct", 40.0)), step=1.0)
        default_duty_pct = d4.number_input("General duty (%)", min_value=0.0, max_value=300.0, value=float(settings.get("default_duty_pct", 18.0)), step=1.0)
        e1, e2, e3 = st.columns(3)
        fx_rate_USD = e1.number_input("USD FX (foreign units per FJD)", min_value=0.0001, max_value=1000.0, value=float(settings.get("fx_rate_USD", 0.435)), format="%.6f")
        fx_rate_AUD = e2.number_input("AUD FX (foreign units per FJD)", min_value=0.0001, max_value=1000.0, value=float(settings.get("fx_rate_AUD", 1.0)), format="%.6f")
        fx_rate_NZD = e3.number_input("NZD FX (foreign units per FJD)", min_value=0.0001, max_value=1000.0, value=float(settings.get("fx_rate_NZD", 0.745)), format="%.6f")
        f1, f2, f3 = st.columns(3)
        default_datecpay_assurance_pct = f1.number_input("DatecPay assurance (%)", min_value=0.0, max_value=300.0, value=float(settings.get("default_datecpay_assurance_pct", 30.0)), step=5.0)
        default_payglobal_assurance_pct = f2.number_input("PayGlobal assurance (%)", min_value=0.0, max_value=300.0, value=float(settings.get("default_payglobal_assurance_pct", 40.0)), step=5.0)
        default_epicor_assurance_pct = f3.number_input("Epicor assurance (%)", min_value=0.0, max_value=300.0, value=float(settings.get("default_epicor_assurance_pct", 0.0)), step=5.0)
        default_datecpay_setup = st.number_input("Default DatecPay setup unit price (FJD VIP)", min_value=0.0, value=float(settings.get("default_datecpay_setup", 1500.0)), step=100.0)
        st.markdown("### Quotation display")
        amount_display = st.selectbox(
            "Customer quotation amount format", ["whole", "decimal"],
            index=0 if str(settings.get("amount_display", "whole")) == "whole" else 1,
            format_func=lambda value: "Whole FJD (rounded to 0 decimals)" if value == "whole" else "Two decimal places (0.00)",
            help="Sets the default for new quotations. You can override it on each quote. Source unit buy/list prices remain editable to cents.",
        )
        save = st.form_submit_button("Save application settings", type="primary")
    if save:
        values = {
            "company_name": company_name, "company_address": company_address, "company_phone": company_phone,
            "company_email": company_email, "branch_address": branch_address, "branch_phone": branch_phone,
            "prepared_by_default": prepared_by_default,
            "prepared_by_phone_default": prepared_by_phone_default,
            "prepared_by_email_default": prepared_by_email_default,
            "default_vat_pct": default_vat_pct,
            "default_margin_pct": default_margin_pct, "default_epicor_margin_pct": default_epicor_margin_pct,
            "default_duty_pct": default_duty_pct,
            "fx_rate_USD": fx_rate_USD, "fx_rate_AUD": fx_rate_AUD, "fx_rate_NZD": fx_rate_NZD,
            "default_datecpay_assurance_pct": default_datecpay_assurance_pct,
            "default_payglobal_assurance_pct": default_payglobal_assurance_pct,
            "default_epicor_assurance_pct": default_epicor_assurance_pct,
            "default_datecpay_setup": default_datecpay_setup, "amount_display": amount_display,
        }
        save_settings(values)
        st.success("Settings saved. They will be used as defaults for new quotes.")
        st.rerun()

    st.markdown("### Data backup")
    st.caption("Download a consistent backup of the local database, including quotations, status history, price book, tier prices and application settings. The file contains customer/contact and commercial pricing data; store it securely.")
    try:
        backup_bytes = database_backup_bytes()
        st.download_button(
            "Download full application backup (.db)", data=backup_bytes,
            file_name=f"Datec_Costing_Quotation_Backup_{date.today().isoformat()}.db",
            mime="application/vnd.sqlite3",
        )
    except Exception as exc:
        st.error(f"Unable to create database backup: {exc}")
