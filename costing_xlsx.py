from __future__ import annotations

from io import BytesIO
from typing import Any

import xlsxwriter

NAVY = "#172F53"
BLUE = "#2D65B4"
PURPLE = "#37308D"
PALE = "#EEF3FA"
INPUT = "#FFF7D6"
ORANGE = "#F2B37A"
INK = "#202938"
GREY = "#657184"
WHITE = "#FFFFFF"
GREEN = "#E8F5EC"


def _num(value: Any, default: float = 0.0) -> float:
    try:
        out = float(value)
        return out if out == out and abs(out) != float("inf") else default
    except (TypeError, ValueError):
        return default


def _round_up_2dp(value: float) -> float:
    import math
    return math.ceil(max(0.0, float(value or 0.0)) * 100 - 1e-9) / 100.0


def _as_list(value: Any) -> list[dict[str, Any]]:
    return value if isinstance(value, list) else []


def _periods(billing: str) -> int:
    b = str(billing or "One-time").lower()
    if b.startswith("month"):
        return 12
    if b.startswith("quarter"):
        return 4
    return 1


def _normalise_option_rows(option: dict[str, Any]) -> list[dict[str, Any]]:
    rows = _as_list(option.get("costing_rows"))
    if rows:
        out = []
        for raw in rows:
            row = dict(raw)
            row.setdefault("description", row.get("item_name", "Custom item"))
            row.setdefault("qty", row.get("default_qty", 1))
            row.setdefault("buy_price", 0.0)
            row.setdefault("buy_currency", "FJD")
            row.setdefault("discount_pct", 0.0)
            row.setdefault("freight_pct", 0.0)
            row.setdefault("billing_frequency", "One-time")
            row.setdefault("assurance_eligible", False)
            row.setdefault("quote_unit_override", None)
            row.setdefault("row_type", "cost")
            out.append(row)
        return out

    # Backwards compatibility for quotes created with app v1.0.
    out = []
    for line in _as_list(option.get("lines")):
        out.append({
            "description": line.get("description", "Custom item"),
            "qty": line.get("qty", 1),
            "buy_price": line.get("buy_unit_price", 0),
            "buy_currency": line.get("buy_currency", "FJD"),
            "discount_pct": 0,
            "freight_pct": 0,
            "billing_frequency": line.get("billing", "One-time"),
            "assurance_eligible": line.get("assurance_eligible", False),
            "quote_unit_override": line.get("unit_vip", 0),
            "row_type": "custom",
        })
    return out


def build_costing_xlsx(quote: dict[str, Any]) -> bytes:
    """Build an editable, formula-driven internal costing workbook for a saved quotation."""
    buffer = BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    workbook.set_calc_mode("auto")
    money_fmt = workbook.add_format({"num_format": '"FJD "#,##0.00;[Red]("FJD "#,##0.00)', "border": 1, "border_color": "#D7DEE8"})
    whole_money_fmt = workbook.add_format({"num_format": '"FJD "#,##0;[Red]("FJD "#,##0)', "border": 1, "border_color": "#D7DEE8"})
    input_money_fmt = workbook.add_format({"num_format": '"FJD "#,##0.00;[Red]("FJD "#,##0.00)', "border": 1, "border_color": "#D7DEE8", "bg_color": INPUT})
    input_price_fmt = workbook.add_format({"num_format": "#,##0.00;[Red](#,##0.00)", "border": 1, "border_color": "#D7DEE8", "bg_color": INPUT})
    rate_fmt = workbook.add_format({"num_format": "0.000000", "border": 1, "border_color": "#D7DEE8", "bg_color": INPUT})
    pct_fmt = workbook.add_format({"num_format": "0.00", "border": 1, "border_color": "#D7DEE8", "bg_color": INPUT})
    qty_fmt = workbook.add_format({"num_format": "0.00", "border": 1, "border_color": "#D7DEE8", "bg_color": INPUT})
    text_input_fmt = workbook.add_format({"border": 1, "border_color": "#D7DEE8", "bg_color": INPUT, "text_wrap": True, "valign": "top"})
    text_fmt = workbook.add_format({"border": 1, "border_color": "#D7DEE8", "text_wrap": True, "valign": "top"})
    title_fmt = workbook.add_format({"bold": True, "font_size": 18, "font_color": WHITE, "bg_color": NAVY, "valign": "vcenter"})
    section_fmt = workbook.add_format({"bold": True, "font_size": 10, "font_color": WHITE, "bg_color": BLUE, "valign": "vcenter"})
    header_fmt = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": NAVY, "border": 1, "border_color": WHITE, "text_wrap": True, "valign": "vcenter", "align": "center"})
    label_fmt = workbook.add_format({"bold": True, "font_color": NAVY, "bg_color": PALE, "border": 1, "border_color": "#D7DEE8", "text_wrap": True, "valign": "top"})
    normal_fmt = workbook.add_format({"font_color": INK, "border": 1, "border_color": "#D7DEE8", "text_wrap": True, "valign": "top"})
    total_label_fmt = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": PURPLE, "border": 1, "border_color": PURPLE})
    total_money_fmt = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": PURPLE, "num_format": '"FJD "#,##0.00', "border": 1, "border_color": PURPLE})
    total_whole_money_fmt = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": PURPLE, "num_format": '"FJD "#,##0', "border": 1, "border_color": PURPLE})
    note_fmt = workbook.add_format({"font_color": GREY, "italic": True, "text_wrap": True, "valign": "top"})
    option_total_fmt = workbook.add_format({"bold": True, "font_color": NAVY, "bg_color": GREEN, "num_format": '"FJD "#,##0.00', "border": 1, "border_color": "#A9D1B3"})
    option_total_whole_fmt = workbook.add_format({"bold": True, "font_color": NAVY, "bg_color": GREEN, "num_format": '"FJD "#,##0', "border": 1, "border_color": "#A9D1B3"})
    summary_header_fmt = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": NAVY, "border": 1, "border_color": WHITE, "text_wrap": True, "valign": "vcenter"})

    options = _as_list(quote.get("options"))
    if not options:
        options = [{
            "title": f"Option 1 - {quote.get('product', '')} {quote.get('variant', '')}".strip(),
            "product": quote.get("product", ""), "variant": quote.get("variant", ""),
            "pricing_mode": quote.get("pricing_mode", ""), "lines": quote.get("lines", []),
            "subtotal": quote.get("subtotal", 0), "grand_total": quote.get("grand_total", 0),
            "assurance_total": quote.get("assurance_total", 0),
            "pricing_snapshot": quote.get("pricing_snapshot", {}),
        }]
    amount_display = quote.get("amount_display")
    if amount_display not in {"whole", "decimal"}:
        # Maintain v1.6 formatting for old quotes where no global display preference was stored.
        amount_display = "whole" if all(str(opt.get("product", "")) == "PayGlobal" for opt in options) else "decimal"
    whole_display = amount_display == "whole"
    display_money_fmt = whole_money_fmt if whole_display else money_fmt
    display_total_fmt = total_whole_money_fmt if whole_display else total_money_fmt
    display_option_total_fmt = option_total_whole_fmt if whole_display else option_total_fmt

    # Summary sheet
    summary = workbook.add_worksheet("Summary")
    summary.hide_gridlines(2)
    summary.set_column("A:A", 7)
    summary.set_column("B:B", 38)
    summary.set_column("C:C", 20)
    summary.set_column("D:D", 25)
    summary.set_column("E:F", 17)
    summary.set_column("G:G", 20)
    summary.merge_range("A1:G1", "DATEC | COSTING VERIFICATION WORKBOOK", title_fmt)
    summary.set_row(0, 32)
    summary.write("A3", "Customer contact", label_fmt)
    summary.write("B3", quote.get("customer", ""), normal_fmt)
    summary.write("D3", "Organisation", label_fmt)
    summary.merge_range("E3:G3", quote.get("organisation", ""), normal_fmt)
    summary.write("A4", "Quotation date", label_fmt)
    summary.write("B4", quote.get("quote_date", ""), normal_fmt)
    summary.write("D4", "Prepared by", label_fmt)
    summary.merge_range("E4:G4", quote.get("prepared_by", ""), normal_fmt)
    summary.merge_range("A6:G6", "OPTION SUMMARY - AMOUNTS ARE FORMULA-LINKED TO EACH OPTION SHEET", section_fmt)
    summary_headers = ["Option", "Option title", "Product", "Solution", "Pricing method", "VAT % / factor", "Option total (FJD)" ]
    for col, header in enumerate(summary_headers):
        summary.write(6, col, header, summary_header_fmt)
    option_sheet_names = []

    # Separate formula-driven Services worksheet, matching ServicesCost in the source workbook.
    service_formula_map: dict[int, dict[str, int]] = {}
    service_sheet = None
    if any(str(opt.get("product", "")) == "PayGlobal" and _as_list(opt.get("services_breakdown")) for opt in options):
        service_sheet = workbook.add_worksheet("Services")
        service_sheet.hide_gridlines(2)
        service_sheet.freeze_panes(4, 1)
        service_sheet.set_column("A:A", 15)
        service_sheet.set_column("B:G", 20)
        service_sheet.set_column("H:H", 14)
        service_sheet.set_column("I:L", 20)
        service_sheet.merge_range("A1:L1", "PAYGLOBAL | IMPLEMENTATION SERVICES COSTING", title_fmt)
        service_sheet.set_row(0, 30)
        service_sheet.merge_range("A2:L2", "Yellow cells are service-hour, rate and discount inputs. Module totals are formula-driven and flow through to each quotation option.", note_fmt)
        svc_row = 3
        service_headers = ["Module", "Setup / Configuration hrs", "Data Migration hrs", "Integration / Technical hrs", "Testing & UAT hrs", "Training hrs", "Go-Live / Implementation hrs", "Total Hours", "Service Rate (FJD/hr)", "Gross Services (FJD)", "Discount (FJD)", "Total Services Cost (FJD)"]
        for idx, opt in enumerate(options, start=1):
            modules = _as_list(opt.get("services_breakdown"))
            if str(opt.get("product", "")) != "PayGlobal" or not modules:
                continue
            service_sheet.merge_range(svc_row, 0, svc_row, 11, f"OPTION {idx} | {opt.get('title') or opt.get('variant', '')}", section_fmt)
            svc_row += 1
            for col, header in enumerate(service_headers):
                service_sheet.write(svc_row, col, header, header_fmt)
            service_sheet.set_row(svc_row, 36)
            svc_row += 1
            service_formula_map[idx] = {}
            for module in modules:
                module_name = str(module.get("module", ""))
                hours_by_activity = module.get("hours_by_activity") or {}
                row_num = svc_row + 1  # one-based Excel row number
                service_formula_map[idx][module_name] = row_num
                service_sheet.write(svc_row, 0, module_name, text_fmt)
                activity_keys = ["Setup / Configuration", "Data Migration", "Integration / Technical", "Testing & UAT", "Training", "Go-Live / Implementation"]
                for col, activity in enumerate(activity_keys, start=1):
                    service_sheet.write_number(svc_row, col, _num(hours_by_activity.get(activity, 0)), qty_fmt)
                total_hours = _num(module.get("total_hours", 0))
                rate = _num(module.get("service_rate", 165), 165)
                discount = _num(module.get("discount_fjd", 0))
                total_cost = _num(module.get("total_services_cost", 0))
                service_sheet.write_formula(svc_row, 7, f"=SUM(B{row_num}:G{row_num})", qty_fmt, total_hours)
                service_sheet.write_number(svc_row, 8, rate, input_money_fmt)
                service_sheet.write_formula(svc_row, 9, f"=H{row_num}*I{row_num}", display_money_fmt, total_hours * rate)
                service_sheet.write_number(svc_row, 10, discount, input_money_fmt)
                service_sheet.write_formula(svc_row, 11, f"=MAX(0,J{row_num}-K{row_num})", display_money_fmt, total_cost)
                svc_row += 1
            total_row_num = svc_row + 1
            service_sheet.merge_range(svc_row, 0, svc_row, 10, f"OPTION {idx} TOTAL SERVICES COST", total_label_fmt)
            first_module_row = total_row_num - len(modules)
            service_sheet.write_formula(svc_row, 11, f"=SUM(L{first_module_row}:L{total_row_num-1})", display_total_fmt,
                                        _num(opt.get("services_total", 0)))
            svc_row += 2
        service_sheet.set_landscape()
        service_sheet.fit_to_pages(1, 0)
        service_sheet.repeat_rows(3)

    for idx, option in enumerate(options, start=1):
        sheet_name = f"Option {idx}"
        option_sheet_names.append(sheet_name)
        ws = workbook.add_worksheet(sheet_name)
        ws.hide_gridlines(2)
        ws.freeze_panes(15, 2)
        widths = [7, 38, 11, 15, 13, 12, 12, 12, 12, 12, 10, 12, 14, 18, 18, 18, 18, 18, 18, 18, 18, 26, 18]
        for col, width in enumerate(widths):
            ws.set_column(col, col, width)
        ws.merge_range("A1:W1", f"COSTING WORKSHEET | {option.get('title') or sheet_name}", title_fmt)
        ws.set_row(0, 30)
        ws.write("A2", "Customer", label_fmt)
        ws.merge_range("B2:D2", quote.get("customer", ""), normal_fmt)
        ws.write("E2", "Organisation", label_fmt)
        ws.merge_range("F2:H2", quote.get("organisation", ""), normal_fmt)
        ws.write("I2", "Date", label_fmt)
        ws.merge_range("J2:L2", quote.get("quote_date", ""), normal_fmt)
        ws.write("M2", "Product / solution", label_fmt)
        ws.merge_range("N2:W2", f"{option.get('product', '')} - {option.get('variant', '')}", normal_fmt)

        snap = option.get("pricing_snapshot") or {}
        assumptions = snap.get("assumptions") or snap
        is_payglobal = str(option.get("product", "")) == "PayGlobal"
        fx = assumptions.get("fx_rates") or {}
        vat = _num(assumptions.get("vat_pct", snap.get("vat_pct", 12.5)), 12.5)
        margin = _num(assumptions.get("margin_pct", 20), 20)
        duty = _num(assumptions.get("duty_pct", 0), 0)
        vat_factor = _num(assumptions.get("vat_factor", 1 + vat / 100.0), 1 + vat / 100.0)
        margin_factor = _num(assumptions.get("margin_factor", 1 + margin / 100.0), 1 + margin / 100.0)
        platform_margin_factor = _num(assumptions.get("platform_margin_factor", margin_factor), margin_factor)
        duty_factor = _num(assumptions.get("duty_factor", 1 + duty / 100.0), 1 + duty / 100.0)
        rates = {
            "USD": _num(fx.get("USD", assumptions.get("fx_usd", 0.435)), 0.435),
            "AUD": _num(fx.get("AUD", assumptions.get("fx_aud", 1.0)), 1.0),
            "NZD": _num(fx.get("NZD", assumptions.get("fx_nzd", 0.745)), 0.745),
            "FJD": 1.0,
        }
        assure_pct = _num(assumptions.get("assurance_pct", snap.get("assurance_pct", 0)), 0)
        assure_enabled = bool(assumptions.get("assurance_enabled", snap.get("assurance_enabled", False)))
        assure_basis = str(assumptions.get("assurance_basis", snap.get("assurance_basis", "Annual licence VIP")))

        ws.merge_range("A4:L4", "Editable pricing assumptions (yellow cells are inputs)", section_fmt)
        if is_payglobal:
            margin_label = "Licence margin factor" if str(option.get("variant", "")) == "Hybrid" else "Margin factor"
            assumption_cells = [
                ("A5", "VAT factor", "B5", vat_factor, rate_fmt),
                ("C5", margin_label, "D5", margin_factor, rate_fmt),
                ("E5", "Duty factor", "F5", duty_factor, rate_fmt),
                ("G5", "USD rate", "H5", rates["USD"], rate_fmt),
                ("I5", "AUD rate", "J5", rates["AUD"], rate_fmt),
                ("K5", "NZD rate", "L5", rates["NZD"], rate_fmt),
            ]
            for label_cell, label, value_cell, value, fmt in assumption_cells:
                ws.write(label_cell, label, label_fmt)
                ws.write(value_cell, value, fmt)
            if str(option.get("variant", "")) == "Hybrid":
                ws.write("A6", "Platform margin factor", label_fmt)
                ws.write("B6", platform_margin_factor, rate_fmt)
            elif str(option.get("variant", "")) == "On-Premise":
                ws.write("A6", "Assurance %", label_fmt)
                ws.write("B6", assure_pct, pct_fmt)
                ws.write("C6", "Assurance enabled", label_fmt)
                ws.write("D6", "Yes" if assure_enabled else "No", text_input_fmt)
                ws.data_validation("D6", {"validate": "list", "source": ["Yes", "No"]})
                ws.write("E6", "Assurance basis", label_fmt)
                ws.write("F6", "Annual landed cost" if "landed" in assure_basis.lower() else "Annual licence VIP", text_input_fmt)
                ws.data_validation("F6", {"validate": "list", "source": ["Annual licence VIP", "Annual landed cost"]})
            ws.merge_range("G6:L6", "FX convention: foreign-currency units per FJD; FJD is always 1.000000.", note_fmt)
        else:
            assumption_cells = [
                ("A5", "VAT %", "B5", vat, pct_fmt),
                ("C5", "Margin %", "D5", margin, pct_fmt),
                ("E5", "Duty %", "F5", duty, pct_fmt),
                ("G5", "USD FX", "H5", rates["USD"], rate_fmt),
                ("I5", "AUD FX", "J5", rates["AUD"], rate_fmt),
                ("K5", "NZD FX", "L5", rates["NZD"], rate_fmt),
            ]
            for label_cell, label, value_cell, value, fmt in assumption_cells:
                ws.write(label_cell, label, label_fmt)
                ws.write(value_cell, value, fmt)
            ws.write("A6", "Assurance %", label_fmt)
            ws.write("B6", assure_pct, pct_fmt)
            ws.write("C6", "Assurance enabled", label_fmt)
            ws.write("D6", "Yes" if assure_enabled else "No", text_input_fmt)
            ws.data_validation("D6", {"validate": "list", "source": ["Yes", "No"]})
            ws.write("E6", "Assurance basis", label_fmt)
            ws.write("F6", "Annual landed cost" if "landed" in assure_basis.lower() else "Annual licence VIP", text_input_fmt)
            ws.data_validation("F6", {"validate": "list", "source": ["Annual licence VIP", "Annual landed cost"]})
            ws.merge_range("G6:L6", "FX convention: foreign currency units per FJD. FJD inputs always use 1.000000.", note_fmt)

        dp = snap.get("datecpay") or {}
        is_datecpay_calculator = str(option.get("pricing_mode", "")).lower().startswith(("employee-tier", "custom annual licence")) and bool(dp)
        if is_datecpay_calculator:
            ws.merge_range("A8:J8", "DatecPay unit-price calculator", section_fmt)
            ws.write("A9", "Employees", label_fmt)
            ws.write("B9", _num(dp.get("employees", dp.get("employee_count", 1)), 1), qty_fmt)
            ws.write("C9", "Monthly unit price / employee (FJD, ex VAT)", label_fmt)
            if dp.get("custom_annual_license_vip") is not None:
                dp_monthly = _num(dp.get("monthly_price_per_employee", 0))
                ws.write("D9", _num(dp.get("custom_annual_license_vip", 0)), input_money_fmt)
                ws.write("C9", "Custom annual licence VIP (FJD)", label_fmt)
            else:
                dp_monthly = _num(dp.get("monthly_price_per_employee", dp.get("monthly_price", 0)))
                ws.write("D9", dp_monthly, input_money_fmt)
            ws.write("E9", "VAT %", label_fmt)
            ws.write_formula("F9", "=$B$5", pct_fmt, vat)
            ws.write("G9", "Assurance %", label_fmt)
            ws.write_formula("H9", "=$B$6", pct_fmt, assure_pct)
            ws.write("I9", "Setup fee VIP (FJD)", label_fmt)
            ws.write("J9", _num(dp.get("setup_fee", dp.get("setup_fee_quote_only", 0))), input_money_fmt)
            if dp.get("custom_annual_license_vip") is not None:
                ws.write("A10", "Annual licence VIP", label_fmt)
                ws.write_formula("B10", "=D9", display_money_fmt, _num(dp.get("custom_annual_license_vip", 0)))
            else:
                ws.write("A10", "Annual VEP licence", label_fmt)
                ws.write_formula("B10", "=B9*D9*12", display_money_fmt, _num(dp.get("employees", 1)) * dp_monthly * 12)
                ws.write("C10", "VAT amount", label_fmt)
                ws.write_formula("D10", "=B10*F9/100", display_money_fmt, _num(dp.get("employees", 1)) * dp_monthly * 12 * vat / 100)
                ws.write("E10", "Annual licence VIP", label_fmt)
                ws.write_formula("F10", "=B10+D10", display_money_fmt, _num(dp.get("annual_license_vip", 0)))
            ws.write("A11", "Annual assurance", label_fmt)
            if dp.get("custom_annual_license_vip") is not None:
                license_cell = "B10"
            else:
                license_cell = "F10"
            cached_license = _num(dp.get("annual_license_vip", dp.get("custom_annual_license_vip", 0)))
            ws.write_formula("B11", f'=IF($D$6="Yes",{license_cell}*$B$6/100,0)', display_money_fmt,
                             cached_license * assure_pct / 100 if assure_enabled else 0)
            ws.write("C11", "Formula note", label_fmt)
            ws.merge_range("D11:J11", "Licence price is calculated from the monthly unit rate × employee count × 12, then VAT. Setup is a separate one-time VIP input.", note_fmt)
        else:
            ws.merge_range("A8:L8", "Only unit prices are entered as inputs. Landed cost, margin, VAT and extended totals below are formulas.", note_fmt)

        # The table uses raw unit prices only. It never imports pre-calculated totals from the source workbooks.
        headers = ["#", "Description", "Qty / units", "Buy / base unit price", "Currency", "Discount %", "Freight %", "Duty factor" if is_payglobal else "Duty %", "FX units / FJD", "Margin factor" if is_payglobal else "Margin %", "VAT factor" if is_payglobal else "VAT %", "Periods / year", "Assurance eligible", "Landed cost / unit FJD", "Landed cost total FJD", "Sell price ex VAT / unit", "Calculated unit VIP FJD", "Manual unit VIP override", "Final unit VIP FJD", "Total VIP FJD", "Billing frequency", "Price basis / source", "List Price (source reference)"]
        header_row = 14  # Excel row 15
        for col, header in enumerate(headers):
            ws.write(header_row, col, header, header_fmt)
        ws.set_row(header_row, 40)

        rows = _normalise_option_rows(option)
        # If relevant lines have annual assurance eligibility, include a verification row even when assurance was off in the quote.
        has_assurance_line = any("assurance" in str(r.get("description", "")).lower() and str(r.get("row_type", "")) == "assurance" for r in rows)
        eligible_present = any(bool(r.get("assurance_eligible", False)) for r in rows)
        if eligible_present and not has_assurance_line:
            rows.append({"description": "Assurance", "qty": 1, "buy_price": 0,
                         "buy_currency": "FJD", "discount_pct": 0, "freight_pct": 0,
                         "billing_frequency": "Annual recurring", "assurance_eligible": False,
                         "quote_unit_override": None, "row_type": "assurance", "price_basis": "Formula from eligible annual lines"})

        first_data_row = header_row + 1
        row_specs: list[dict[str, Any]] = []
        for excel_row, row in enumerate(rows, start=first_data_row + 1):
            # xlsxwriter APIs below use zero-based row indexes.
            r0 = excel_row - 1
            row_type = str(row.get("row_type", "cost")).lower()
            if "assurance" in str(row.get("description", "")).lower() and row_type != "custom":
                row_type = "assurance"
            description = str(row.get("description", row.get("item_name", "Custom item")))
            qty = _num(row.get("qty", row.get("default_qty", 1)), 1)
            buy_price = _num(row.get("buy_price", 0))
            currency = str(row.get("buy_currency", "FJD") or "FJD").upper()
            disc = _num(row.get("discount_pct", 0))
            freight = _num(row.get("freight_pct", 0))
            billing = str(row.get("billing_frequency", "One-time") or "One-time")
            periods = _periods(billing)
            assurance_eligible = bool(row.get("assurance_eligible", False))
            raw_override = row.get("quote_unit_override")
            override = None
            if raw_override is not None and str(raw_override).strip().lower() not in {"", "none", "nan"}:
                override = _num(raw_override, 0)
            source = str(row.get("price_basis", row.get("notes", "")) or "")
            if row_type == "datecpay_tier":
                source = source or "DatecPay monthly unit price per employee"
            elif row_type == "fixed":
                source = source or "Fixed master unit VIP price"
            elif row_type == "custom":
                source = source or "Quote-only custom VIP unit price"
            elif row_type == "datecpay_custom_annual":
                source = source or "DatecPay custom annual VIP unit price input"
            elif row_type == "datecpay_setup":
                source = source or "DatecPay setup VIP unit price input"
            elif row_type == "assurance":
                source = source or "Formula based on eligible annual lines"
            elif row_type == "service_module":
                source = "Formula linked to Services worksheet"
            else:
                source = source or "Supplier unit buy price; calculated in this workbook"

            ws.write_number(r0, 0, excel_row - first_data_row, normal_fmt)
            ws.write(r0, 1, description, text_input_fmt)
            ws.write_number(r0, 2, qty, qty_fmt)
            if row_type == "datecpay_tier":
                ws.write_formula(r0, 3, "=$D$9", input_price_fmt, buy_price)
            else:
                ws.write_number(r0, 3, buy_price, input_price_fmt)
            ws.write(r0, 4, currency, text_input_fmt)
            ws.write_number(r0, 5, disc, pct_fmt)
            ws.write_number(r0, 6, freight, pct_fmt)
            ws.write_formula(r0, 7, "=$F$5", rate_fmt if is_payglobal else pct_fmt, duty_factor if is_payglobal else duty)
            fx_formula = f'=IF(E{excel_row}="USD",$H$5,IF(E{excel_row}="AUD",$J$5,IF(E{excel_row}="NZD",$L$5,1)))'
            fx_cached = rates.get(currency, 1.0)
            ws.write_formula(r0, 8, fx_formula, rate_fmt, fx_cached)
            row_margin_pct = _num(row.get("margin_pct_override", margin), margin)
            row_margin_factor = 1.0 + row_margin_pct / 100.0
            if is_payglobal and "margin_pct_override" in row and str(option.get("variant", "")) == "Hybrid":
                # Keep Hybrid's local-licence and platform margin controls live in the exported workbook.
                margin_formula = f'=IF(E{excel_row}="FJD",$D$5,$B$6)'
            elif is_payglobal and "margin_pct_override" in row:
                margin_formula = "=$D$5"
            else:
                margin_formula = "=$D$5"
            ws.write_formula(r0, 9, margin_formula, rate_fmt if is_payglobal else pct_fmt, row_margin_factor if is_payglobal else row_margin_pct)
            ws.write_formula(r0, 10, "=$B$5", rate_fmt if is_payglobal else pct_fmt, vat_factor if is_payglobal else vat)
            ws.write_number(r0, 11, periods, qty_fmt)
            ws.write(r0, 12, "Yes" if assurance_eligible else "No", text_input_fmt)
            ws.data_validation(r0, 12, r0, 12, {"validate": "list", "source": ["Yes", "No"]})

            if row_type == "assurance":
                ws.write_formula(r0, 13, "=0", money_fmt, 0)
                ws.write_formula(r0, 14, "=0", money_fmt, 0)
                ws.write_formula(r0, 15, "=0", money_fmt, 0)
                ws.write_formula(r0, 16, "=0", money_fmt, 0)
                ws.write_blank(r0, 17, None, input_money_fmt)
                # Formula references only rows above this one to avoid circular references.
                last_prev = excel_row - 1
                first_data_excel = first_data_row + 1
                if last_prev >= first_data_excel:
                    landed_range = f"$O${first_data_excel}:$O${last_prev}"
                    total_range = f"$T${first_data_excel}:$T${last_prev}"
                    eligible_range = f"$M${first_data_excel}:$M${last_prev}"
                    annual_landed_formula = f'SUMPRODUCT(({eligible_range}="Yes")*({landed_range})*($L${first_data_excel}:$L${last_prev}))'
                    base_formula = f'IF($F$6="Annual landed cost",{annual_landed_formula},SUMIFS({total_range},{eligible_range},"Yes"))'
                else:
                    base_formula = "0"
                ass_formula = f'=IF($D$6="Yes",({base_formula})*$B$6/100,0)'
                assurance_cached = _num(option.get("assurance_total", 0)) if assure_enabled else 0
                ws.write_formula(r0, 18, ass_formula, money_fmt, assurance_cached)
            else:
                if is_payglobal:
                    landed_formula = f'=IFERROR((D{excel_row}*(1-F{excel_row}/100)*(1+G{excel_row}/100)*H{excel_row})/I{excel_row},0)'
                    unit_landed_cached = (buy_price * (1-disc/100) * (1+freight/100) * duty_factor / (fx_cached or 1)) if currency != "FJD" else buy_price * (1-disc/100) * (1+freight/100) * duty_factor
                    effective_margin_factor = row_margin_factor
                    unit_vep_cached = unit_landed_cached * effective_margin_factor
                    unit_vip_cached = _round_up_2dp(unit_vep_cached * vat_factor)
                else:
                    landed_formula = f'=IFERROR((D{excel_row}*(1-F{excel_row}/100)*(1+G{excel_row}/100)*(1+H{excel_row}/100))/I{excel_row},0)'
                    unit_landed_cached = (buy_price * (1-disc/100) * (1+freight/100) * (1+duty/100) / (fx_cached or 1)) if currency != "FJD" else buy_price * (1-disc/100) * (1+freight/100) * (1+duty/100)
                    effective_margin_factor = 1 + row_margin_pct / 100.0
                    unit_vep_cached = unit_landed_cached * effective_margin_factor
                    unit_vip_cached = unit_vep_cached * (1 + vat/100)
                ws.write_formula(r0, 13, landed_formula, money_fmt, unit_landed_cached)
                # Total landed cost is the current item quantity, not annualised subscription periods.
                ws.write_formula(r0, 14, f'=C{excel_row}*N{excel_row}', money_fmt, qty * unit_landed_cached)
                vep_formula = f'=N{excel_row}*J{excel_row}' if is_payglobal else f'=N{excel_row}*(1+J{excel_row}/100)'
                ws.write_formula(r0, 15, vep_formula, money_fmt, unit_vep_cached)
                # DatecPay's tier price is a unit sell rate before VAT; no supplier duty/margin is applied to this unit rate.
                if row_type == "datecpay_tier":
                    ws.write_formula(r0, 13, f'=D{excel_row}', money_fmt, buy_price)
                    ws.write_formula(r0, 14, f'=C{excel_row}*N{excel_row}*L{excel_row}', money_fmt, qty * buy_price * periods)
                    ws.write_formula(r0, 15, f'=N{excel_row}', money_fmt, buy_price)
                    ws.write_formula(r0, 16, f'=P{excel_row}*(1+K{excel_row}/100)', money_fmt, buy_price * (1+vat/100))
                elif row_type == "datecpay_custom_annual":
                    custom_vip = _num(dp.get("custom_annual_license_vip", 0))
                    ws.write_formula(r0, 13, "=0", money_fmt, 0)
                    ws.write_formula(r0, 14, "=0", money_fmt, 0)
                    ws.write_formula(r0, 15, "=IFERROR($D$9/(1+$B$5/100),0)", money_fmt, custom_vip / (1 + vat/100) if (1 + vat/100) else custom_vip)
                    ws.write_formula(r0, 16, "=$D$9", money_fmt, custom_vip)
                elif row_type == "datecpay_setup":
                    setup_vip = _num(dp.get("setup_fee", 0))
                    ws.write_formula(r0, 13, "=0", money_fmt, 0)
                    ws.write_formula(r0, 14, "=0", money_fmt, 0)
                    ws.write_formula(r0, 15, "=$J$9", money_fmt, setup_vip)
                    ws.write_formula(r0, 16, "=$J$9", money_fmt, setup_vip)
                else:
                    vip_formula = f'=ROUNDUP(P{excel_row}*K{excel_row},2)' if is_payglobal else f'=P{excel_row}*(1+K{excel_row}/100)'
                    ws.write_formula(r0, 16, vip_formula, money_fmt, unit_vip_cached)
                if row_type == "service_module" and service_formula_map.get(idx):
                    module_name = description.replace("Implementation Services - ", "", 1)
                    service_row_num = service_formula_map[idx].get(module_name)
                    if service_row_num:
                        ws.write_formula(r0, 17, f"='Services'!$L${service_row_num}", input_money_fmt, override or 0)
                    elif override is not None:
                        ws.write_number(r0, 17, override, input_money_fmt)
                    else:
                        ws.write_blank(r0, 17, None, input_money_fmt)
                elif override is None or row_type in {"datecpay_custom_annual", "datecpay_setup"}:
                    ws.write_blank(r0, 17, None, input_money_fmt)
                else:
                    ws.write_number(r0, 17, override, input_money_fmt)
                if row_type == "datecpay_custom_annual":
                    ws.write_formula(r0, 18, "=$B$10", money_fmt, _num(dp.get("custom_annual_license_vip", 0)))
                elif row_type == "datecpay_setup":
                    ws.write_formula(r0, 18, "=$J$9", money_fmt, _num(dp.get("setup_fee", 0)))
                else:
                    if override is not None and override > 0:
                        final_unit_cached = override
                    elif row_type == "datecpay_tier":
                        final_unit_cached = buy_price * (1 + vat / 100.0)
                    else:
                        # Cache must agree with the same unit margin factor as the live formula,
                        # especially for Hybrid local licence rows using a 1.000 margin factor.
                        final_unit_cached = unit_vip_cached
                    ws.write_formula(r0, 18, f'=IF(AND(ISNUMBER(R{excel_row}),R{excel_row}>0),R{excel_row},Q{excel_row})', money_fmt, final_unit_cached)

            total_cached = _num(row.get("calculated_total", 0))
            if row_type == "assurance":
                total_cached = _num(option.get("assurance_total", 0)) if assure_enabled else 0
            elif override is not None and override > 0:
                total_cached = qty * override * periods
            elif row_type == "datecpay_tier":
                total_cached = qty * buy_price * (1+vat/100) * periods
            elif row_type == "datecpay_custom_annual":
                total_cached = _num(dp.get("custom_annual_license_vip", 0)) * qty * periods
            elif row_type == "datecpay_setup":
                total_cached = _num(dp.get("setup_fee", 0)) * qty * periods
            else:
                if is_payglobal:
                    landed_cached = buy_price * (1-disc/100) * (1+freight/100) * duty_factor / (fx_cached or 1) if currency != "FJD" else buy_price * (1-disc/100) * (1+freight/100) * duty_factor
                    # Excel source rounds calculated unit VIP upward to cents before extension.
                    rounded_unit_vip = _round_up_2dp(landed_cached * row_margin_factor * vat_factor)
                    total_cached = qty * rounded_unit_vip * periods
                else:
                    landed_cached = buy_price * (1-disc/100) * (1+freight/100) * (1+duty/100) / (fx_cached or 1) if currency != "FJD" else buy_price * (1-disc/100) * (1+freight/100) * (1+duty/100)
                    total_cached = qty * landed_cached * (1+row_margin_pct/100) * (1+vat/100) * periods
            total_formula = f'=S{excel_row}*C{excel_row}*L{excel_row}'
            total_fmt_for_row = display_money_fmt
            ws.write_formula(r0, 19, total_formula, total_fmt_for_row, total_cached)
            ws.write(r0, 20, billing, text_fmt)
            ws.write(r0, 21, source, text_fmt)
            list_price = _num(row.get("list_price", 0))
            if list_price <= 0 and buy_price > 0:
                list_price = buy_price
            ws.write_number(r0, 22, list_price, input_price_fmt)
            row_specs.append({"row": excel_row, "type": row_type, "eligible": assurance_eligible})

        last_data_row0 = first_data_row + len(rows)
        last_excel_row = first_data_row + len(rows)
        # Add one summary line. Totals recalculate when any yellow input changes.
        total_row0 = last_data_row0
        total_excel = total_row0 + 1
        ws.merge_range(total_row0, 0, total_row0, 18, "OPTION TOTAL (FJD VIP - YEAR 1 / ONE-TIME + ANNUAL RECURRING)", total_label_fmt)
        sum_formula = f'=SUM(T{first_data_row+1}:T{last_excel_row})' if rows else "=0"
        summary_cached = _num(option.get("grand_total", 0))
        total_line_fmt = display_total_fmt
        ws.write_formula(total_row0, 19, sum_formula, total_line_fmt, summary_cached)
        ws.merge_range(total_row0, 20, total_row0, 22, "Internal worksheet - not for customer issue", total_label_fmt)
        ws.freeze_panes(header_row + 1, 2)
        ws.autofilter(header_row, 0, max(header_row, last_data_row0 - 1), len(headers) - 1)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(header_row)

        # Summary row. Formula points directly to the option worksheet's total cell.
        sum_row = 7 + idx - 1
        summary.write_number(sum_row, 0, idx, normal_fmt)
        summary.write(sum_row, 1, option.get("title") or f"Option {idx}", text_fmt)
        summary.write(sum_row, 2, option.get("product", ""), text_fmt)
        summary.write(sum_row, 3, option.get("variant", ""), text_fmt)
        summary.write(sum_row, 4, option.get("pricing_mode", ""), text_fmt)
        summary.write_number(sum_row, 5, vat_factor if is_payglobal else vat, rate_fmt if is_payglobal else pct_fmt)
        summary.write_formula(sum_row, 6, f"='{sheet_name}'!T{total_excel}",
                              display_option_total_fmt, summary_cached)

    summary.merge_range(7 + len(options) + 1, 0, 7 + len(options) + 1, 6,
                        "Source policy: only unit buy/base prices are inputs. Extended totals from the original spreadsheets or previous quotations are not imported. Formula outputs recalculate when inputs change.", note_fmt)
    summary.freeze_panes(7, 0)
    summary.autofilter(6, 0, 6 + len(options), 6)
    summary.set_landscape()
    summary.fit_to_pages(1, 1)
    workbook.close()
    return buffer.getvalue()
