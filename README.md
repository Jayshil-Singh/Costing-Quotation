# Datec Costing & Quotation App — v1.9

A local Python web application for preparing professional Datec customer quotations and editable costing-verification workbooks for PayGlobal, DatecPay 2.0 and Epicor.

## What it does

- Creates one customer quotation with up to **8 alternative options**. Each option can use a different product, deployment, pricing method, assumptions and unit rates.
- Supports cost-based pricing from **unit buy prices only**. VAT, margin, duty, currency conversion, assurance and extended totals are recalculated by the app.
- Keeps source prices, saved master prices and quote-only price overrides separate. A quote override does not update the master price book.
- Supports fixed unit VIP prices, custom unit VIP prices and the DatecPay employee-tier calculator.
- Offers currency selection for USD, AUD, NZD and FJD. The exchange-rate convention follows the costing workbooks: foreign currency units per FJD (for example, 0.745 means FJD 1 = NZD 0.745).
- Uses a tactile skeuomorphic application interface with raised cards, inset input fields, subtle gradients and clear visual sections.
- Generates a Datec-style A4 customer PDF using the attached quotation layout. The in-app preview renders each PDF page as an image (rather than an embedded data-URL PDF iframe) to avoid Chrome’s “This page has been blocked” preview error. The customer PDF intentionally omits the quotation number and a separate “Valid Until” header field. Validity wording can remain in Terms & Conditions as shown in the source quotation example.
- Uses the PayGlobal customer-quote layout from the latest reference: dark navy section bars, S/No. / Description / Qty / Unit VIP FJD / Total VIP FJD columns, a separate implementation-services block and subtotal, a highlighted recurring-fees note for Cloud, and a prominent grand total. Cloud subscription components are grouped into one annual package line with included components listed underneath. The customer-facing quote supports whole-FJD or two-decimal amount display, selectable per quotation.
- Offers two PayGlobal costing views: the full costing table and a quotation view showing only the requested fields for Cloud, On-Premise or Hybrid. Source unit prices are editable to cents (for example, NZD 65.27). PayGlobal calculated unit VIP prices use the source workbook formula ROUNDUP(unit VIP, 2) before extension by quantity and periods. Extended totals retain full precision internally and use the amount-display preference selected for that quotation. Buy, list, landed-cost, FX and other calculation inputs retain decimal precision.
- Calculates PayGlobal implementation services separately by Payroll, ESS, HR, T&A and Rostering. Service hours are entered by activity and priced at an editable FJD/hour rate, with discounts and totals linked into the quotation.
- Scope-specific PayGlobal service bullets respect the selected product and service scope. “Build ESS” appears only when an ESS product/subscription is included and ESS implementation work has been costed; it is omitted when ESS is not included.
- Quotation greeting uses the contact name entered in the `TO` field. The quotation metadata shows the customer email/phone and the phone/email of the consultant preparing the quote; the Datec head-office details remain in the page header/footer. Consultant contact defaults can be saved in Application Settings. The consultant phone and email must be provided before the PDF can be generated.
- Epicor buy costs are quote-specific and start at zero for each new cost-based quote, with quantity 1 on the main technical cost line. The preparer must enter the current supplier cost; the app does not prefill an older project cost and blocks generation if an included Epicor cost line has no price.
- Generates an Excel costing verification workbook with a Summary sheet and a separate sheet for every option, plus a formula-driven Services sheet where PayGlobal services are included. Yellow cells are editable unit inputs and assumptions; calculation columns use formulas for landed cost, sell price, VAT, assurance and totals.
- Saves quotes, price-history entries, price book data and settings in SQLite. The quote's costing worksheet is downloadable again from Quote history.
- Adds a dedicated **Terms & Conditions** navigation page with an editable standard terms template, preview, and restore-default action. New quotes use the current saved terms; old quote snapshots retain their original wording.
- Lets users choose the amount display per quotation: **whole FJD** (rounded to zero decimal places) or **two decimal places**. The preference also follows through to calculated total formats in the Excel verification workbook; source buy/list unit prices still retain cents.
- Adds quote status tracking (Draft, Sent, Accepted, Rejected, Expired, Cancelled), quote-history CSV export, and downloadable SQLite database backup.

## Source price policy

The default price book imports raw unit buy/base prices from the supplied costing workbooks. It does **not** seed previously calculated selling amounts, extended totals, or figures adjusted in older customer quotes as master unit prices. The workbook's monthly DatecPay tier value is treated as a unit rate per employee per month; the app then multiplies it by employee count and 12 months and calculates VAT. The DatecPay setup input defaults to FJD 1,500, from the costing calculator, not the FJD 2,500 adjusted in the sample customer quotation. The custom annual VIP mode starts at zero and requires a user-entered approved unit price.

PayGlobal costing uses workbook factors: VAT factor 1.125 and duty factor 1.000 by default; Cloud and On-Premise margin factor 1.20; Hybrid uses a local licence margin factor of 1.00 and a platform/subscription margin factor of 1.20. On-Premise Assurance defaults to 50% from the supplied workbook and is shown to customers only as “Assurance”, without a percentage in the quote line. Epicor uses the HFC costing workbook margin of 40% and duty of 18%; these settings are editable per option. Exchange rates are editable and should be confirmed before use.

The supplied DatecPay employee tier prices for tiers 6–10 are preserved as uploaded. They are editable from the DatecPay tiers page; confirm them with the owner of the source price list before use.

## PayGlobal costing details

The quotation view contains the requested fields for each variant. Cloud shows Monthly and Yearly VIP totals; On-Premise and Hybrid show Yearly VIP totals. Hybrid combines local licence and PaaS/platform lines in one costing table; the full view shows each line's buy currency. “Datec Total Buy” stays in the source buy currency, while Landed Cost and VIP values are in FJD.

## Run on Windows

1. Install Python 3.10 or newer and ensure the Python launcher (`py`) is available.
2. Extract the ZIP file to a stable working directory.
3. Double-click `start_app.bat`.
4. On first run, the launcher creates `.venv`, installs the packages in `requirements.txt` (an internet connection is required) and opens Streamlit in the browser.
5. Your SQLite data is stored in `data/costing_quotes.db`. Back up the `data` folder to preserve settings, price history and saved quotes.

## Run manually

```powershell
py -3 -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

This starter is intended for local/internal use. For multiple staff on a network, deploy to a secured internal server and add user authentication and role-based permissions before production use.


## Version 1.8 changes

- Edit saved quotations using customer fields, item descriptions, quantities and quote-only unit VIP prices. Master price-book prices remain unchanged.
- Permanently delete saved quotations after an explicit confirmation.
- Preview the generated customer PDF in the app before downloading. Downloading the PDF updates the workflow status to **Quote Sent**.
- Track **Quote Sent**, **Quote Accepted** and **Quote Declined**. A reason is required for declined quotations. The dashboard shows won and lost tallies.
- Customise costing Excel downloads by selecting worksheets, columns and rows. The workbook retains formulas and hides unselected items rather than deleting formula support cells.
- Updated SQLite migration adds the loss-reason field to existing installations. Back up the database before upgrading.

## PDF preview compatibility (v1.9)

The quotation preview is rendered page-by-page with PyMuPDF and displayed as images in Streamlit. This avoids Chrome policies that can block embedded `data:application/pdf` iframes. The downloaded PDF remains the original vector PDF for printing and sharing.


## Version 2.0 — PDF preview dependency repair

- Corrected the Windows launcher version marker. Existing installations that previously skipped dependency installation will now install/check the current `requirements.txt` on the next launch.
- The requirements include `PyMuPDF`, which provides the `fitz` module used to render PDF pages as preview images.
- If the preview still reports a missing `fitz` module, close Streamlit, run `start_app.bat` again, and check that the package installation completes without errors. The app will continue to provide the PDF download even if preview rendering is unavailable.

## Persistent Supabase PostgreSQL storage

The app uses Supabase PostgreSQL when `SUPABASE_DB_URL` is present in Streamlit Secrets. If that secret is absent, local development continues to use `data/costing_quotes.db` (SQLite).

### Streamlit Cloud

In **App → Settings → Secrets**, configure:

```toml
SUPABASE_DB_URL = "postgresql://postgres:YOUR_PASSWORD@YOUR_HOST:5432/postgres"
```

Use the connection URI from the Supabase dashboard. Prefer the Session Pooler URI if a direct connection is not reachable from your environment. Keep this secret out of GitHub.

### One-time SQLite migration

1. Make a backup copy of `data/costing_quotes.db` before doing anything else.
2. On a computer with Python installed, install the packages in `requirements.txt`.
3. Set `SUPABASE_DB_URL` in your local shell to the same database URI stored in Streamlit Secrets. Do not put the URI in source code.
4. Run `python migrate_sqlite_to_supabase.py` from this folder.
5. Confirm the printed counts against the source database and check the Supabase Table Editor before relying on the app.
6. Commit and push the code changes to GitHub, then wait for Streamlit Cloud to redeploy.

The migration script refuses to run if the target Supabase database already contains quotations. It imports settings, catalogue items, DatecPay tiers, price history and saved quotations. It does not delete the local SQLite backup.
