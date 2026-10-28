---
name: dcf-builder
description: Build a universal DCF workbook (Quarterly, DCF, Valuation Summary, Sensitivity) for any US-listed ticker from SEC filings, with sourced inputs, written justifications, peer forward P/E and a bond-based cost of debt. Also refreshes an existing DCF.
---

# DCF Builder

Builds `<TICKER>_DCF.xlsx` from a fixed template: a universal version of a full Airbnb (ABNB) DCF model. The template is the model. **Never change its formulas.** Only fill the blue input cells and the source/justification cells.

Use it when the user says "DCF for COST", "value NVDA", "build the DCF for…", or "refresh my DCF" (refresh mode is at the end). For manual mode, give them the blank template (Step 0) and stop.

## House rules (non-negotiable)

- **Colors:** Blue = input you edit | Black = formula | Green = link from another sheet | Yellow highlight = key assumption. These come from the template. Don't restyle.
- **Every typed input gets a source.** On Quarterly the source goes in column L of the same row. On DCF it goes in C (key assumptions), M (revenue bridge) or Q (cost of capital). On Valuation Summary it goes in C or M. Give each source a hyperlink to the exact document (EDGAR filing, 8-K Ex. 99.1, FRED, Kroll, stockanalysis.com…).
- **Justifications explain why:** short, quote the number and name the document. Example: `Airbnb Q2'26 shareholder letter outlook "at least 35.5%"`. Or: `Highest full-year margin FY23–FY25 – a level already proven, not a new high`.
- **Units:** $ millions, shares in millions. Expenses are entered as positive numbers. Capex is negative. Working-capital changes keep their cash-flow sign.
- **Methodology is fixed** (it is part of the model's design):
  - Adj. EBITDA = operating income + D&A + SBC + other non-recurring add-backs.
  - Taxes are charged on EBIT (EBITDA − D&A), and D&A is added back once.
  - FY+1 revenue = YTD actual + guided quarter + rest of year.
  - Margin path rises evenly from the guidance floor to the best FY margin of the last 3 years by FY+5.
  - Capex % = 3-year average. NWC % = 2-year average.
  - Cost of debt = principal-weighted YTM of the notes outstanding.
  - Blended target = DCF, P/E and EV/FCF, ⅓ each, rounded with MROUND to $5.
- Excel compatibility: the risk-free cell `DCF!P5` pulls FRED live only on Windows Excel. On Mac Excel, Google Sheets and LibreOffice it uses the fallback in `DCF!P9`, so **always refresh P9 and its date in Q9**.

## Step 0: Set up the kit

Write the three scripts from the appendix into a working folder exactly as given: `build_dcf_template.py`, `fill_dcf.py`, `sec_pull.py`. Then build the blank template:

```bash
python3 build_dcf_template.py DCF_Template.xlsx
```

Recalculate with LibreOffice (`soffice --headless --convert-to xlsx --outdir recalculated/ FILE.xlsx`), or with an xlsx skill's recalc script if your environment has one.

## Step 1: Identify the company

Get the ticker and the valuation date (today unless the user gives one). Confirm the company name from SEC. If it files a 20-F or 40-F (foreign issuer, often IFRS), tell the user before continuing. The template still works, but the YTD columns may be semiannual.

## Step 2: Pull the filings

```bash
python3 sec_pull.py TICKER data/ --email <user's contact email>
```

SEC's fair-access policy asks for a contact email in the User-Agent. Ask the user for one the first time; don't invent one.

This writes `filings.json` (latest 10-K and the newer 10-Q, with URLs), the IS/BS/CF statements as CSV, `statements.md` (read this), and `xbrl_facts.json`, which you use to cross-check FY totals. If the shell can't reach sec.gov, fetch the same EDGAR pages with WebFetch.

## Step 3: Fill the Quarterly sheet

Columns: **B = FY-2, C = FY-1, D = FY0** (all from the latest 10-K, which shows 3 years of IS and CF). **F = YTD prior year, G = YTD current year** (both from the latest 10-Q's year-to-date columns). Never use the 3-month columns for F and G. I (TTM) is a formula.

| Cells | What | Notes |
|---|---|---|
| B3, B4 | Company name, ticker | |
| D7, G7 | FY0 period end, latest 10-Q period end | `{"date": "YYYY-MM-DD"}` |
| D8, G8 | 10-K URL, 10-Q URL (plain URL text) | Row-6 headers hyperlink to these |
| G9 | Fiscal quarters in YTD (0–3) | 0 if the 10-K is the latest filing; then leave F and G blank |
| 11 | Revenue | |
| 12 | Cost of revenue | |
| 14 | SG&A total | Sum of S&M + G&A + ops & support and other SG&A-type lines. List what was summed in L14 |
| 15 | R&D | Product development / technology & development |
| 16 | Other operating expense (income) | Impairments, restructuring, other lines above operating income |
| 19 | Operating income as reported | Check: row 20 must be 0 in every column |
| 21 | Interest expense (positive) | If it's buried in "other", enter 0 and say so in L21 |
| 22 | Other non-operating, net, incl. interest income | Income positive. Pre-tax income must tie |
| 24 | Income tax provision | A benefit is negative |
| 26 | Net income as reported | Check: row 27 must be 0 |
| 31 | D&A | CF statement. If it isn't a separate line, use the MD&A Adj. EBITDA reconciliation or the earnings release and cite it |
| 32 | SBC | CF statement |
| 33 | Other add-backs | Items in the company's Adj. EBITDA reconciliation not covered above. 0 if the company reports no Adj. EBITDA |
| 35 | Adj. EBITDA as reported (optional) | Check: row 36 must be 0 when filled |
| 39 | CFO | |
| 40 | Capex (negative) | Purchases of PP&E. If not on the CF face, use the FCF reconciliation |
| 42 (B:D only) | Changes in operating assets & liabilities, net | Sum of the CF working-capital lines, cash-flow sign |
| G45–G48 | Cash, ST investments, current debt, LT debt | Latest balance sheet. Exclude operating leases |
| G51 | Diluted weighted-avg. shares, latest 3-month period | 10-Q IS |
| G54 | Prior-year revenue of the next (guided) quarter | 3-month column of last year's 10-Q. For a guided Q4 = prior FY − prior 9M YTD |
| A77:F84 | Debt schedule: name, principal, coupon, maturity, price | See Step 4 |

Unit check: convert filings reported in thousands or billions. Cross-check D11, D18 and D25 against `xbrl_facts.json`.

## Step 4: Debt schedule (cost of debt)

- From the latest 10-K/10-Q **Debt** note, enter one row per note: name, principal ($mm), coupon (as a fraction, e.g. 0.044), and maturity date.
- Price (% of par): the last trade on FINRA TRACE (finra-markets.morningstar.com → Bond Center → search the issuer). Use the browser if WebFetch can't render it. The YTM is computed with `YIELD()` off the valuation date in DCF!P2.
- If there's no price, leave it blank. YTM then falls back to the coupon, and L must say so.
- For convertible notes, the convert's price isn't a straight-debt yield. Leave the price blank and type the note's effective interest rate in the coupon cell, saying so in L.
- With no notes listed (bank loans only, or no debt), the DCF falls back to interest expense ÷ debt.

## Step 5: Market inputs (DCF sheet)

| Cell | Input | Source |
|---|---|---|
| P2 | Valuation date | |
| B13 | Share price at close on P2 | IBKR connector (`get_price_snapshot` or price history) or Yahoo. State the date in C13 |
| P9 / Q9 | 10Y Treasury (DGS10) latest close as a fraction, plus its date | `curl 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10&cosd=<date − 10 days>'`. Take the last non-empty value |
| P6 / Q6 | Beta (5Y monthly) | Yahoo Finance statistics page, cite the URL |
| P7 / Q7 | Equity risk premium | Kroll's current recommended U.S. ERP. Verify it is still current |
| B8 / C8 | Terminal growth | Default logic: Fed longer-run real GDP (latest SEP median) + 2.0% inflation. Verify the SEP number and cite it |

## Step 6: Guidance and the growth path

- Use the latest earnings release or shareholder letter (8-K Ex. 99.1), and cite the URL.
- Next-quarter revenue guidance low/high goes in **L7/L8** (M7/M8 sources). If there's no guidance, leave them blank: L9 then grows the prior-year quarter at the YTD rate. Say so in M7.
- **B14**: current-year Adj. EBITDA margin guidance (floor), quoted in C14. If there's no margin guidance, type the TTM margin (`Quarterly!I60` value) and write "No margin guidance – TTM Adj. EBITDA margin" in C14.
- **L13**: consensus revenue for the current FY (cite the source). M13 shows the gap. Flag it if it's above 3%.
- **L14** (growth for FY+2E) and **L15** (annual step-down) are key assumptions. Anchor them to guidance and consensus, explain the reasoning in M14/M15, and write the revenue story in **H19**. Explain the margin path in **H21**.

## Step 7: Peers and forward P/E (Valuation Summary)

- Choose 2–5 competitors. Start with those named in the 10-K, Item 1 "Competition", then same-industry comparables. Explain who and why in C2.
- K2:K6: peer name, hyperlinked to the peer's latest earnings release or 10-Q. L2:L6: forward P/E from `stockanalysis.com/stocks/<ticker>/statistics/` ("Forward PE"). M2:M6: `stockanalysis.com (S&P)` hyperlinked to that page.
- **B3**: consensus EPS for FY+2 (the label in A3 shows the year). Source it from MarketWatch analyst estimates or stockanalysis.com forecast, and cite it in C3.

## Step 8: Fill, recalculate, verify

1. Write `inputs.json` with every value and source note. Text with a link uses `{"text": "...", "url": "..."}`, dates use `{"date": "YYYY-MM-DD"}`, and numbers are plain.
2. Run `python3 fill_dcf.py DCF_Template.xlsx inputs.json TICKER_DCF.xlsx`. It refuses to overwrite formulas.
3. Recalculate, then verify. Every item must pass before delivery:
   - Zero formula errors.
   - Quarterly rows 20, 27 and 36 are 0 (or blank) in every column.
   - FY0 revenue, operating income and net income match XBRL companyfacts.
   - WACC > terminal growth, and WACC is between 6% and 14% (explain if outside).
   - Terminal value as % of EV (DCF!B50): flag it if above 80%.
   - Sensitivity center cell D8 equals DCF!B47.
   - Every blue cell is filled or intentionally blank, with the reason in its source cell.
   - Every source cell has a URL.
4. Render the DCF sheet to an image (LibreOffice → PDF → PNG) and look at it for clipped text or wrong colors.

## Step 9: Deliver

Send `TICKER_DCF.xlsx`. If a folder on the user's computer is connected, also save it there. Then summarize in 3–5 lines: implied DCF price, blended price target and upside, the 2–3 assumptions that move the value most (use the Sensitivity sheet), and any flags (TV % of EV, consensus gap, missing guidance, coupon-fallback YTM).

## Refresh mode ("refresh my DCF", "update the rates")

Read the existing file and update only P2, B13/C13, P9/Q9, peer P/Es (L2:L6) and B3. Re-run the verify checklist, then report the old and new implied price and price target. Never touch the formulas.

## Appendix: scripts (write these files exactly)

### `build_dcf_template.py`

```python
#!/usr/bin/env python3
"""Build the universal DCF template (generalized from a full Airbnb (ABNB) DCF model).

Usage:  python3 build_dcf_template.py OUTPUT.xlsx

Sheets: Quarterly (all company inputs) | DCF | Valuation Summary | Sensitivity
Colors: Blue = input you edit | Black = formula | Green = link from another sheet |
        Yellow fill = key assumption.
Every blue input cell is empty in the template; fill_dcf.py writes values into them.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- styles
NAVY, BLUE, GREEN, LINK, GRAY = "1F3864", "0000FF", "00B050", "0563C1", "595959"
F_NAVY = PatternFill("solid", fgColor=NAVY)
F_YEL = PatternFill("solid", fgColor="FFFF00")
F_BASE = PatternFill("solid", fgColor="DDEBF7")
F_SUB = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin")
NUM = '#,##0;\\(#,##0\\);\\-'
NUM1 = '#,##0.0;\\(#,##0.0\\);\\-'
PCT = '0.0%;\\(0.0%\\);\\-'
PCT2 = '0.00%;\\(0.00%\\);\\-'
USD = '\\$#,##0;"($"#,##0\\);\\-'
USD2 = '\\$#,##0.00;"($"#,##0.00\\);\\-'
MULT = '0.0"x";\\(0.0"x"\\);\\-'
DATE = 'yyyy-mm-dd'
DEC3 = '0.000'
DEC2 = '0.00'


def font(color="000000", size=10, bold=False, italic=False, underline=None):
    return Font(name="Arial", size=size, bold=bold, italic=italic, color="FF" + color, underline=underline)


class S:
    """Tiny writer bound to one sheet with a default font size."""

    def __init__(self, ws, size):
        self.ws, self.size = ws, size

    def put(self, ref, value=None, kind="formula", fmt=None, bold=False, yellow=False,
            wrap=False, align=None, italic=False, size=None):
        c = self.ws[ref]
        if value is not None:
            c.value = value
        sz = size or self.size
        color = {"input": BLUE, "formula": "000000", "link": GREEN, "note": GRAY,
                 "url": LINK, "head": "FFFFFF", "label": "000000"}[kind]
        c.font = font(color, sz, bold=bold or kind == "head", italic=italic or kind == "note",
                      underline="single" if kind == "url" else None)
        if fmt:
            c.number_format = fmt
        if yellow:
            c.fill = F_YEL
        if kind == "head":
            c.fill = F_NAVY
        if wrap or align:
            c.alignment = Alignment(wrap_text=wrap, horizontal=align, vertical="center" if wrap else None)
        return c

    def band(self, row, c1, c2, text=None, align=None):
        for col in range(c1, c2 + 1):
            self.put(f"{get_column_letter(col)}{row}", kind="head")
        if text is not None:
            self.put(f"{get_column_letter(c1)}{row}", text, kind="head", align=align)


# ======================================================================
# QUARTERLY  (universal input sheet)
# ======================================================================
FLOW_COLS = ["B", "C", "D", "F", "G"]      # FY-2, FY-1, FY0, YTD prior year, YTD current year
FY_COLS = ["B", "C", "D"]


def build_quarterly(wb):
    ws = wb.active
    ws.title = "Quarterly"
    q = S(ws, 9)
    widths = {"A": 52, "B": 11, "C": 11, "D": 11, "E": 2, "F": 11, "G": 11, "H": 2, "I": 11, "J": 2,
              "K": 50, "L": 44}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.sheet_view.zoomScale = 115

    q.put("A1", '=IF($B$3="","Company",$B$3)&" – annual and year-to-date financials ($mm, except per share; shares in mm)"',
          size=12, bold=True).font = font(NAVY, 12, bold=True)
    q.put("A2", "Every blue number is typed from the SEC filing linked in its column header (click the header). "
                "One 10-K gives FY-2 to FY0; the latest 10-Q gives both YTD columns and the balance sheet. "
                "TTM = FY0 + YTD current − YTD prior. Black = formulas.", kind="note")
    q.put("A3", "Company name", kind="label"); q.put("B3", kind="input")
    q.put("A4", "Ticker", kind="label"); q.put("B4", kind="input")

    # header band row 6
    q.band(6, 1, 12)
    q.put("A6", "Metric", kind="head")
    for col, off, dflt in (("B", -24, "FY-2"), ("C", -12, "FY-1"), ("D", 0, "FY0")):
        q.put(f"{col}6", f'=HYPERLINK($D$8,IF($D$7="","{dflt}","FY"&YEAR({col}7)))', kind="head", align="center")
    q.put("F6", '=HYPERLINK($G$8,IF($G$7="","YTD prior yr","YTD Q"&$G$9&" FY"&RIGHT(YEAR($D$7)),2)))', kind="head", align="center")
    q.put("G6", '=HYPERLINK($G$8,IF($G$7="","YTD current","YTD Q"&$G$9&" FY"&RIGHT(YEAR($D$7)+1,2)))', kind="head", align="center")
    q.put("I6", '=IF($G$7="","TTM","TTM "&TEXT($G$7,"mmm-yy"))', kind="head", align="center")
    q.put("K6", "Where to find it (manual entry guide)", kind="head")
    q.put("L6", "Source / notes", kind="head")
    # F6 formula above has a stray paren risk – rebuild cleanly
    ws["F6"].value = '=HYPERLINK($G$8,IF($G$7="","YTD prior yr","YTD Q"&$G$9&" FY"&RIGHT(YEAR($D$7),2)))'

    q.put("A7", "Period end date", kind="label")
    q.put("B7", '=IF($D$7="","",EDATE($D$7,-24))', fmt=DATE)
    q.put("C7", '=IF($D$7="","",EDATE($D$7,-12))', fmt=DATE)
    q.put("D7", kind="input", fmt=DATE)
    q.put("F7", '=IF($G$7="","",EDATE($G$7,-12))', fmt=DATE)
    q.put("G7", kind="input", fmt=DATE)
    q.put("I7", '=G7', fmt=DATE)
    q.put("K7", "FY0 = fiscal year of the latest 10-K. YTD = latest 10-Q period end (leave blank if no 10-Q yet).", kind="note")
    q.put("A8", "Source filing URL (headers in row 6 link here)", kind="label")
    q.put("B8", "same 10-K →", kind="note"); q.put("C8", "same 10-K →", kind="note")
    q.put("D8", kind="input", size=8); q.put("F8", "same 10-Q →", kind="note"); q.put("G8", kind="input", size=8)
    q.put("K8", "EDGAR link to the 10-K (D8) and latest 10-Q (G8). One 10-K carries 3 years of IS and CF.", kind="note")
    q.put("A9", "Fiscal quarters in YTD period (0–3)", kind="label")
    q.put("G9", kind="input", fmt="0")
    q.put("K9", "Q1 10-Q = 1, Q2 = 2, Q3 = 3. Use 0 when the 10-K is the latest filing.", kind="note")

    def section(row, text):
        q.band(row, 1, 12, text)

    def flow_row(r, label, guide, cols=FLOW_COLS, ttm=True, bold=False, formula=None, fmt=NUM):
        q.put(f"A{r}", label, kind="label", bold=bold)
        for col in cols:
            if formula:
                q.put(f"{col}{r}", formula(col), fmt=fmt, bold=bold)
            else:
                q.put(f"{col}{r}", kind="input", fmt=fmt)
        if ttm:
            q.put(f"I{r}", f"=D{r}+G{r}-F{r}", fmt=fmt, bold=bold)
        if guide:
            q.put(f"K{r}", guide, kind="note")

    # ---------------- income statement
    section(10, "Income statement – standardized (LSEG / IBKR layout). Enter expenses as positive numbers.")
    flow_row(11, "Revenue (total)", "IS top line: 'Total revenue' / 'Net sales'.")
    flow_row(12, "Cost of revenue", "IS: cost of revenue / cost of sales / cost of goods sold.")
    flow_row(13, "Gross profit", None, formula=lambda c: f"={c}11-{c}12", bold=True)
    flow_row(14, "Selling/General/Admin. expenses, total",
             "Sum of S&M + G&A (+ ops & support, other SG&A-type lines).")
    flow_row(15, "Research & development", "IS: R&D / product development / technology & development.")
    flow_row(16, "Other operating expense (income), net",
             "Impairments, restructuring, other operating lines above operating income.")
    flow_row(17, "Total operating expense", None, formula=lambda c: f"={c}12+{c}14+{c}15+{c}16")
    flow_row(18, "Operating income", None, formula=lambda c: f"={c}11-{c}17", bold=True)
    flow_row(19, "  Operating income as reported (check)", "IS: 'Income from operations' – proves the lines above tie.")
    flow_row(20, "  Difference (should be 0)", None, ttm=False,
             formula=lambda c: f'=IF({c}19="","",{c}18-{c}19)')
    flow_row(21, "Interest expense", "IS or debt note. Enter positive. 0 if it is buried in 'other' (say so in notes).")
    flow_row(22, "Other non-operating income (expense), net – incl. interest income",
             "IS: interest income + other income/expense. Income positive, expense negative.")
    flow_row(23, "Pre-tax income", None, formula=lambda c: f"={c}18-{c}21+{c}22", bold=True)
    flow_row(24, "Provision for income taxes", "IS: income tax expense (benefit negative).")
    flow_row(25, "Net income", None, formula=lambda c: f"={c}23-{c}24", bold=True)
    flow_row(26, "  Net income as reported (check)", "IS: net income attributable to the company.")
    flow_row(27, "  Difference (should be 0)", None, ttm=False,
             formula=lambda c: f'=IF({c}26="","",{c}25-{c}26)')

    # ---------------- adjusted EBITDA
    section(29, "Adjusted EBITDA – universal build (operating income + D&A + SBC + other add-backs)")
    flow_row(30, "Operating income", None, formula=lambda c: f"={c}18")
    flow_row(31, "Depreciation & amortization", "Cash flow statement, first adjustments under net income.")
    flow_row(32, "Stock-based compensation", "Cash flow statement: stock-based compensation expense.")
    flow_row(33, "Other non-recurring add-backs (0 if none)",
             "Company's Adj. EBITDA reconciliation items not above (e.g. restructuring, legal/tax reserves).")
    flow_row(34, "Adjusted EBITDA", None, formula=lambda c: f"=SUM({c}30:{c}33)", bold=True)
    flow_row(35, "  Adjusted EBITDA as reported (optional check)", "MD&A / earnings release reconciliation. Leave blank if not reported.")
    flow_row(36, "  Difference (should be 0)", None, ttm=False,
             formula=lambda c: f'=IF({c}35="","",{c}34-{c}35)')

    # ---------------- cash flow
    section(38, "Cash flow statement")
    flow_row(39, "Net cash provided by operating activities", "CF statement: operating activities total.")
    flow_row(40, "Capital expenditures (enter negative)", "CF: purchases of property & equipment (+ capitalized software if material).")
    flow_row(41, "Free cash flow (= CFO + capex)", None, formula=lambda c: f"={c}39+{c}40", bold=True)
    flow_row(42, "Changes in operating assets & liabilities, net (cash-flow sign)",
             "CF: sum of the 'changes in operating assets and liabilities' lines. Negative = cash invested.",
             cols=FY_COLS, ttm=False)

    # ---------------- balance sheet
    section(44, "Balance sheet & shares – latest 10-Q (column G; use FY0 10-K if no 10-Q yet)")
    for r, label, guide in ((45, "Cash and cash equivalents", "Balance sheet, current assets."),
                            (46, "Short-term investments / marketable securities", "Balance sheet, current assets."),
                            (47, "Debt – current portion", "Balance sheet: current debt / notes due within 1 year."),
                            (48, "Debt – long-term, net", "Balance sheet: long-term debt (excl. operating leases).")):
        q.put(f"A{r}", label, kind="label"); q.put(f"G{r}", kind="input", fmt=NUM); q.put(f"K{r}", guide, kind="note")
    q.put("A49", "Total debt", kind="label", bold=True); q.put("G49", "=G47+G48", fmt=NUM, bold=True)
    q.put("A50", "Cash + short-term investments", kind="label", bold=True); q.put("G50", "=G45+G46", fmt=NUM, bold=True)
    q.put("A51", "Diluted weighted-avg. shares, latest quarter (mm)", kind="label")
    q.put("G51", kind="input", fmt=NUM1)
    q.put("K51", "10-Q income statement: diluted weighted-average shares for the latest 3-month period.", kind="note")

    # ---------------- guidance bridge
    section(53, "Revenue bridge input (for current-year estimate on the DCF sheet)")
    q.put("A54", "Prior-year revenue of the next (guided) quarter", kind="label")
    q.put("G54", kind="input", fmt=NUM)
    q.put("K54", "Same quarter last year (e.g. guiding Q3'26 → Q3'25 revenue from last year's 10-Q).", kind="note")

    # ---------------- derived metrics
    section(56, "Derived metrics (formulas)")
    def ratio_row(r, label, num, den, cols=("B", "C", "D", "F", "G", "I"), fmt=PCT, sign=""):
        q.put(f"A{r}", label, kind="label")
        for c in cols:
            q.put(f"{c}{r}", f"=IF({c}{den}=0,0,{sign}{c}{num}/{c}{den})", fmt=fmt)
    q.put("A57", "Revenue growth YoY", kind="label")
    q.put("C57", "=IF(B11=0,0,C11/B11-1)", fmt=PCT)
    q.put("D57", "=IF(C11=0,0,D11/C11-1)", fmt=PCT)
    q.put("G57", "=IF(F11=0,0,G11/F11-1)", fmt=PCT)
    ratio_row(58, "Gross margin", 13, 11)
    ratio_row(59, "Operating margin", 18, 11)
    ratio_row(60, "Adj. EBITDA margin", 34, 11)
    ratio_row(61, "Net income margin", 25, 11)
    ratio_row(62, "FCF margin", 41, 11)
    ratio_row(63, "FCF / Adj. EBITDA", 41, 34)
    ratio_row(64, "SBC % of revenue", 32, 11)
    ratio_row(65, "D&A % of revenue", 31, 11)
    ratio_row(66, "Capex % of revenue", 40, 11, sign="-")
    ratio_row(67, "Effective tax rate (tax / pre-tax income)", 24, 23)
    q.put("A68", "NWC investment % of change in revenue", kind="label")
    q.put("C68", "=IF(C11-B11=0,0,-C42/(C11-B11))", fmt=PCT)
    q.put("D68", "=IF(D11-C11=0,0,-D42/(D11-C11))", fmt=PCT)

    section(70, "Values the DCF uses (column D)")
    q.put("A71", "Capex % of revenue – 3-year average", kind="label"); q.put("D71", "=AVERAGE(B66:D66)", fmt=PCT)
    q.put("A72", "NWC % of change in revenue – 2-year average", kind="label"); q.put("D72", "=AVERAGE(C68:D68)", fmt=PCT)
    q.put("A73", "Highest full-year Adj. EBITDA margin (FY-2 to FY0)", kind="label"); q.put("D73", "=MAX(B60:D60)", fmt=PCT)
    q.put("K71", "Feeds DCF capex % (FY+1 to FY+5).", kind="note")
    q.put("K72", "Feeds DCF NWC %. Override in notes if one year is distorted by a one-off.", kind="note")
    q.put("K73", "Feeds the DCF long-term margin target – a level the company has already delivered.", kind="note")

    # ---------------- debt schedule
    section(75, "Debt schedule – notes outstanding (10-K debt footnote; prices from FINRA TRACE). Feeds cost of debt.")
    for col, text in (("A", "Note / facility"), ("B", "Principal ($mm)"), ("C", "Coupon"), ("D", "Maturity"),
                      ("F", "Price (% of par)"), ("G", "YTM"), ("I", "Weight"), ("K", "Where to find it"), ("L", "Source / notes")):
        q.put(f"{col}76", text, kind="label", bold=True).fill = F_SUB
    for col in ("E", "H", "J"):
        ws[f"{col}76"].fill = F_SUB
    for r in range(77, 85):
        q.put(f"A{r}", kind="input")
        q.put(f"B{r}", kind="input", fmt=NUM)
        q.put(f"C{r}", kind="input", fmt=PCT2)
        q.put(f"D{r}", kind="input", fmt=DATE)
        q.put(f"F{r}", kind="input", fmt=DEC2)
        q.put(f"G{r}", f'=IF(B{r}="","",IF(F{r}="",C{r},IFERROR(YIELD(DCF!$P$2,D{r},C{r},F{r},100,2,0),C{r})))', fmt=PCT2)
        q.put(f"I{r}", f'=IF(OR(B{r}="",$B$85=0),"",B{r}/$B$85)', fmt=PCT)
    q.put("K77", "10-K note 'Debt': each note's principal, coupon, maturity.", kind="note")
    q.put("K78", "Price: FINRA TRACE last trade (finra-markets.morningstar.com).", kind="note")
    q.put("K79", "No price → YTM falls back to the coupon.", kind="note")
    q.put("K80", "Converts: use straight-debt yield, not the convert's price.", kind="note")
    q.put("A85", "Total / principal-weighted YTM", kind="label", bold=True)
    q.put("B85", "=SUM(B77:B84)", fmt=NUM, bold=True)
    q.put("G85", "=IF(B85=0,0,SUMPRODUCT(B77:B84,G77:G84)/B85)", fmt=PCT2, bold=True)
    for col in ("B", "G"):
        ws[f"{col}85"].border = Border(top=THIN)

    q.put("A87", "Notes:", kind="label", bold=True)
    q.put("A88", "• Adjusted EBITDA here = operating income + D&A + SBC + other add-backs; the check row ties it to the company's reported figure when one exists.", kind="note")
    q.put("A89", "• Units: $ millions. Convert filings reported in thousands (÷1,000) or billions (×1,000).", kind="note")
    q.put("A90", "• YTD columns are the cumulative 10-Q figures (10-Q cash flow statements are already year-to-date).", kind="note")
    ws.freeze_panes = "B7"
    return ws


# ======================================================================
# DCF
# ======================================================================
def build_dcf(wb):
    ws = wb.create_sheet("DCF")
    d = S(ws, 10)
    ws.sheet_view.showGridLines = False
    for k, v in {"A": 47.5, "B": 13, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11,
                 "I": 14, "J": 14, "K": 16, "L": 12, "M": 43.2, "N": 21.8, "O": 52, "P": 13.8, "Q": 40}.items():
        ws.column_dimensions[k].width = v
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 24

    yr = "YEAR(Quarterly!$D$7)"
    lab = lambda n, suf: f'=IF(Quarterly!$D$7="","FY0{"+" + str(n) if n else ""}{suf}","FY"&({yr}+{n})&"{suf}")'

    d.put("A1", "Discounted Cash Flow (DCF)", size=20, bold=True)
    c = d.put("A2", "=IF(Quarterly!$B$4=\"\",\"TICKER\",Quarterly!$B$4)", kind="link", size=18, bold=True)
    d.put("A3", "Blue = input you edit  |  Black = formula  |  Green = link from another sheet  |  Yellow = key assumptions.",
          kind="note", size=9, italic=False)
    d.put("O2", "Valuation date", kind="label", bold=True)
    d.put("P2", kind="input", fmt=DATE)
    d.put("Q2", "Pricing date for share price, Rf and bond YTMs", kind="note", size=9)

    # key assumptions band
    d.band(4, 1, 7, "KEY ASSUMPTIONS ($ in millions, except per share)")
    d.band(4, 8, 14)
    d.put("H4", lab(1, "E"), kind="head")
    d.put("M4", "Source/ Justification", kind="head")
    d.band(4, 15, 16, "COST OF EQUITY (CAPM)")

    rows = [
        (5, "Base year revenue (FY0)", "=Quarterly!D11", "link", USD, False, "Latest actual fiscal-year revenue"),
        (6, "Tax rate", "=Quarterly!D67", "link", PCT, False, "FY0 effective rate (tax ÷ pre-tax income), applied to EBIT"),
        (7, "WACC (discount rate)", "=P25", "formula", PCT, False, "Build separately (CAPM cost of equity + after-tax cost of debt)"),
        (8, "Terminal growth rate", None, "input", PCT, True, "[Justify: long-term nominal GDP growth – cite source]"),
        (9, "Cash + short-term investments", "=Quarterly!G50", "link", USD, False, "Latest 10-Q balance sheet"),
        (10, "Total debt", "=Quarterly!G49", "link", USD, False, "Latest 10-Q balance sheet"),
        (11, "Net cash", "=B9-B10", "formula", USD, False, None),
        (12, "Diluted shares outstanding (mm)", "=Quarterly!G51", "link", NUM1, False, "Latest-quarter diluted weighted-average shares (10-Q)"),
        (13, "Current share price ($)", None, "input", USD2, False, "[Closing price on valuation date – source]"),
        (14, "Current-year Adj. EBITDA margin guidance (floor)", None, "input", PCT, True, "[Company outlook – quote and cite]"),
        (15, "Long-term Adj. EBITDA margin target (reached FY+5)", "=Quarterly!D73", "link", PCT, True,
         "Highest full-year Adj. EBITDA margin of the last 3 fiscal years – a level already proven, not a new high."),
    ]
    for r, label, f, kind, fmt, yel, note in rows:
        d.put(f"A{r}", label, kind="label")
        d.put(f"B{r}", f, kind=kind, fmt=fmt, yellow=yel)
        if note:
            d.put(f"C{r}", note, kind="note", size=9)
    ws.merge_cells("C15:G15"); ws["C15"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[15].height = 39.75
    d.put("A14", '="FY"&IF(Quarterly!$D$7="","0+1",YEAR(Quarterly!$D$7)+1)&"E Adj. EBITDA margin guidance (floor)"', kind="label")

    # revenue bridge H/L/M
    br = [
        (5, "YTD revenue (actual, latest 10-Q)", "=Quarterly!G11", "link", "Latest 10-Q (Quarterly sheet, col G)"),
        (6, "Prior-year revenue of guided quarter", "=Quarterly!G54", "link", "Last year's 10-Q"),
        (7, "Next-quarter revenue guidance – low", None, "input", "[Company outlook – cite]"),
        (8, "Next-quarter revenue guidance – high", None, "input", "[Same]"),
        (9, "Guided quarter revenue (midpoint; YTD growth if no guidance)",
         "=IF(COUNT(L7:L8)=2,AVERAGE(L7:L8),L6*(1+Quarterly!G57))", "formula", None),
        (10, "Guided quarter implied YoY growth", "=IF(L6=0,0,L9/L6-1)", "formula", "vs same quarter last year"),
        (11, "Rest of year = prior-year remainder × (1 + guided growth)",
         "=(Quarterly!D11-Quarterly!F11-Quarterly!G54)*(1+L10)", "formula",
         "Assumes the remaining quarters grow at the rate management guided"),
        (12, None, "=L5+L9+L11", "formula", None),
        (13, "  Cross-check: analyst consensus revenue", None, "input", None),
        (14, "Revenue growth, FY+2E (first year after guidance)", None, "input", "[Justify – analyst assumption / consensus]"),
        (15, "Annual growth step-down after FY+2E", None, "input", "[Justify – e.g. growth slows 1 pt per year]"),
    ]
    for r, label, f, kind, note in br:
        if label:
            d.put(f"H{r}", label, kind="label")
        fmt = PCT if r in (10, 14, 15) else USD
        d.put(f"L{r}", f, kind=kind, fmt=fmt, yellow=r in (14, 15))
        if note:
            d.put(f"M{r}", note, kind="note", size=9)
    d.put("H12", '="FY"&IF(Quarterly!$D$7="","0+1",YEAR(Quarterly!$D$7)+1)&"E revenue"', kind="label", bold=True)
    ws["L12"].font = font("000000", 10, bold=True)
    d.put("M13", '=IF(N(L13)=0,"",(L12-L13)/L13)', fmt=PCT)

    # cost of equity / debt / WACC (O/P/Q)
    fred = ('VALUE(TRIM(RIGHT(SUBSTITUTE(TRIM(CLEAN(_xlfn.WEBSERVICE("https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10&cosd="'
            '&TEXT(TODAY()-10,"yyyy-mm-dd")))),",",REPT(" ",200)),200)))/100')
    d.put("O5", "Risk-free rate (10Y Treasury, FRED DGS10)", kind="label")
    d.put("P5", f"=IFERROR({fred},P9)", fmt=PCT2)
    d.put("Q5", "FRED DGS10 – live pull on Windows Excel; Mac uses the fallback below", kind="url", size=9)
    ws["Q5"].hyperlink = "https://fred.stlouisfed.org/series/DGS10"
    d.put("O6", "BETA", kind="label"); d.put("P6", kind="input", fmt=DEC2)
    d.put("Q6", "[Beta source]", kind="note", size=9)
    d.put("O7", "Equity risk premium", kind="label"); d.put("P7", kind="input", fmt=PCT2, yellow=True)
    d.put("Q7", "[ERP source – e.g. Kroll recommended U.S. ERP]", kind="note", size=9)
    d.put("O8", "Cost of equity  (Rf + Beta x ERP)", kind="label", bold=True)
    d.put("P8", "=P5+P6*P7", fmt=PCT2, bold=True)
    d.put("O9", "  10Y yield fallback (last FRED close; used on Mac)", kind="note", size=9)
    d.put("P9", kind="input", fmt=PCT2)
    d.put("Q9", "[FRED DGS10 close and its date]", kind="note", size=9)

    d.band(10, 15, 16, "COST OF DEBT")
    d.put("O11", "Principal-weighted YTM on notes outstanding", kind="label")
    d.put("P11", "=Quarterly!G85", kind="link", fmt=PCT2)
    d.put("Q11", "Debt schedule on the Quarterly sheet (10-K debt note + FINRA TRACE)", kind="note", size=9)
    d.put("O12", "Pre-tax cost of debt used", kind="label")
    d.put("P12", "=IF(P11>0,P11,IF(B10=0,0,Quarterly!D21/B10))", fmt=PCT2)
    d.put("Q12", "Falls back to FY0 interest expense ÷ total debt if no notes are listed", kind="note", size=9)
    d.put("O13", "Tax rate", kind="label"); d.put("P13", "=B6", fmt=PCT)
    d.put("O14", "After-tax cost of debt", kind="label", bold=True); d.put("P14", "=P12*(1-P13)", fmt=PCT2, bold=True)

    d.band(16, 15, 16, "CAPITAL WEIGHT AND WACC")
    d.put("O17", "Equity value – market cap ($mm)", kind="label"); d.put("P17", "=B13*B12", fmt=USD)
    d.put("O18", "Debt ($mm)", kind="label"); d.put("P18", "=B10", fmt=USD)
    d.put("O19", "Equity weight", kind="label"); d.put("P19", "=IF(P17+P18=0,0,P17/(P17+P18))", fmt=PCT)
    d.put("O20", "Debt weight", kind="label"); d.put("P20", "=IF(P17+P18=0,0,P18/(P17+P18))", fmt=PCT)
    d.band(22, 15, 16, "WACC")
    d.put("O23", "Equity contribution  (Equity weight x Cost of equity)", kind="label"); d.put("P23", "=P19*P8", fmt=PCT2)
    d.put("O24", "Debt contribution  (Debt weight x after-tax cost of debt)", kind="label"); d.put("P24", "=P20*P14", fmt=PCT2)
    d.put("O25", "WACC", kind="label", bold=True); d.put("P25", "=P23+P24", fmt=PCT2, bold=True)
    ws["P25"].border = Border(top=THIN)

    # FCF projection
    d.band(17, 1, 7, "FREE CASH FLOW PROJECTION ($ in millions)")
    cols = ["B", "C", "D", "E", "F", "G"]
    d.put("B17", lab(0, "A"), kind="head", align="right")
    for i, col in enumerate(cols[1:], start=1):
        d.put(f"{col}17", lab(i, "E"), kind="head", align="right")

    def row(r, label, fs, fmt=USD, bold=False, kinds=None, yellow=False):
        d.put(f"A{r}", label, kind="label", bold=bold)
        for i, col in enumerate(cols):
            f = fs(col, i)
            if f is None:
                continue
            k = (kinds or {}).get(col, "formula")
            d.put(f"{col}{r}", f, kind=k, fmt=fmt, bold=bold)
    prev = {c: cols[i - 1] for i, c in enumerate(cols) if i}

    row(18, "Discount period (years)", lambda c, i: "0" if i == 0 else f"={prev[c]}18+1", fmt="0", kinds={"B": "input"})
    ws["B18"].value = 0
    row(19, "Revenue growth %",
        lambda c, i: ["=Quarterly!D57", "=IF(B20=0,0,(C20-B20)/B20)", "=$L$14"][i] if i < 3 else f"={prev[c]}19-$L$15",
        fmt=PCT, kinds={"B": "link"})
    row(20, "Revenue", lambda c, i: ["=B5", "=L12"][i] if i < 2 else f"={prev[c]}20*(1+{c}19)", bold=True)
    row(21, "Adj. EBITDA margin %",
        lambda c, i: ["=IF(B20=0,0,B22/B20)", "=B14"][i] if i < 2 else f"=$C$21+($B$15-$C$21)*{i - 1}/4", fmt=PCT)
    row(22, "Adj. EBITDA", lambda c, i: "=Quarterly!D34" if i == 0 else f"={c}20*{c}21", kinds={"B": "link"})
    row(23, "D&A % of revenue",
        lambda c, i: ["=Quarterly!D65", "=Quarterly!I65"][i] if i < 2 else "=$C$23", fmt=PCT, kinds={"B": "link", "C": "link"})
    row(24, "Less: D&A", lambda c, i: f"=-{c}20*{c}23")
    row(25, "EBIT", lambda c, i: f"={c}22+{c}24", bold=True)
    row(26, "Less: taxes on EBIT", lambda c, i: f"=-MAX({c}25,0)*$B$6")
    row(27, "Net Operating Profit After Tax", lambda c, i: f"={c}25+{c}26", bold=True)
    row(28, "Plus: D&A", lambda c, i: f"=-{c}24")
    row(29, "Capex % of revenue",
        lambda c, i: "=Quarterly!D66" if i == 0 else "=Quarterly!$D$71", fmt=PCT, kinds={c: "link" for c in cols})
    row(30, "Less: capital expenditures", lambda c, i: f"=-{c}20*{c}29")
    row(31, "NWC investment % of change in revenue", lambda c, i: None if i == 0 else "=Quarterly!$D$72",
        fmt=PCT, kinds={c: "link" for c in cols})
    row(32, "Less: increase in net working capital", lambda c, i: None if i == 0 else f"=-({c}20-{prev[c]}20)*{c}31")
    row(33, "Unlevered free cash flow (UFCF)", lambda c, i: None if i == 0 else f"={c}27+{c}28+{c}30+{c}32", bold=True)
    row(34, "Discount factor", lambda c, i: None if i == 0 else f"=1/(1+$B$7)^{c}18", fmt=DEC3)
    row(35, "PV of UFCF", lambda c, i: None if i == 0 else f"={c}33*{c}34", bold=True)
    for col in cols[1:]:
        ws[f"{col}33"].border = Border(top=THIN)

    # justification notes (H column)
    d.put("H19", "[Revenue path: current-year guidance, then FY+2E growth stepping down each year – cite guidance]",
          kind="note", size=9, wrap=True)
    ws.merge_cells("H19:N19"); ws.row_dimensions[19].height = 25
    d.put("H21", "[Margin path: guidance floor rising evenly to the proven long-term level by FY+5 – explain why]",
          kind="note", size=9)
    d.put("H23", "D&A held at the TTM % of revenue", kind="note", size=9)
    d.put("H26", "Taxes on EBIT (EBITDA − D&A) – keeps the D&A tax shield; D&A is added back once below", kind="note", size=9)
    d.put("H29", "3-year average capex % of revenue (Quarterly sheet)", kind="note", size=9)
    d.put("H31", "2-year average NWC investment per $ of revenue growth (Quarterly sheet)", kind="note", size=9)
    d.put("H34", "Formula: 1 ÷ (1 + discount rate) ^ years.  e.g. 0.826 means $1 received in year 2 is worth $0.826 today",
          kind="note", size=9)

    # valuation
    d.band(38, 1, 2, "VALUATION ($ in millions, except per share)")
    val = [
        (39, "Sum of PV of UFCF (FY+1 to FY+5)", "=SUM(C35:G35)", USD, False),
        (40, "Terminal year UFCF (FY+5 x (1+g))", "=G33*(1+B8)", USD, False),
        (41, "Terminal value (Gordon growth)", "=IF(B7>B8,B40/(B7-B8),IF(B7=0,0,NA()))", USD, False),
        (42, "PV of terminal value", "=B41*G34", USD, False),
        (43, "Enterprise value", "=B39+B42", USD, True),
        (44, "Plus: net cash", "=B11", USD, False),
        (45, "Equity value", "=B43+B44", USD, True),
        (46, "Diluted shares (mm)", "=B12", NUM1, False),
        (47, "Implied share price ($)", "=IF(B46=0,0,B45/B46)", USD2, True),
        (48, "Current share price ($)", "=B13", USD2, False),
        (49, "Upside / (downside)", "=IF(B48=0,0,B47/B48-1)", PCT, False),
        (50, "  Terminal value as % of enterprise value", "=IF(B43=0,0,B42/B43)", PCT, False),
    ]
    for r, label, f, fmt, b in val:
        d.put(f"A{r}", label, kind="label", bold=b)
        d.put(f"B{r}", f, fmt=fmt, bold=b)
    for r in (43, 45, 47):
        ws[f"B{r}"].border = Border(top=THIN)
    ws["B47"].fill = F_BASE
    ws["A50"].font = font(GRAY, 9, italic=True); ws["B50"].font = font(GRAY, 9, italic=True)
    return ws


# ======================================================================
# VALUATION SUMMARY
# ======================================================================
def build_vs(wb):
    ws = wb.create_sheet("Valuation Summary")
    v = S(ws, 10)
    ws.sheet_view.zoomScale = 125
    for k, w in {"A": 40.5, "B": 12, "C": 11, "D": 9, "K": 22, "L": 9, "M": 30}.items():
        ws.column_dimensions[k].width = w
    fy2 = 'IF(Quarterly!$D$7="","FY0+2",YEAR(Quarterly!$D$7)+2)'
    v.band(1, 1, 9)
    v.put("A1", f'="Valuation Drivers for FY"&{fy2}', kind="head")
    for col, t in (("K", "Company"), ("L", "Fwd P/E"), ("M", "Valuation data")):
        v.put(f"{col}1", t, kind="head", align="center", size=9, wrap=True)
    for r in range(2, 7):
        v.put(f"K{r}", kind="url")
        v.put(f"L{r}", kind="input", fmt="0.0")
        v.put(f"M{r}", kind="url")
    v.put("A2", f'="Target P/E on FY"&{fy2}&"E EPS"', kind="label")
    v.put("B2", "=IF(COUNT(L2:L6)=0,0,AVERAGE(L2:L6))", fmt="0.0", yellow=True)
    v.put("C2", "[Average forward P/E of named competitors – who and why]", kind="note")
    v.put("A3", "Avg. EPS Estimate", kind="label")
    v.put("B3", kind="input", fmt=DEC2, yellow=True)
    v.put("C3", "[Consensus EPS source for FY+2]", kind="note")
    v.put("A5", "Enterprise Value", kind="label"); v.put("B5", "=DCF!P17-DCF!B11", kind="link", fmt='"$"#,##0.00')
    v.put("A6", "FCF / Adj. EBITDA (conversion)", kind="label")
    v.put("B6", "=IF(Quarterly!I34=0,0,Quarterly!I41/Quarterly!I34)", kind="link", fmt=PCT)
    v.put("C6", "Actual TTM (trailing twelve months) free cash flow ÷ Adj. EBITDA", kind="label")
    v.put("A7", "Free Cash Flow", kind="label"); v.put("B7", "=DCF!D22*B6", kind="link", fmt='"$"#,##0.00')
    v.put("C7", "FY+2E Adj. EBITDA × TTM conversion", kind="label")
    v.put("A8", '="Target EV / FCF on FY"&' + fy2 + '&"E FCF"', kind="label")
    v.put("B8", "=IF(Quarterly!I41=0,0,B5/Quarterly!I41)", kind="link", fmt="0.0")
    v.put("C8", "Company's own current EV ÷ TTM FCF – assumes no multiple expansion", kind="label")

    v.band(13, 1, 9)
    v.put("A13", "Valuation summary (12-month target)", kind="head")
    v.put("C13", "Upside", kind="head", align="center"); v.put("D13", "Weight", kind="head")
    items = [(14, "DCF", "=DCF!B47", "link"), (15, "P/E", "=B3*B2", "formula"),
             (16, "EV / FCF", "=IF(DCF!B12=0,0,(B8*B7+DCF!B11)/DCF!B12)", "formula")]
    for r, label, f, k in items:
        v.put(f"A{r}", label, kind="label")
        v.put(f"B{r}", f, kind=k, fmt='"$"#,##0.00')
        v.put(f"C{r}", f"=IF(DCF!$B$13=0,0,(B{r}-DCF!$B$13)/DCF!$B$13)", fmt=PCT2)
        v.put(f"D{r}", 1/3, kind="input", fmt=DEC2)
    v.put("A17", "Blended value", kind="label", bold=True)
    v.put("B17", "=IF(SUM(D14:D16)=0,0,SUMPRODUCT(B14:B16,D14:D16)/SUM(D14:D16))", fmt='"$"#,##0.00')
    v.put("C17", "=IF(DCF!$B$13=0,0,(B17-DCF!$B$13)/DCF!$B$13)", fmt=PCT2)
    v.put("A18", "Price target (rounded to $5)", kind="label", bold=True)
    v.put("B18", "=MROUND(B17,5)", fmt='"$"#,##0', bold=True).fill = PatternFill("solid", fgColor="D9EAD3")
    v.put("C18", "=IF(DCF!$B$13=0,0,(B18-DCF!$B$13)/DCF!$B$13)", fmt=PCT2, bold=True).fill = PatternFill("solid", fgColor="D9EAD3")
    v.put("A19", '="Implied P/E on FY"&' + fy2 + '&"E at target"', kind="label")
    v.put("B19", "=IF(B3=0,0,B18/B3)", fmt=DEC2)
    return ws


# ======================================================================
# SENSITIVITY
# ======================================================================
def build_sens(wb):
    ws = wb.create_sheet("Sensitivity")
    s = S(ws, 10)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    for col in "BCDEF":
        ws.column_dimensions[col].width = 13
    s.put("A1", "Sensitivity Analysis", size=14, bold=True)
    s.put("A2", "Axes are centered on the base case in the DCF sheet and update automatically. "
                "Each cell recomputes the full DCF (no Data Table needed).", kind="note", size=9)
    steps = [-0.01, -0.005, 0, 0.005, 0.01]
    pv = "SUMPRODUCT(DCF!$C$33:$G$33/(1+{w})^DCF!$C$18:$G$18)"
    tv = "DCF!$G$33*(1+{g})/({w}-{g})/(1+{w})^DCF!$G$18"
    for top, title, eq in ((4, "IMPLIED SHARE PRICE ($)", True), (13, "ENTERPRISE VALUE ($ in millions)", False)):
        s.band(top, 1, 6, title)
        hr = top + 1
        s.put(f"A{hr}", "WACC (x) / Terminal Growth Rate (y)", kind="head")
        for j, col in enumerate("BCDEF"):
            s.put(f"{col}{hr}", f"=DCF!$B$8+({steps[j]})", kind="link", fmt=PCT, align="right")
            ws[f"{col}{hr}"].fill = F_NAVY
            ws[f"{col}{hr}"].font = font("FFFFFF", 10, bold=True)
        for i in range(5):
            r = hr + 1 + i
            s.put(f"A{r}", f"=DCF!$B$7+({steps[i]})", kind="link", fmt=PCT, align="right")
            for col in "BCDEF":
                w, g = f"$A{r}", f"{col}${hr}"
                core = pv.format(w=w) + "+" + tv.format(w=w, g=g)
                if eq:
                    body = f"({core}+DCF!$B$44)/DCF!$B$12"
                    fmt = USD2
                else:
                    body = core
                    fmt = USD
                s.put(f"{col}{r}", f"=IF(DCF!$B$12=0,0,IF({w}>{g},{body},NA()))", fmt=fmt)
        mid = hr + 3
        ws[f"D{mid}"].fill = F_BASE
        ws[f"D{mid}"].font = font("000000", 10, bold=True)
    return ws


def build(path):
    wb = Workbook()
    build_quarterly(wb)
    build_dcf(wb)
    build_vs(wb)
    build_sens(wb)
    wb.active = 1
    wb.save(path)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "DCF_Template.xlsx")
    print("ok")
```

### `fill_dcf.py`

```python
#!/usr/bin/env python3
"""Fill the DCF template's blue input cells from a JSON file, keeping all template formatting.

Usage:  python3 fill_dcf.py TEMPLATE.xlsx inputs.json OUTPUT.xlsx

inputs.json:
{
  "Quarterly!B3": "Airbnb, Inc.",                      # plain value (number or text)
  "Quarterly!D7": {"date": "2025-12-31"},              # date
  "DCF!M7": {"text": "Q2'26 shareholder letter", "url": "https://..."},   # text with hyperlink
  "DCF!C13": {"text": "Closing price on 2026-10-01 (IBKR)"}               # note text (style kept)
}
Only writes the cells listed. Refuses to overwrite a template formula unless the key ends with '!'
(e.g. "DCF!H19!" to force) – formulas are the model, inputs are the only thing that changes.
"""
import datetime as dt
import json
import sys

from openpyxl import load_workbook
from openpyxl.styles import Font


def main(tpl, js, out):
    wb = load_workbook(tpl)
    data = json.load(open(js))
    for key, val in data.items():
        force = key.endswith("!")
        key = key.rstrip("!")
        sheet, ref = key.rsplit("!", 1)
        c = wb[sheet][ref]
        if isinstance(c.value, str) and c.value.startswith("=") and not force:
            raise SystemExit(f"{key} holds a formula ({c.value[:40]}…) – not overwriting")
        default_font = c.font is None or c.font.name in (None, "Calibri")
        if isinstance(val, dict):
            if default_font and "date" not in val:
                # unstyled cell (e.g. a Source / notes column): note style, or link style if it has a URL
                c.font = Font(name="Arial", size=9, italic=not val.get("url"),
                              color="FF0563C1" if val.get("url") else "FF595959",
                              underline="single" if val.get("url") else None)
            if "date" in val:
                c.value = dt.datetime.strptime(val["date"], "%Y-%m-%d")
            else:
                c.value = val.get("text")
            if val.get("url"):
                c.hyperlink = val["url"]
        else:
            c.value = val
    wb.save(out)
    print(f"filled {len(data)} cells -> {out}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
```

### `sec_pull.py`

```python
#!/usr/bin/env python3
"""Pull the latest 10-K and 10-Q financial statements for a ticker from SEC EDGAR.

Usage:  python3 sec_pull.py TICKER OUTDIR [--email you@example.com]

Writes to OUTDIR:
  filings.json            – CIK, company name, latest 10-K / 10-Q (accession, dates, document URL)
  <form>_<statement>.csv  – each primary statement as rendered by EDGAR (IS, BS, CF, and any
                            'Statements' report), company line labels kept verbatim
  statements.md           – the same tables as text, for reading and mapping to template rows
  xbrl_facts.json         – companyfacts JSON (for cross-checks and debt-note tags)

SEC asks for a descriptive User-Agent with a contact email (fair-access policy).
"""
import io, json, os, re, sys, time, urllib.request
import xml.etree.ElementTree as ET

import pandas as pd

UA = "dcf-builder research contact@example.com"


def get(url, binary=False):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Encoding": "identity"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                data = r.read()
            time.sleep(0.15)  # stay well under SEC's 10 req/s limit
            return data if binary else data.decode("utf-8", "replace")
        except Exception:
            if attempt == 2:
                raise
            time.sleep(1.5)


def cik_for(ticker):
    m = json.loads(get("https://www.sec.gov/files/company_tickers.json"))
    for row in m.values():
        if row["ticker"].upper() == ticker.upper():
            return int(row["cik_str"]), row["title"]
    raise SystemExit(f"Ticker {ticker} not found in SEC company_tickers.json")


def latest_filings(cik):
    sub = json.loads(get(f"https://data.sec.gov/submissions/CIK{cik:010d}.json"))
    rec = sub["filings"]["recent"]
    out = {}
    for i, form in enumerate(rec["form"]):
        if form in ("10-K", "10-Q", "20-F", "40-F") and form not in out:
            acc = rec["accessionNumber"][i]
            out[form] = {
                "form": form, "accession": acc, "filingDate": rec["filingDate"][i],
                "reportDate": rec["reportDate"][i], "primaryDocument": rec["primaryDocument"][i],
                "url": f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/{rec['primaryDocument'][i]}",
                "folder": f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/",
            }
    annual = out.get("10-K") or out.get("20-F") or out.get("40-F")
    q = out.get("10-Q")
    # a 10-Q only matters if it is newer than the annual report
    if annual and q and q["reportDate"] <= annual["reportDate"]:
        q = None
    return sub.get("name"), annual, q


STMT_PATTERNS = {
    "IS": r"(statements? of (consolidated )?(operations|income|earnings)|income statements?)(?!.*comprehensive)",
    "BS": r"balance sheets?|statements? of financial (position|condition)",
    "CF": r"statements? of cash flows?|cash flows? statements?",
}


def statements(filing):
    fs = get(filing["folder"] + "FilingSummary.xml")
    root = ET.fromstring(fs)
    picked = {}
    for rep in root.iter("Report"):
        name = (rep.findtext("ShortName") or "").strip()
        cat = (rep.findtext("MenuCategory") or "").strip()
        html = rep.findtext("HtmlFileName")
        if cat != "Statements" or not html:
            continue
        low = name.lower()
        if "parenthetical" in low:
            continue
        for key, pat in STMT_PATTERNS.items():
            if key not in picked and re.search(pat, low):
                picked[key] = (name, html)
    tables = {}
    for key, (name, html) in picked.items():
        raw = get(filing["folder"] + html)
        df = pd.read_html(io.StringIO(raw))[0]
        tables[key] = (name, df)
    return tables


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    ticker, outdir = sys.argv[1], sys.argv[2]
    global UA
    if "--email" in sys.argv:
        UA = f"dcf-builder research {sys.argv[sys.argv.index('--email') + 1]}"
    os.makedirs(outdir, exist_ok=True)
    cik, title = cik_for(ticker)
    name, annual, q = latest_filings(cik)
    info = {"ticker": ticker.upper(), "cik": cik, "name": name or title, "annual": annual, "quarterly": q}
    json.dump(info, open(os.path.join(outdir, "filings.json"), "w"), indent=2)
    md = [f"# {info['name']} ({ticker.upper()}) – CIK {cik}\n"]
    for f in (annual, q):
        if not f:
            continue
        md.append(f"\n## {f['form']} period {f['reportDate']} filed {f['filingDate']}\n{f['url']}\n")
        for key, (sname, df) in statements(f).items():
            df.to_csv(os.path.join(outdir, f"{f['form']}_{key}.csv"), index=False)
            md.append(f"\n### {key}: {sname}\n")
            md.append(df.fillna("").to_string(index=False, max_colwidth=70))
            md.append("\n")
    open(os.path.join(outdir, "statements.md"), "w").write("\n".join(md))
    facts = get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json")
    open(os.path.join(outdir, "xbrl_facts.json"), "w").write(facts)
    print(json.dumps(info, indent=2))


if __name__ == "__main__":
    main()
```
