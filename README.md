# DCF Builder

**A universal discounted-cash-flow model for any US-listed company: one Excel template, three small
Python scripts, and an AI skill that fills it from SEC filings with a source on every input.**

[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
![Excel](https://img.shields.io/badge/Excel-template-217346.svg)
![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)

The template started as a full, hand-built DCF of Airbnb (ABNB) and was generalized so the same model
works for any company. **The template is the model:** every number is a live formula, and you only
type about 65 blue input cells from one 10-K, the latest 10-Q and the debt note. Everything else
(the forecast, WACC, terminal value, valuation summary and sensitivity tables) calculates itself.

---

## What's inside

| | |
|---|---|
| [`DCF_Template.xlsx`](DCF_Template.xlsx) | Blank template. Fill the blue cells; every sheet calculates. |
| [`examples/ABNB_DCF_example.xlsx`](examples/ABNB_DCF_example.xlsx) | The template filled for Airbnb as a proof test, with every input sourced. |
| [`examples/abnb_inputs.json`](examples/abnb_inputs.json) | The inputs behind that example, in the format `fill_dcf.py` reads. |
| [`scripts/sec_pull.py`](scripts/sec_pull.py) | Pulls the latest 10-K and 10-Q statements and XBRL facts from SEC EDGAR. No API key. |
| [`scripts/fill_dcf.py`](scripts/fill_dcf.py) | Writes an inputs JSON into the template. Refuses to overwrite a formula. |
| [`scripts/build_dcf_template.py`](scripts/build_dcf_template.py) | Rebuilds the blank template from code, so the model is reviewable and versionable. |
| [`skill/dcf-builder/SKILL.md`](skill/dcf-builder/SKILL.md) | AI skill: "DCF for COST" and it does the whole thing, sourced and verified. |

### The four sheets

- **Quarterly**: three fiscal years plus year-to-date from the filings, standardized; TTM is a formula. Built-in checks (operating income, net income and Adj. EBITDA must tie to the filing).
- **DCF**: revenue bridge from guidance, margin path, FCF build, CAPM cost of equity, bond-based cost of debt, WACC, Gordon terminal value, implied share price.
- **Valuation Summary**: DCF, peer forward P/E and EV/FCF, blended ⅓ each into a 12-month price target rounded to $5.
- **Sensitivity**: implied price across WACC × terminal growth and the other assumptions that move value most.

**Color code:** blue = input you type · black = formula · green = link from another sheet · yellow = key assumption.

---

## Methodology

These choices are fixed in the template, so every valuation built with it is comparable:

- **Adj. EBITDA** = operating income + D&A + SBC + other non-recurring add-backs.
- **Taxes are charged on EBIT** (EBITDA − D&A), and D&A is added back once.
- **Next fiscal year revenue** = year-to-date actual + guided quarter + rest of year.
- **Margin path** rises evenly from the company's guidance floor to its best full-year margin of the last three years by year 5: a level already proven, not a new high.
- **Capex %** = 3-year average; **NWC %** = 2-year average.
- **Cost of debt** = principal-weighted yield to maturity of the notes outstanding, priced from FINRA TRACE (falls back to the coupon when a note has no recent trade, and says so).
- **Cost of equity** = 10Y Treasury (FRED DGS10) + beta × equity risk premium (Kroll).
- **Terminal growth** = Fed long-run real GDP (latest SEP median) + 2% inflation.
- **Price target** = DCF, P/E and EV/FCF weighted ⅓ each, rounded to $5.

### Proof test: Airbnb

| | |
|---|---|
| Implied DCF share price | **$150.05** (original hand-built model: $152.90) |
| WACC | 10.9% |
| Terminal value as % of EV | 76% |
| Blended price target (DCF · P/E · EV/FCF) | **$180** |

The $2.85 gap to the original model is fully explained, item by item, by fixes the universal version made:

| Change | Effect |
|---|---:|
| FY+2 growth formula bug fixed | −$1.12 |
| D&A double-count removed | −$2.84 |
| Taxes charged on EBIT | +$0.57 |
| D&A % held at TTM | +$0.03 |
| Bond-based cost of debt | +$0.51 |

Rebuilding the template from code and filling it with `examples/abnb_inputs.json` reproduces the
example to the cent.

---

## Use it by hand

1. Open `DCF_Template.xlsx` and fill the blue cells on **Quarterly** from the latest 10-K (three fiscal
   years) and 10-Q (year-to-date columns, never the 3-month ones). Units are $ millions and millions of
   shares; expenses are positive numbers, capex is negative.
2. Fill the market inputs and assumptions on **DCF** (valuation date, share price, 10Y yield, beta, ERP,
   terminal growth, guidance) and the peers on **Valuation Summary**.
3. Put a source next to every input, in the source column beside it. The checks on Quarterly rows 20, 27
   and 36 must read 0.

`sec_pull.py` saves most of the typing:

```bash
pip install -r requirements.txt
python3 scripts/sec_pull.py COST data/ --email you@example.com   # SEC asks for a contact email
# → data/statements.md, IS/BS/CF CSVs, filings.json (with URLs), xbrl_facts.json for cross-checks
```

## Use it from code

```bash
python3 scripts/fill_dcf.py DCF_Template.xlsx my_inputs.json COST_DCF.xlsx
python3 scripts/build_dcf_template.py DCF_Template.xlsx          # rebuild the blank template
```

`inputs.json` maps cells to values. Plain numbers and text go in as-is; dates are `{"date": "YYYY-MM-DD"}`;
a sourced note is `{"text": "...", "url": "https://..."}`. See `examples/abnb_inputs.json`. `fill_dcf.py`
leaves every formula alone unless a key explicitly ends in `!`.

openpyxl writes formulas but doesn't calculate them. Open the file in Excel, or recalculate headless:
`soffice --headless --convert-to xlsx --outdir recalculated/ COST_DCF.xlsx`.

## Use it with an AI assistant

[`skill/dcf-builder/SKILL.md`](skill/dcf-builder/SKILL.md) is a complete, self-contained skill for Claude
(it embeds the three scripts, so it needs nothing else from this repo). Ask *"DCF for COST"* or
*"refresh my DCF"* and the assistant:

1. pulls the filings from EDGAR and fills **Quarterly**, cross-checking FY totals against XBRL;
2. builds the debt schedule from the debt note and FINRA TRACE prices;
3. sources the market inputs (FRED, beta, Kroll ERP, Fed SEP) and the guidance from the latest earnings release;
4. picks 2–5 peers (starting from the 10-K's competition section) and their forward P/Es;
5. fills, recalculates and runs a verification checklist: zero formula errors, statements tie, WACC
   between 6% and 14% and above terminal growth, terminal value share of EV flagged above 80%, every
   input sourced with a link;
6. delivers `TICKER_DCF.xlsx` with a 3–5 line summary of the price, target and the assumptions that matter most.

**Install:** in Claude Code, copy the folder to `~/.claude/skills/dcf-builder/`. In the Claude app, add
`skill/dcf-builder` as a custom skill (zip the folder and upload it in your skills settings). Other
assistants: paste SKILL.md in as the instructions. The assistant needs a shell with Python and web
access for EDGAR, FRED and peer multiples.

**Refresh mode** updates only the market inputs (date, price, 10Y yield, peer P/Es, consensus EPS) and
reports the old and new implied price and target, without touching anything else.

---

## Notes and limits

- **US filers (10-K / 10-Q).** Foreign issuers filing 20-F or 40-F (often IFRS, semiannual) work with care; the skill warns you first.
- **Risk-free rate:** `DCF!P5` pulls FRED live only in Windows Excel (`WEBSERVICE`). Mac Excel, Google Sheets and LibreOffice use the fallback value in `DCF!P9`; update it and its date in `Q9`.
- **Bond YTMs** use the FINRA TRACE price when one is entered; with no price, YTM falls back to the coupon. For convertible notes, use the effective interest rate instead.
- A DCF is only as good as its assumptions. The Sensitivity sheet is there to show how much they matter.

## Disclaimer

For education and research. Nothing here is investment advice or a recommendation to buy or sell any
security. Check every input against the filings yourself.

## License

[MIT](LICENSE) © 2026 Arturo Fraile
