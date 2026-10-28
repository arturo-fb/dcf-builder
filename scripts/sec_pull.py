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
