#!/usr/bin/env python3
"""
fotorentas_risk.py — Client & Payment-Risk analyzer for Fotorentas
===================================================================

Reads a Costa Rican Hacienda electronic-invoicing export ("Reporte de ventas
general") and turns it into a client-level risk and revenue view, then writes a
self-contained HTML dashboard.

WHY THIS EXISTS
---------------
Fotorentas rents high-value camera/video/drone gear with NO security deposit and
an extraordinarily clean loss record (~0.6% of clients ever default, ~0.014% of
rentals never returned). So this is deliberately NOT a "fraud detector" — that
would be a control wildly oversized for the actual risk. Instead it does the
honest, useful thing for a clean book:

  1. Links every invoice to a client by cédula (the national-ID number).
  2. Separates repeat clients from one-timers, and individuals (física) from
     companies (jurídica) and foreigners (extranjero).
  3. Flags the only payment risk that actually exists in the data: invoices that
     are unpaid (Estado = Pendiente) or reversed for non-payment (Nota de
     crédito / Anulada). Any flagged client is put on a WATCH list so they're
     caught the next time they book.
  4. Anonymizes PII by default, so the output is safe to share / put on GitHub.

Right-sizing a control to the real risk — and building privacy in by default —
is itself the Trust & Safety skill this is meant to demonstrate.

USAGE
-----
    python3 fotorentas_risk.py INPUT.xlsx [-o dashboard.html] [--show-pii]

    --show-pii   Disable anonymization (real names + full cédulas). Use only on
                 your own machine, never for a public/portfolio copy.

Author: César B. Miranda — Fotorentas / Servicios XBM S.A.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------- #
# Loading & cleaning the Hacienda export
# --------------------------------------------------------------------------- #

# The export has a multi-row preamble (company info, record count, a category
# banner) before the real header row. We find the header dynamically rather than
# hard-coding a skiprows count, so the tool keeps working if the preamble grows.
HEADER_ANCHOR = "fecha de emision"  # accent-stripped marker of the header row


def _norm(text) -> str:
    """Lower-case and strip accents so column matching is robust to ó/é/°."""
    if text is None:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return text.strip().lower()


def _find_col(columns, *targets):
    """Return the real column label whose normalized name matches a target."""
    norm_map = {_norm(c): c for c in columns}
    for t in targets:
        if _norm(t) in norm_map:
            return norm_map[_norm(t)]
    raise KeyError(f"None of {targets} found in columns: {list(columns)}")


def _to_date(value):
    """Normalize an emission date to a Timestamp.

    openpyxl/pandas already parse date-formatted cells into datetimes, but some
    exports store the date as a raw Excel serial (days since 1899-12-30). Handle
    both so the tool is portable across export settings.
    """
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, pd.Timestamp):
        return value
    try:
        return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=int(float(value))))
    except (ValueError, TypeError):
        return pd.NaT


def _reserva_of(note: str):
    """First reservation number, anchored on 'Reserva:' so we don't accidentally
    grab years or amounts from the surrounding date text."""
    m = re.search(r"[Rr]eservas?:\s*([\d,\s y]+)", str(note))
    if not m:
        return pd.NA
    nums = re.findall(r"\d{3,5}", m.group(1))
    return int(nums[0]) if nums else pd.NA


def _parse_sheet(raw: pd.DataFrame):
    """Turn one raw sheet into tidy invoice rows, or None if it has no header."""
    header_row = next((i for i in range(len(raw))
                       if any(_norm(v) == HEADER_ANCHOR for v in raw.iloc[i].tolist())),
                      None)
    if header_row is None:
        return None
    df = raw.iloc[header_row + 1:].copy()
    df.columns = raw.iloc[header_row].tolist()

    cols = {
        "emitted": _find_col(df.columns, "Fecha de emisión"),
        "id_type": _find_col(df.columns, "Tipo de identificación"),
        "cedula": _find_col(df.columns, "N° de identificación"),
        "name": _find_col(df.columns, "Nombre"),
        "doc_type": _find_col(df.columns, "Tipo de documento"),
        "terms": _find_col(df.columns, "Condición de venta"),
        "state": _find_col(df.columns, "Estado"),
        "gross": _find_col(df.columns, "Total"),
        "notes": _find_col(df.columns, "Comentarios"),
    }
    df = df[list(cols.values())]
    df.columns = list(cols.keys())
    return df[df["cedula"].notna() & (df["cedula"].astype(str).str.strip() != "")]


def load_invoices(paths) -> pd.DataFrame:
    """Load one or many Hacienda exports into a tidy one-row-per-invoice frame.

    Handles the real-world quirks of these exports:
      * a file can hold several sheets (e.g. a mid-period v4.3 -> v4.4 format
        change), so every sheet with a valid header is read and concatenated;
      * 'N° de documento' is blank in these reports, so duplicates from
        overlapping sheets are removed on a composite key instead.
    """
    if isinstance(paths, (str, Path)):
        paths = [paths]

    frames = []
    for path in paths:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
        for raw in sheets.values():
            parsed = _parse_sheet(raw)
            if parsed is not None and len(parsed):
                frames.append(parsed)
    if not frames:
        sys.exit("ERROR: no sheet with a 'Fecha de emisión' header was found. "
                 "Are these Hacienda 'Reporte de ventas' exports?")

    df = pd.concat(frames, ignore_index=True)

    # Coerce types.
    df["gross"] = pd.to_numeric(df["gross"], errors="coerce").fillna(0.0)
    df["emitted"] = df["emitted"].apply(_to_date)
    for col in ("id_type", "name", "doc_type", "terms", "state", "notes"):
        df[col] = df[col].astype(str).str.strip()
    df["cedula"] = df["cedula"].astype(str).str.strip()
    df["reserva"] = df["notes"].apply(_reserva_of)

    # De-duplicate across overlapping sheets (the empty doc number forces a
    # content-based key).
    df = df.drop_duplicates(subset=["emitted", "cedula", "gross", "doc_type", "notes"])
    return df.reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Risk logic & client roll-up
# --------------------------------------------------------------------------- #

# An invoice "needs review" if it was annulled or credited. NOTE: validated
# against operator-confirmed records, this status is only a *weak* proxy for
# non-payment (~40% precision) — most annulments are billing corrections (wrong
# recipient, amount adjustments). Confirmed non-payment comes from the operations
# log; see the roadmap in README. We surface these for review, not as a verdict.
FLAG_STATES = {"pendiente", "anulada", "anulado"}
FLAG_DOCS = {"nota de credito", "nota de crédito"}


def is_flagged(state: str, doc_type: str) -> bool:
    return _norm(state) in FLAG_STATES or _norm(doc_type) in FLAG_DOCS


def short_id_type(id_type: str) -> str:
    n = _norm(id_type)
    if "juridica" in n:
        return "Company"
    if "extranjero" in n or "dimex" in n or "pasaporte" in n:
        return "Foreigner"
    return "Individual"


def build_clients(df: pd.DataFrame) -> pd.DataFrame:
    """Roll invoices up to one row per client (keyed on cédula)."""
    df = df.copy()
    df["flagged"] = df.apply(lambda r: is_flagged(r["state"], r["doc_type"]), axis=1)
    df["client_type"] = df["id_type"].apply(short_id_type)

    g = df.groupby("cedula")
    clients = pd.DataFrame({
        "name": g["name"].first(),
        "client_type": g["client_type"].first(),
        "invoices": g.size(),
        "gross": g["gross"].sum(),
        "flagged_invoices": g["flagged"].sum(),
    }).reset_index()

    def tier(row):
        if row["flagged_invoices"] > 0:
            return "REVIEW"           # has an annulled / credited invoice
        if row["invoices"] > 1:
            return "REPEAT"           # returning client
        return "SINGLE"               # one invoice so far
    clients["tier"] = clients.apply(tier, axis=1)
    return clients.sort_values("gross", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Anonymization (on by default)
# --------------------------------------------------------------------------- #

def anonymize(clients: pd.DataFrame) -> pd.DataFrame:
    """Replace names with stable Client IDs and mask cédulas."""
    clients = clients.copy()
    # Stable ID derived from a hash of the cédula -> deterministic across runs.
    def label(ced: str) -> str:
        h = int(hashlib.sha1(ced.encode()).hexdigest(), 16) % 1000
        return f"Client {h:03d}"
    clients["name"] = clients["cedula"].apply(label)
    clients["cedula"] = clients["cedula"].apply(lambda c: "•••" + str(c)[-3:])
    return clients


# --------------------------------------------------------------------------- #
# HTML dashboard  (camera-instrument readout aesthetic)
# --------------------------------------------------------------------------- #

def _money(v) -> str:
    return f"₡{v:,.0f}"


def build_dashboard(df: pd.DataFrame, clients: pd.DataFrame,
                    period: str, anonymized: bool) -> str:
    total_gross = df["gross"].sum()
    n_invoices = len(df)
    # One row per real client (grouping happened on the true cédula, before any
    # masking), so the row count is the correct unique-client total even when
    # anonymization collapses some masked IDs.
    n_clients = len(clients)
    flagged = clients[clients["tier"] == "REVIEW"]
    n_flagged = len(flagged)

    # Client-type mix.
    mix = clients["client_type"].value_counts().to_dict()
    mix_rows = "".join(
        f'<div class="mix-row"><span>{k}</span>'
        f'<span class="mono">{v}</span></div>'
        for k, v in mix.items()
    )

    # Invoice-state split.
    paid = int((~df.apply(lambda r: is_flagged(r["state"], r["doc_type"]), axis=1)).sum())
    pending = n_invoices - paid
    paid_pct = (paid / n_invoices * 100) if n_invoices else 0

    # Top clients by spend.
    top = clients.head(8)
    max_gross = top["gross"].max() if len(top) else 1
    top_rows = ""
    for _, r in top.iterrows():
        w = (r["gross"] / max_gross * 100) if max_gross else 0
        badge = {"REVIEW": "watch", "REPEAT": "repeat", "SINGLE": "single"}[r["tier"]]
        top_rows += (
            f'<div class="bar-row">'
            f'<div class="bar-label">{r["name"]}'
            f'<span class="tier {badge}">{r["tier"]}</span></div>'
            f'<div class="bar-track"><div class="bar-fill" style="width:{w:.1f}%"></div></div>'
            f'<div class="bar-val mono">{_money(r["gross"])}</div>'
            f'</div>'
        )

    # Annulled / credited invoices — surfaced for review, NOT a non-payment verdict.
    if n_flagged:
        watch_rows = "".join(
            f'<tr><td>{r["name"]}</td><td class="mono">{r["cedula"]}</td>'
            f'<td>{r["client_type"]}</td>'
            f'<td class="mono">{int(r["flagged_invoices"])}</td>'
            f'<td class="mono">{_money(r["gross"])}</td></tr>'
            for _, r in flagged.iterrows()
        )
        watch_block = f"""
        <table class="watch-table">
          <thead><tr><th>Client</th><th>Cédula</th><th>Type</th>
          <th>Annulled/credited</th><th>Lifetime</th></tr></thead>
          <tbody>{watch_rows}</tbody>
        </table>"""
    else:
        watch_block = ('<div class="empty">No annulled or credited invoices in '
                       'this period.</div>')

    pii_badge = ("PII ANONYMIZED" if anonymized else "PII VISIBLE — PRIVATE COPY")
    pii_class = "anon" if anonymized else "pii"

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fotorentas — Client &amp; Payment Risk</title>
<style>
  :root {{
    --bg:#0E0F11; --panel:#16181C; --panel2:#1B1E23; --line:#2A2E35;
    --ink:#E8E6E1; --muted:#868C96; --amber:#F2B544; --amber-dim:#7a5e23;
    --risk:#E5564B; --ok:#5FB286;
  }}
  * {{ box-sizing:border-box; }}
  body {{
    margin:0; background:var(--bg); color:var(--ink);
    font-family:ui-sans-serif,-apple-system,"Helvetica Neue",Arial,sans-serif;
    -webkit-font-smoothing:antialiased; line-height:1.4;
  }}
  .mono {{ font-family:ui-monospace,"SF Mono","JetBrains Mono",Menlo,monospace;
           font-variant-numeric:tabular-nums; }}
  .wrap {{ max-width:1040px; margin:0 auto; padding:30px 22px 60px;
           position:relative; }}
  /* viewfinder crop-marks: the signature element */
  .wrap::before,.wrap::after {{
    content:""; position:absolute; width:18px; height:18px; pointer-events:none;
  }}
  .wrap::before {{ top:14px; left:8px;
    border-top:2px solid var(--amber); border-left:2px solid var(--amber); }}
  .wrap::after {{ bottom:40px; right:8px;
    border-bottom:2px solid var(--amber); border-right:2px solid var(--amber); }}

  header {{ display:flex; justify-content:space-between; align-items:flex-end;
           flex-wrap:wrap; gap:14px; border-bottom:1px solid var(--line);
           padding-bottom:16px; margin-bottom:24px; }}
  .title {{ font-size:13px; letter-spacing:.32em; text-transform:uppercase;
            color:var(--muted); }}
  .title b {{ color:var(--ink); }}
  .title .amber {{ color:var(--amber); }}
  .meta {{ text-align:right; font-size:11px; letter-spacing:.12em; color:var(--muted); }}
  .badge {{ display:inline-block; margin-top:6px; font-size:10px; letter-spacing:.14em;
            padding:3px 8px; border-radius:2px; border:1px solid; }}
  .badge.anon {{ color:var(--ok); border-color:#2c4a3c; }}
  .badge.pii {{ color:var(--risk); border-color:#5a2b28; }}

  .kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:26px; }}
  .kpi {{ background:var(--panel); border:1px solid var(--line); border-radius:4px;
          padding:16px 16px 14px; }}
  .kpi .lab {{ font-size:10px; letter-spacing:.18em; text-transform:uppercase;
               color:var(--muted); margin-bottom:10px; }}
  .kpi .num {{ font-size:26px; font-weight:600; }}
  .kpi.flag .num {{ color:var(--risk); }}
  .kpi .num small {{ font-size:13px; color:var(--muted); font-weight:400; }}

  .grid {{ display:grid; grid-template-columns:1.55fr 1fr; gap:16px; margin-bottom:16px; }}
  .card {{ background:var(--panel); border:1px solid var(--line); border-radius:4px;
           padding:18px 18px 20px; }}
  .card h2 {{ font-size:11px; letter-spacing:.2em; text-transform:uppercase;
              color:var(--amber); margin:0 0 16px; font-weight:600; }}

  .bar-row {{ display:grid; grid-template-columns:1fr 90px; align-items:center;
              gap:8px 12px; margin-bottom:13px; }}
  .bar-label {{ font-size:13px; grid-column:1 / 2; }}
  .tier {{ font-size:9px; letter-spacing:.1em; padding:2px 6px; border-radius:2px;
           margin-left:8px; vertical-align:middle; }}
  .tier.repeat {{ color:var(--amber); background:#3a2e11; }}
  .tier.single {{ color:var(--muted); background:#23262c; }}
  .tier.watch {{ color:var(--risk); background:#3a201e; }}
  .bar-track {{ grid-column:1 / 2; height:6px; background:#23262c; border-radius:3px; }}
  .bar-fill {{ height:100%; background:linear-gradient(90deg,var(--amber-dim),var(--amber));
               border-radius:3px; }}
  .bar-val {{ grid-column:2 / 3; grid-row:1 / 3; text-align:right; font-size:13px;
              color:var(--ink); }}

  .mix-row {{ display:flex; justify-content:space-between; padding:9px 0;
              border-bottom:1px solid var(--line); font-size:13px; }}
  .mix-row:last-child {{ border-bottom:none; }}
  .mix-row .mono {{ color:var(--amber); }}

  .status {{ margin-top:6px; }}
  .status-bar {{ height:10px; border-radius:5px; overflow:hidden; display:flex;
                 background:var(--risk); margin-bottom:10px; }}
  .status-bar .paid {{ background:var(--ok); height:100%; }}
  .status-legend {{ display:flex; justify-content:space-between; font-size:12px;
                    color:var(--muted); }}
  .status-legend .mono {{ color:var(--ink); }}

  .watch h2 {{ color:var(--risk); }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th {{ text-align:left; font-size:10px; letter-spacing:.12em; text-transform:uppercase;
        color:var(--muted); border-bottom:1px solid var(--line); padding:0 0 10px; }}
  td {{ padding:11px 0; border-bottom:1px solid var(--line); }}
  tr:last-child td {{ border-bottom:none; }}
  .empty {{ color:var(--ok); font-size:13px; padding:6px 0; }}

  footer {{ margin-top:26px; padding-top:16px; border-top:1px solid var(--line);
            font-size:11px; color:var(--muted); line-height:1.7; }}
  footer .amber {{ color:var(--amber-dim); }}

  @media (max-width:720px) {{
    .kpis {{ grid-template-columns:repeat(2,1fr); }}
    .grid {{ grid-template-columns:1fr; }}
  }}
</style></head>
<body><div class="wrap">

  <header>
    <div class="title"><span class="amber">●</span> FOTORENTAS &nbsp;·&nbsp;
      <b>Client &amp; Payment Risk</b></div>
    <div class="meta">PERIOD &nbsp;{period}<br>
      SOURCE &nbsp;HACIENDA · REPORTE DE VENTAS GENERAL
      <br><span class="badge {pii_class}">{pii_badge}</span></div>
  </header>

  <div class="kpis">
    <div class="kpi"><div class="lab">Gross billed</div>
      <div class="num">{_money(total_gross)}</div></div>
    <div class="kpi"><div class="lab">Invoices</div>
      <div class="num">{n_invoices}</div></div>
    <div class="kpi"><div class="lab">Unique clients</div>
      <div class="num">{n_clients}</div></div>
    <div class="kpi flag"><div class="lab">To review</div>
      <div class="num">{n_flagged}<small> &nbsp;clients</small></div></div>
  </div>

  <div class="grid">
    <div class="card">
      <h2>Top clients by lifetime spend</h2>
      {top_rows}
    </div>
    <div>
      <div class="card" style="margin-bottom:16px;">
        <h2>Client mix</h2>
        {mix_rows}
      </div>
      <div class="card">
        <h2>Invoice state</h2>
        <div class="status">
          <div class="status-bar"><div class="paid" style="width:{paid_pct:.1f}%"></div></div>
          <div class="status-legend">
            <span><span class="mono">{paid}</span> active</span>
            <span><span class="mono">{pending}</span> annulled / credited</span>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="card watch">
    <h2>● Annulled / credited invoices — for review</h2>
    {watch_block}
  </div>

  <footer>
    Generated by <span class="amber">fotorentas_risk.py</span> from real Hacienda
    exports. Clients here have an annulled or credited invoice, surfaced
    <em>for review</em> — not flagged as non-payers. Validated against
    operator-confirmed records, annulment status is only a weak proxy for
    non-payment (~40% precision, ~43% recall): most annulments are routine
    billing corrections.<br>
    Confirmed non-payment lives in the operations log, not the invoicing system —
    integrating that ground-truth signal is the next step (see README roadmap).
  </footer>

</div></body></html>"""


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main():
    ap = argparse.ArgumentParser(description="Fotorentas client & payment-risk analyzer")
    ap.add_argument("input", nargs="+",
                    help="One or more Hacienda 'Reporte de ventas' .xlsx exports")
    ap.add_argument("-o", "--output", default="fotorentas_dashboard.html",
                    help="Output HTML path")
    ap.add_argument("--show-pii", action="store_true",
                    help="Disable anonymization (private use only)")
    args = ap.parse_args()

    df = load_invoices(args.input)
    clients = build_clients(df)
    anonymized = not args.show_pii
    if anonymized:
        clients = anonymize(clients)

    # Period label from the emission dates, falling back gracefully.
    dates = df["emitted"].dropna()
    if len(dates):
        period = f"{dates.min():%b %Y}" if dates.min().strftime("%b %Y") == \
                 dates.max().strftime("%b %Y") else f"{dates.min():%b %Y}–{dates.max():%b %Y}"
    else:
        period = "—"

    html = build_dashboard(df, clients, period, anonymized)
    Path(args.output).write_text(html, encoding="utf-8")

    # Console summary.
    print(f"Loaded {len(df)} invoices · {len(clients)} clients · "
          f"period {period}")
    print(f"Gross billed: {_money(df['gross'].sum())}")
    print(f"Annulled/credited (for review): {(clients['tier'] == 'REVIEW').sum()} clients")
    print(f"PII: {'anonymized' if anonymized else 'VISIBLE'}  ->  {args.output}")


if __name__ == "__main__":
    main()
