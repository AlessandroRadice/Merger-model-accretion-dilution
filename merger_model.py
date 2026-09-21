# %% [markdown]
# # Accretion / Dilution Merger Model
#
# **Does the deal create value for the acquirer's shareholders?**
#
# This notebook models a public-to-public acquisition the way an M&A analyst does before a board meeting:
#
# | Module | Output |
# |---|---|
# | 1. Market data | Price, shares, consensus EPS, net debt, EBITDA and book equity for acquirer and target |
# | 2. Offer & consideration | Offer price, premium, cash / stock mix, exchange ratio, implied multiples |
# | 3. Sources & uses | How the deal is funded: balance-sheet cash, new debt, new shares, fees |
# | 4. Purchase price allocation | Intangible write-up, deferred tax liability, goodwill |
# | 5. Pro forma EPS | Year 1–3 accretion / dilution, GAAP and cash EPS, EPS bridge |
# | 6. Breakeven synergies | Synergies needed for the deal to be EPS neutral each year |
# | 7. Sensitivities | Accretion vs premium × % stock, and synergies × % stock |
# | 8. Ownership & credit | Pro forma ownership vs contribution, pro forma leverage |
# | 9. Exports | **Excel model with live formulas** + charts |
#
# **How to use:** edit the *Configuration* cell, then `Runtime → Run all`.
#
# > Educational project. Consensus EPS and balance-sheet data from Yahoo Finance can be incomplete;
# > reconcile key inputs with the latest 10-K/10-Q and broker consensus before relying on the output.

# %%
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "yfinance", "xlsxwriter"], check=False)

# %%
import os, warnings, datetime as dt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

warnings.filterwarnings("ignore")
pd.set_option("display.float_format", lambda x: f"{x:,.2f}")

# %% [markdown]
# ## 0. Configuration

# %%
ACQUIRER = "MDLZ"          # buyer
TARGET   = "HSY"           # company being acquired
PROJECT_NAME = "Project Cocoa"
AUTHOR = "Alessandro Radice"
DATA_MODE = os.environ.get("MERGER_MODE", "live")   # "live" (Yahoo Finance) or "demo" (synthetic, offline)

DEAL = dict(
    premium        = 0.30,    # offer premium to the unaffected share price
    pct_stock      = 0.50,    # share of the purchase price paid in acquirer stock
    cost_of_debt   = 0.060,   # pre-tax rate on new acquisition debt
    interest_on_cash = 0.040, # pre-tax yield forgone on balance-sheet cash used
    min_cash       = None,    # cash the acquirer keeps on balance sheet; None -> 50% of current cash
    advisory_fees  = 0.012,   # % of target enterprise value, expensed at close (excluded from EPS)
    financing_fees = 0.020,   # % of new debt, capitalised and amortised
    financing_term = 7,       # years
    writeup_pct    = 0.25,    # share of excess purchase price allocated to amortisable intangibles
    intangible_life = 15,     # years
    synergies      = None,    # run-rate pre-tax cost synergies, $mm; None -> 4% of target revenue
    phase_in       = [0.40, 0.80, 1.00],   # share of run-rate synergies achieved in years 1-3
    tax_rate       = 0.23,
    acq_eps_growth = 0.07,    # EPS growth used for years 2-3
    tgt_eps_growth = 0.07,
)

OUTPUT_DIR = "merger_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)
TODAY = dt.date.today()

# %% [markdown]
# ## 1. Market data
# One standardized record per company (amounts in **$ millions**, except per-share data).

# %%
def fetch_live(tk):
    import yfinance as yf
    i = yf.Ticker(tk).info or {}
    price = i.get("currentPrice") or i.get("regularMarketPrice")
    shares = i.get("sharesOutstanding")
    eps = i.get("forwardEps")
    if not (price and shares and eps):
        raise ValueError(f"{tk}: missing price, shares or forward EPS")
    m = 1e6
    return dict(ticker=tk, name=i.get("longName", tk), price=float(price), shares=shares / m, eps=float(eps),
                debt=(i.get("totalDebt") or 0) / m, cash=(i.get("totalCash") or 0) / m,
                ebitda=(i.get("ebitda") or 0) / m, revenue=(i.get("totalRevenue") or 0) / m,
                book=(i.get("bookValue") or 0) * shares / m)

def demo_universe():
    acq = dict(ticker="ACQ.DEMO", name="Acquirer Co. (synthetic)", price=85.0, shares=1200.0, eps=4.50,
               debt=22000.0, cash=4000.0, ebitda=9800.0, revenue=52000.0, book=28000.0)
    tgt = dict(ticker="TGT.DEMO", name="Target Co. (synthetic)", price=42.0, shares=400.0, eps=2.40,
               debt=5000.0, cash=800.0, ebitda=2050.0, revenue=12000.0, book=6500.0)
    return acq, tgt

IS_DEMO = DATA_MODE != "live"
if not IS_DEMO:
    try:
        acq, tgt = fetch_live(ACQUIRER), fetch_live(TARGET)
    except Exception as e:
        print(f"Live data unavailable ({e}). Falling back to synthetic DEMO data.")
        IS_DEMO = True
if IS_DEMO:
    acq, tgt = demo_universe()
    if DEAL["synergies"] is None:
        DEAL["synergies"] = 500.0

for c in (acq, tgt):
    c["mcap"] = c["price"] * c["shares"]
    c["net_debt"] = c["debt"] - c["cash"]
    c["ev"] = c["mcap"] + c["net_debt"]
    c["ni"] = c["eps"] * c["shares"]
    c["pe"] = c["price"] / c["eps"]

if DEAL["synergies"] is None:
    DEAL["synergies"] = round(0.04 * tgt["revenue"], -1)
if DEAL["min_cash"] is None:
    DEAL["min_cash"] = 0.5 * acq["cash"]

DATA_LABEL = "SYNTHETIC DEMO DATA" if IS_DEMO else f"Yahoo Finance, consensus EPS; market data as of {TODAY:%d %b %Y}"
companies = pd.DataFrame([acq, tgt]).set_index("ticker")[["name", "price", "shares", "mcap", "net_debt", "ev", "ebitda", "revenue", "eps", "pe", "book"]]
companies

# %% [markdown]
# ## 2–4. Offer, sources & uses, purchase price allocation
# The cash portion is funded first from excess balance-sheet cash (above `min_cash`), then with new debt.
# Financing fees are grossed up into the debt raised. Advisory fees are expensed at close.

# %%
def merger(d=DEAL, a=acq, t=tgt):
    r = {}
    tax = d["tax_rate"]
    r["offer_price"] = t["price"] * (1 + d["premium"])
    r["equity_value"] = r["offer_price"] * t["shares"]
    r["enterprise_value"] = r["equity_value"] + t["net_debt"]
    r["cash_consid"] = r["equity_value"] * (1 - d["pct_stock"])
    r["stock_consid"] = r["equity_value"] * d["pct_stock"]
    r["advisory"] = d["advisory_fees"] * r["enterprise_value"]
    r["cash_used"] = min(max(a["cash"] - d["min_cash"], 0), r["cash_consid"] + r["advisory"])
    r["new_debt"] = (r["cash_consid"] + r["advisory"] - r["cash_used"]) / (1 - d["financing_fees"])
    r["fin_fees"] = r["new_debt"] * d["financing_fees"]
    r["new_shares"] = r["stock_consid"] / a["price"]
    r["exchange_ratio"] = r["offer_price"] / a["price"] * d["pct_stock"]
    # PPA
    r["excess"] = r["equity_value"] - t["book"]
    r["writeup"] = d["writeup_pct"] * r["excess"]
    r["dtl"] = r["writeup"] * tax
    r["goodwill"] = r["excess"] - r["writeup"] + r["dtl"]
    # Pro forma P&L, years 1-3
    rows = []
    for y in range(3):
        a_eps = a["eps"] * (1 + d["acq_eps_growth"]) ** y
        a_ni = a_eps * a["shares"]
        t_ni = t["eps"] * (1 + d["tgt_eps_growth"]) ** y * t["shares"]
        syn = d["synergies"] * d["phase_in"][y]
        int_new = r["new_debt"] * d["cost_of_debt"]
        int_lost = r["cash_used"] * d["interest_on_cash"]
        amort = r["writeup"] / d["intangible_life"]
        fin_am = r["fin_fees"] / d["financing_term"]
        pretax_adj = syn - int_new - int_lost - amort - fin_am
        ni = a_ni + t_ni + pretax_adj * (1 - tax)
        sh = a["shares"] + r["new_shares"]
        eps = ni / sh
        cash_eps = (ni + amort * (1 - tax)) / sh
        ni_ex_syn = ni - syn * (1 - tax)
        breakeven = max(0.0, (a_eps * sh - ni_ex_syn) / (1 - tax))
        rows.append({"Year": f"Year {y + 1}", "Acquirer NI": a_ni, "Target NI": t_ni, "Synergies": syn,
                     "Interest on new debt": -int_new, "Interest forgone on cash": -int_lost,
                     "Intangible amortisation": -amort, "Financing fee amortisation": -fin_am,
                     "Tax on adjustments": -pretax_adj * tax, "Pro forma NI": ni, "Pro forma shares": sh,
                     "Standalone EPS": a_eps, "Pro forma EPS": eps, "Accretion": eps / a_eps - 1,
                     "Cash EPS": cash_eps, "Cash accretion": cash_eps / a_eps - 1,
                     "Breakeven synergies": breakeven})
    r["pf"] = pd.DataFrame(rows).set_index("Year")
    # Ownership, contribution, credit
    sh = a["shares"] + r["new_shares"]
    r["own_acq"], r["own_tgt"] = a["shares"] / sh, r["new_shares"] / sh
    r["pf_net_debt"] = a["net_debt"] + t["net_debt"] + r["new_debt"] + r["cash_used"]
    r["pf_lev"] = r["pf_net_debt"] / (a["ebitda"] + t["ebitda"])
    r["pf_lev_syn"] = r["pf_net_debt"] / (a["ebitda"] + t["ebitda"] + d["synergies"])
    r["offer_pe"] = r["offer_price"] / t["eps"]
    r["offer_ev_ebitda"] = r["enterprise_value"] / t["ebitda"]
    return r

R = merger()
pf = R["pf"]
print(f"Offer ${R['offer_price']:.2f} ({DEAL['premium']:.0%} premium) | equity value ${R['equity_value']:,.0f}mm | "
      f"EV ${R['enterprise_value']:,.0f}mm | {R['offer_ev_ebitda']:.1f}x EBITDA | {R['offer_pe']:.1f}x P/E")
print(f"New debt ${R['new_debt']:,.0f}mm | cash used ${R['cash_used']:,.0f}mm | new shares {R['new_shares']:,.1f}mm")
print(f"Goodwill ${R['goodwill']:,.0f}mm | intangibles ${R['writeup']:,.0f}mm")
pf.T

# %% [markdown]
# ## 5–6. Accretion / dilution and breakeven synergies

# %%
summary = pd.DataFrame({"Standalone EPS": pf["Standalone EPS"], "Pro forma EPS": pf["Pro forma EPS"],
                        "Accretion": pf["Accretion"].map("{:+.1%}".format),
                        "Cash accretion": pf["Cash accretion"].map("{:+.1%}".format),
                        "Synergies": pf["Synergies"], "Breakeven synergies": pf["Breakeven synergies"]})
summary

# %% [markdown]
# ## 7. Sensitivities (Year 2 accretion)

# %%
prem_grid = [0.15, 0.225, 0.30, 0.375, 0.45]
stock_grid = [0.0, 0.25, 0.50, 0.75, 1.0]
syn_grid = [DEAL["synergies"] * k for k in (0.0, 0.5, 1.0, 1.5, 2.0)]

def acc(y=1, **kw):
    return merger({**DEAL, **kw})["pf"]["Accretion"].iloc[y]

sens_prem = pd.DataFrame([[acc(premium=p, pct_stock=s) for s in stock_grid] for p in prem_grid], index=prem_grid, columns=stock_grid)
sens_syn = pd.DataFrame([[acc(synergies=v, pct_stock=s) for s in stock_grid] for v in syn_grid], index=syn_grid, columns=stock_grid)
print("Year 2 accretion: premium (rows) x % stock (columns)")
print((sens_prem * 100).round(1))
print("\nYear 2 accretion: run-rate synergies (rows) x % stock (columns)")
print((sens_syn * 100).round(1))

# %% [markdown]
# ## 8. Cost of each currency, ownership and contribution
# A deal is accretive when the acquirer "pays" with a currency cheaper than the earnings yield it buys.

# %%
t_rate = DEAL["tax_rate"]
currency = pd.Series({"Acquirer stock (1 / P/E)": 1 / acq["pe"], "New debt (after tax)": DEAL["cost_of_debt"] * (1 - t_rate),
                      "Cash on hand (after tax)": DEAL["interest_on_cash"] * (1 - t_rate),
                      "Target earnings yield at offer": 1 / R["offer_pe"]})
contrib = pd.DataFrame({"Acquirer": [acq["revenue"], acq["ebitda"], acq["ni"], R["own_acq"]],
                        "Target": [tgt["revenue"], tgt["ebitda"], tgt["ni"], R["own_tgt"]]},
                       index=["Revenue", "EBITDA", "Net income (NTM)", "Pro forma ownership"])
contrib_pct = contrib.div(contrib.sum(axis=1), axis=0)
print(currency.map("{:.2%}".format).to_string())
print(f"\nPro forma net debt / EBITDA: {R['pf_lev']:.2f}x ({R['pf_lev_syn']:.2f}x with run-rate synergies)")
contrib_pct.map("{:.1%}".format)

# %% [markdown]
# ## 9. Charts

# %%
BLACK, GREY, LIGHT, INK, MUTED = "#1A1A18", "#8A8980", "#C8C7BF", "#1A1A18", "#66655E"
plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#b5b3ad", "xtick.color": MUTED, "ytick.color": MUTED})

fig, ax = plt.subplots(figsize=(9, 4.5))
x = np.arange(3)
ax.bar(x - 0.18, pf["Accretion"] * 100, 0.36, color=BLACK, label="GAAP EPS")
ax.bar(x + 0.18, pf["Cash accretion"] * 100, 0.36, color=LIGHT, label="Cash EPS (ex. amortisation)")
for i, (g, c) in enumerate(zip(pf["Accretion"], pf["Cash accretion"])):
    ax.text(i - 0.18, g * 100 + (0.15 if g >= 0 else -0.35), f"{g:+.1%}", ha="center", fontsize=9, color=INK)
    ax.text(i + 0.18, c * 100 + (0.15 if c >= 0 else -0.35), f"{c:+.1%}", ha="center", fontsize=9, color=INK)
ax.axhline(0, color=INK, lw=1); ax.margins(y=0.15)
ax.set_xticks(x, pf.index); ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
ax.set_title("EPS accretion / (dilution) to the acquirer", loc="left", fontsize=12, color=INK)
ax.legend(frameon=False, fontsize=9)
fig.text(0.01, 0.01, f"Source: {DATA_LABEL}", fontsize=7, color=MUTED)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(f"{OUTPUT_DIR}/accretion.png", dpi=200, metadata={"Title": "EPS accretion / (dilution) to the acquirer", "Author": AUTHOR, "Software": None}); plt.show()

# %% [markdown]
# ## 10. Excel export: live-formula model
# Blue = hard-coded input, black = formula, green = link to another sheet. The sensitivity grids are formulas too.

# %%
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

XLSX = f"{OUTPUT_DIR}/{acq['ticker'].split('.')[0]}_{tgt['ticker'].split('.')[0]}_merger_model.xlsx"
wb = xlsxwriter.Workbook(XLSX)
wb.set_properties({"title": f"{PROJECT_NAME}: {acq['name']} / {tgt['name']} merger model", "author": AUTHOR,
                   "subject": "Accretion / dilution analysis", "keywords": "M&A, merger model, accretion dilution",
                   "comments": "Generated by the Accretion / Dilution Merger Model", "created": dt.datetime.now()})
F = dict(font_name="Arial", font_size=10)
fmt = lambda **k: wb.add_format({**F, **k})
f_title = fmt(bold=True, font_size=14, font_color="#1A1A18"); f_sub = fmt(italic=True, font_color="#66655E")
f_hdr = fmt(bold=True, font_color="white", bg_color="#1A1A18", align="center", valign="vcenter", text_wrap=True)
f_b = fmt(bold=True)
IN = dict(font_color="#0000FF"); LK = dict(font_color="#008000")
NUM, PCT, PX, X_ = '#,##0.0;(#,##0.0)', '0.0%', '$#,##0.00', '0.0"x"'
f_in = {k: fmt(**IN, num_format=v) for k, v in dict(num=NUM, pct=PCT, px=PX, x=X_).items()}
f_lk = {k: fmt(**LK, num_format=v) for k, v in dict(num=NUM, pct=PCT, px=PX, x=X_).items()}
f_ = {k: fmt(num_format=v) for k, v in dict(num=NUM, pct=PCT, px=PX, x=X_).items()}
f_tot = fmt(num_format=NUM, bold=True, top=1)
f_key = fmt(num_format='+0.0%;-0.0%', bold=True, bg_color="#FFF2CC", border=1)
f_keypx = fmt(num_format='$#,##0.000', bold=True, border=1)
f_sens = fmt(num_format='+0.0%;-0.0%', border=1, align="center")
f_sh_p = fmt(num_format='0.0%', bold=True, bg_color="#ECEBE5", border=1, align="center")
f_sh_n = fmt(num_format='#,##0', bold=True, bg_color="#ECEBE5", border=1, align="center")

# ---------- Inputs ----------
wi = wb.add_worksheet("Inputs"); wi.hide_gridlines(2)
wi.set_column("A:A", 2); wi.set_column("B:B", 40); wi.set_column("C:E", 14)
wi.write("B2", f"{PROJECT_NAME}: inputs ($mm, except per share)", f_title)
wi.write("B3", f"Source: {DATA_LABEL}", f_sub)
wi.write_row(4, 1, ["Company data", "Acquirer", "Target"], f_hdr)
REF = {}
comp_rows = [("Ticker", "ticker", None), ("Share price", "price", "px"), ("Diluted shares (mm)", "shares", "num"),
             ("NTM EPS (consensus)", "eps", "px"), ("Total debt", "debt", "num"), ("Cash", "cash", "num"),
             ("LTM EBITDA", "ebitda", "num"), ("LTM revenue", "revenue", "num"), ("Book equity", "book", "num")]
for i, (lab, k, f) in enumerate(comp_rows):
    r = 5 + i
    wi.write(r, 1, lab)
    for j, c in enumerate((acq, tgt)):
        wi.write(r, 2 + j, c[k], f_in[f] if f else fmt(**IN))
        REF[("a", "t")[j] + "_" + k] = f"Inputs!${'CD'[j]}${r + 1}"
derived = [("Market capitalisation", "mcap", "={p}*{s}"), ("Net debt", "nd", "={d}-{c}"), ("P/E (NTM)", "pe", "={p}/{e}"),
           ("NTM net income", "ni", "={e}*{s}")]
for i, (lab, k, form) in enumerate(derived):
    r = 14 + i
    wi.write(r, 1, lab)
    for j, side in enumerate("at"):
        g = lambda key: REF[f"{side}_{key}"]
        wi.write_formula(r, 2 + j, form.format(p=g("price"), s=g("shares"), d=g("debt"), c=g("cash"), e=g("eps")),
                         f_["x"] if k == "pe" else f_["num"])
        REF[f"{side}_{k}"] = f"Inputs!${'CD'[j]}${r + 1}"
wi.write(19, 1, "Deal assumptions", f_b)
deal_rows = [("Offer premium", "premium", "pct"), ("% of price paid in stock", "pct_stock", "pct"),
             ("Pre-tax cost of new debt", "cost_of_debt", "pct"), ("Pre-tax interest on cash", "interest_on_cash", "pct"),
             ("Minimum cash retained by acquirer", "min_cash", "num"), ("Advisory fees (% of target EV)", "advisory_fees", "pct"),
             ("Financing fees (% of new debt)", "financing_fees", "pct"), ("Financing fee amortisation (years)", "financing_term", "num"),
             ("Intangible write-up (% of excess price)", "writeup_pct", "pct"), ("Intangible life (years)", "intangible_life", "num"),
             ("Run-rate pre-tax synergies", "synergies", "num"), ("Tax rate", "tax_rate", "pct"),
             ("Acquirer EPS growth (years 2-3)", "acq_eps_growth", "pct"), ("Target EPS growth (years 2-3)", "tgt_eps_growth", "pct")]
for i, (lab, k, f) in enumerate(deal_rows):
    r = 20 + i
    wi.write(r, 1, lab); wi.write(r, 2, DEAL[k], f_in[f]); REF[k] = f"Inputs!$C${r + 1}"
r = 20 + len(deal_rows)
wi.write(r, 1, "Synergy phase-in (years 1-3)")
for y in range(3):
    wi.write(r, 2 + y, DEAL["phase_in"][y], f_in["pct"]); REF[f"phase{y}"] = f"Inputs!${'CDE'[y]}${r + 1}"

# ---------- Transaction (offer, sources & uses, PPA) ----------
ws = wb.add_worksheet("Transaction"); ws.hide_gridlines(2)
ws.set_column("A:A", 2); ws.set_column("B:B", 40); ws.set_column("C:C", 14); ws.set_column("D:D", 3); ws.set_column("E:E", 34); ws.set_column("F:F", 14)
ws.write("B2", "Offer, sources & uses and purchase price allocation ($mm)", f_title)
T = lambda k: REF[k]
lines = [("Offer price per share", f"={T('t_price')}*(1+{T('premium')})", "px", "offer_price"),
         ("Equity purchase price", "=C5*{s}".format(s=T("t_shares")), "num", "equity"),
         ("(+) Target net debt assumed", f"={T('t_nd')}", "num", None),
         ("Transaction enterprise value", "=C6+C7", "num", "tev"),
         ("Offer EV / LTM EBITDA", f"=C8/{T('t_ebitda')}", "x", None),
         ("Offer P/E (NTM)", f"=C5/{T('t_eps')}", "x", "offer_pe"),
         ("Cash consideration", f"=C6*(1-{T('pct_stock')})", "num", "cash_c"),
         ("Stock consideration", f"=C6*{T('pct_stock')}", "num", "stock_c"),
         ("Advisory fees", f"={T('advisory_fees')}*C8", "num", "adv"),
         ("Balance-sheet cash used", f"=MIN(MAX({T('a_cash')}-{T('min_cash')},0),C11+C13)", "num", "cash_used"),
         ("New debt raised (incl. financing fees)", f"=(C11+C13-C14)/(1-{T('financing_fees')})", "num", "new_debt"),
         ("Financing fees", f"=C15*{T('financing_fees')}", "num", "fin_fees"),
         ("New acquirer shares issued (mm)", f"=C12/{T('a_price')}", "num", "new_sh"),
         ("Exchange ratio (stock component)", f"=C5/{T('a_price')}*{T('pct_stock')}", None, None)]
ws.write_row(3, 1, ["Offer and funding", ""], f_hdr)
for i, (lab, form, f, key) in enumerate(lines):
    r = 4 + i
    ws.write(r, 1, lab); ws.write_formula(r, 2, form, f_[f] if f else fmt(num_format="0.0000"))
    if key: REF[key] = f"Transaction!$C${r + 1}"
ws.write_row(3, 4, ["Sources", "$mm"], f_hdr)
for i, (lab, form) in enumerate([("New acquirer shares", "=C12"), ("New debt", "=C15"), ("Balance-sheet cash", "=C14")]):
    ws.write(4 + i, 4, lab); ws.write_formula(4 + i, 5, form, f_["num"])
ws.write(7, 4, "Total sources", f_b); ws.write_formula(7, 5, "=SUM(F5:F7)", f_tot)
ws.write_row(9, 4, ["Uses", "$mm"], f_hdr)
for i, (lab, form) in enumerate([("Purchase of target equity", "=C6"), ("Advisory fees", "=C13"), ("Financing fees", "=C16")]):
    ws.write(10 + i, 4, lab); ws.write_formula(10 + i, 5, form, f_["num"])
ws.write(13, 4, "Total uses", f_b); ws.write_formula(13, 5, "=SUM(F11:F13)", f_tot)
ws.write(14, 4, "Check (sources − uses)"); ws.write_formula(14, 5, "=F8-F14", f_["num"])
ws.write_row(19, 1, ["Purchase price allocation", ""], f_hdr)
ppa = [("Equity purchase price", "=C6", "num"), ("(−) Target book equity", f"=-{T('t_book')}", "num"),
       ("Excess purchase price", "=C21+C22", "num"), ("Intangible write-up", f"=C23*{T('writeup_pct')}", "num"),
       ("Deferred tax liability on write-up", f"=C24*{T('tax_rate')}", "num"), ("Goodwill", "=C23-C24+C25", "num"),
       ("Annual intangible amortisation", f"=C24/{T('intangible_life')}", "num")]
for i, (lab, form, f) in enumerate(ppa):
    ws.write(20 + i, 1, lab); ws.write_formula(20 + i, 2, form, f_tot if lab == "Goodwill" else f_[f])
REF["writeup"], REF["amort"] = "Transaction!$C$24", "Transaction!$C$27"

# ---------- Pro forma ----------
wp = wb.add_worksheet("Pro Forma"); wp.hide_gridlines(2)
wp.set_column("A:A", 2); wp.set_column("B:B", 38); wp.set_column("C:E", 14)
wp.write("B2", "Pro forma EPS: accretion / (dilution) ($mm, except per share)", f_title)
wp.write_row(3, 1, [""] + list(pf.index), f_hdr)
PL = ["Acquirer standalone EPS", "Acquirer net income", "Target net income", "Synergies (pre-tax)", "Interest on new debt",
      "Interest forgone on cash", "Intangible amortisation", "Financing fee amortisation", "Tax on adjustments",
      "Pro forma net income", "Acquirer shares", "New shares issued", "Pro forma shares", "Pro forma EPS",
      "Accretion / (dilution)", "Cash EPS (ex. amortisation)", "Cash EPS accretion / (dilution)", "Breakeven pre-tax synergies"]
PR = {lab: 4 + i for i, lab in enumerate(PL)}
for lab, r in PR.items():
    wp.write(r, 1, lab, f_b if lab in ("Pro forma net income", "Pro forma EPS", "Accretion / (dilution)") else None)
for y in range(3):
    c = 2 + y; C = xl_col_to_name(c)
    rr = lambda lab: f"{C}{PR[lab] + 1}"
    wp.write_formula(PR["Acquirer standalone EPS"], c, f"={T('a_eps')}*(1+{T('acq_eps_growth')})^{y}", f_keypx)
    wp.write_formula(PR["Acquirer net income"], c, f"={rr('Acquirer standalone EPS')}*{T('a_shares')}", f_["num"])
    wp.write_formula(PR["Target net income"], c, f"={T('t_eps')}*(1+{T('tgt_eps_growth')})^{y}*{T('t_shares')}", f_["num"])
    wp.write_formula(PR["Synergies (pre-tax)"], c, f"={T('synergies')}*{T(f'phase{y}')}", f_["num"])
    wp.write_formula(PR["Interest on new debt"], c, f"=-{T('new_debt')}*{T('cost_of_debt')}", f_["num"])
    wp.write_formula(PR["Interest forgone on cash"], c, f"=-{T('cash_used')}*{T('interest_on_cash')}", f_["num"])
    wp.write_formula(PR["Intangible amortisation"], c, f"=-{T('amort')}", f_lk["num"])
    wp.write_formula(PR["Financing fee amortisation"], c, f"=-{T('fin_fees')}/{T('financing_term')}", f_["num"])
    wp.write_formula(PR["Tax on adjustments"], c, f"=-SUM({C}{PR['Synergies (pre-tax)'] + 1}:{C}{PR['Financing fee amortisation'] + 1})*{T('tax_rate')}", f_["num"])
    wp.write_formula(PR["Pro forma net income"], c, f"=SUM({C}{PR['Acquirer net income'] + 1}:{C}{PR['Tax on adjustments'] + 1})", f_tot)
    wp.write_formula(PR["Acquirer shares"], c, f"={T('a_shares')}", f_lk["num"])
    wp.write_formula(PR["New shares issued"], c, f"={T('new_sh')}", f_lk["num"])
    wp.write_formula(PR["Pro forma shares"], c, f"={rr('Acquirer shares')}+{rr('New shares issued')}", f_tot)
    wp.write_formula(PR["Pro forma EPS"], c, f"={rr('Pro forma net income')}/{rr('Pro forma shares')}", f_keypx)
    wp.write_formula(PR["Accretion / (dilution)"], c, f"={rr('Pro forma EPS')}/{rr('Acquirer standalone EPS')}-1", f_key)
    wp.write_formula(PR["Cash EPS (ex. amortisation)"], c, f"=({rr('Pro forma net income')}-{rr('Intangible amortisation')}*(1-{T('tax_rate')}))/{rr('Pro forma shares')}", f_keypx)
    wp.write_formula(PR["Cash EPS accretion / (dilution)"], c, f"={rr('Cash EPS (ex. amortisation)')}/{rr('Acquirer standalone EPS')}-1", fmt(num_format='+0.0%;-0.0%'))
    wp.write_formula(PR["Breakeven pre-tax synergies"], c, f"=MAX(0,({rr('Acquirer standalone EPS')}*{rr('Pro forma shares')}-({rr('Pro forma net income')}-{rr('Synergies (pre-tax)')}*(1-{T('tax_rate')})))/(1-{T('tax_rate')}))", f_["num"])
REF["acc_y"] = [f"'Pro Forma'!${'CDE'[y]}${PR['Accretion / (dilution)'] + 1}" for y in range(3)]

# ---------- Sensitivity (live closed-form formulas) ----------
def acc_formula(prem, stk, syn, y=1):
    """Year-(y+1) accretion as one Excel formula, given cell refs (or literals) for premium, % stock and synergies."""
    P = f"({T('t_price')}*(1+{prem}))"; E = f"({P}*{T('t_shares')})"
    Cc = f"({E}*(1-{stk}))"; adv = f"({T('advisory_fees')}*({E}+{T('t_nd')}))"
    Cb = f"MIN(MAX({T('a_cash')}-{T('min_cash')},0),{Cc}+{adv})"
    Dn = f"(({Cc}+{adv}-{Cb})/(1-{T('financing_fees')}))"
    N = f"({E}*{stk}/{T('a_price')})"; W = f"({T('writeup_pct')}*({E}-{T('t_book')}))"
    adj = (f"({syn}*{T(f'phase{y}')}-{Dn}*{T('cost_of_debt')}-{Cb}*{T('interest_on_cash')}"
           f"-{W}/{T('intangible_life')}-{Dn}*{T('financing_fees')}/{T('financing_term')})")
    aeps = f"({T('a_eps')}*(1+{T('acq_eps_growth')})^{y})"
    ni = f"({aeps}*{T('a_shares')}+{T('t_eps')}*(1+{T('tgt_eps_growth')})^{y}*{T('t_shares')}+{adj}*(1-{T('tax_rate')}))"
    return f"={ni}/({T('a_shares')}+{N})/{aeps}-1"

wz = wb.add_worksheet("Sensitivity"); wz.hide_gridlines(2)
wz.set_column("A:A", 2); wz.set_column("B:B", 16); wz.set_column("C:G", 12)
wz.write("B2", "Year 2 EPS accretion / (dilution): every cell is a full merger-model formula", f_title)
def grid(top, title, row_vals, row_fmt, cell):
    wz.write(top, 1, title, f_b)
    wz.write(top + 1, 1, "↓ rows / % stock →", f_sub)
    for j, s in enumerate(stock_grid):
        wz.write(top + 1, 2 + j, s, f_sh_p)
    for i, v in enumerate(row_vals):
        r = top + 2 + i
        wz.write(r, 1, v, row_fmt)
        for j in range(len(stock_grid)):
            wz.write_formula(r, 2 + j, cell(f"$B${r + 1}", f"{xl_col_to_name(2 + j)}${top + 2}"), f_sens)
grid(3, "Offer premium (rows) × % stock (columns)", prem_grid, f_sh_p, lambda rv, cv: acc_formula(rv, cv, T("synergies")))
grid(11, "Run-rate synergies, $mm (rows) × % stock (columns)", syn_grid, f_sh_n, lambda rv, cv: acc_formula(T("premium"), cv, rv))

# ---------- Ownership & credit ----------
wo = wb.add_worksheet("Ownership & Credit"); wo.hide_gridlines(2)
wo.set_column("A:A", 2); wo.set_column("B:B", 36); wo.set_column("C:E", 14)
wo.write("B2", "Contribution vs ownership, and pro forma credit ($mm)", f_title)
wo.write_row(3, 1, ["Contribution", "Acquirer", "Target", "Target %"], f_hdr)
for i, (lab, a_, t_) in enumerate([("LTM revenue", T("a_revenue"), T("t_revenue")), ("LTM EBITDA", T("a_ebitda"), T("t_ebitda")),
                                   ("NTM net income", T("a_ni"), T("t_ni")), ("Pro forma shares (mm)", T("a_shares"), T("new_sh"))]):
    r = 4 + i
    wo.write(r, 1, lab); wo.write_formula(r, 2, f"={a_}", f_lk["num"]); wo.write_formula(r, 3, f"={t_}", f_lk["num"])
    wo.write_formula(r, 4, f"=D{r + 1}/(C{r + 1}+D{r + 1})", f_["pct"])
cred = [("Acquirer net debt", f"={T('a_nd')}"), ("Target net debt", f"={T('t_nd')}"), ("New debt", f"={T('new_debt')}"),
        ("Balance-sheet cash used", f"={T('cash_used')}"), ("Pro forma net debt", "=SUM(C11:C14)"),
        ("Combined LTM EBITDA", f"={T('a_ebitda')}+{T('t_ebitda')}"), ("Pro forma net debt / EBITDA", "=C15/C16"),
        ("… including run-rate synergies", f"=C15/(C16+{T('synergies')})")]
wo.write_row(9, 1, ["Pro forma credit", ""], f_hdr)
for i, (lab, form) in enumerate(cred):
    wo.write(10 + i, 1, lab); wo.write_formula(10 + i, 2, form, f_["x"] if "/" in lab or "…" in lab else f_["num"])
wb.close()
print("Saved", XLSX)

# %% [markdown]
# ## 11. Download

# %%
try:
    from google.colab import files
    for fpath in [XLSX, f"{OUTPUT_DIR}/accretion.png"]:
        files.download(fpath)
except ImportError:
    print("Outputs saved in:", os.path.abspath(OUTPUT_DIR))
