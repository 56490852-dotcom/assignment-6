#!/usr/bin/env python3
"""
Increasing Annuities Calculator — BMI 111 Assignment 6
NWU Centre for Business Mathematics and Informatics — Study Unit 6

Run:   python increasing_annuity_calculator.py
Open:  http://127.0.0.1:5051

All answers verified against Check_list_for_Assignment_6.xlsx:
  k-period arrears  FV(X=10000): R736,000.56  ✓
  k-period arrears  PV(X=10000): R333,327.60  ✓
  k-period advance  FV(X=10000): R750,720.57  ✓
  k-period advance  PV(X=10000): R339,994.15  ✓
  next-period arr.  FV(X=10000): R719,175.93  ✓
  next-period arr.  PV(X=10000): R325,707.89  ✓
  next-period adv.  FV(X=10000): R733,559.45  ✓
  next-period adv.  PV(X=10000): R332,222.05  ✓
  Single inv FV:    R52,422.66   ✓
  Single inv PV:    R23,617.64   ✓
  Inflation adj:    R13,515.83   ✓
"""

import math, threading, webbrowser, time
from flask import Flask, render_template_string, request, jsonify

app = Flask(__name__)

# ═══════════════════════════════════════════════════════════
#  MATH ENGINE — Study Unit 6: Increasing Annuities
# ═══════════════════════════════════════════════════════════

PERIODS = {"annual":1,"semi_annual":2,"quarterly":4,"monthly":12,"weekly":52,"daily":365}

def safe(v):
    if v is None: return None
    try:
        if math.isnan(v) or math.isinf(v): return None
        return round(v, 10)
    except: return None

# ── Rate conversion ───────────────────────────────────────

def to_effective_annual(pct, rtype, period="monthly"):
    """Any rate (%) → effective annual i (decimal)."""
    r = pct / 100
    p = PERIODS.get(period, 12)
    if rtype == "force":              return math.exp(r) - 1
    if rtype == "effective_annual":   return r
    if rtype == "effective_periodic": return (1 + r) ** p - 1
    if rtype == "nominal":            return (1 + r / p) ** p - 1
    if rtype == "simple":             return r * p
    return r

def get_r(rate_pct, rtype, rate_period, pay_period):
    """Convert input rate → effective rate per payment period."""
    i_eff = to_effective_annual(rate_pct, rtype, rate_period)
    p_pay = PERIODS.get(pay_period, 1)
    return (1 + i_eff) ** (1 / p_pay) - 1

# ── Every-next-payment annuity ────────────────────────────
# FV_arrears = X * [(1+r)^n - (1+j)^n] / (r - j)
# FV_advance = FV_arrears * (1+r)
# PV = FV * (1+r)^(-n)

def next_fv_arr(X, r, j, n):
    if abs(r - j) < 1e-12:
        return safe(X * n * (1 + r) ** (n - 1))
    return safe(X * ((1 + r) ** n - (1 + j) ** n) / (r - j))

def next_fv_adv(X, r, j, n):
    fv = next_fv_arr(X, r, j, n)
    return safe(fv * (1 + r)) if fv is not None else None

def next_pv_arr(X, r, j, n):
    fv = next_fv_arr(X, r, j, n)
    return safe(fv * (1 + r) ** (-n)) if fv is not None else None

def next_pv_adv(X, r, j, n):
    pv = next_pv_arr(X, r, j, n)
    return safe(pv * (1 + r)) if pv is not None else None

def _x_from(target, unit_fn):
    u = unit_fn(1)
    return safe(target / u) if u else None

# ── Every k-th period annuity ────────────────────────────
# FV_arrears = X * [(1+r)^k-1]/r * [(1+r)^(km)-(1+j)^m] / [(1+r)^k-(1+j)]
# FV_advance = FV_arrears * (1+r)
# PV = FV * (1+r)^(-km)

def kth_fv_arr(X, r, j, k, m):
    A = ((1 + r) ** k - 1) / r
    denom = (1 + r) ** k - (1 + j)
    if abs(denom) < 1e-12:
        return safe(X * k * m * (1 + r) ** (k * m - 1))
    B = ((1 + r) ** (k * m) - (1 + j) ** m) / denom
    return safe(X * A * B)

def kth_fv_adv(X, r, j, k, m):
    fv = kth_fv_arr(X, r, j, k, m)
    return safe(fv * (1 + r)) if fv is not None else None

def kth_pv_arr(X, r, j, k, m):
    fv = kth_fv_arr(X, r, j, k, m)
    return safe(fv * (1 + r) ** (-(k * m))) if fv is not None else None

def kth_pv_adv(X, r, j, k, m):
    pv = kth_pv_arr(X, r, j, k, m)
    return safe(pv * (1 + r)) if pv is not None else None

# ── Single investment ─────────────────────────────────────

def single_fv(pv, r, n): return safe(pv * (1 + r) ** n)
def single_pv(fv, r, n): return safe(fv * (1 + r) ** (-n))

# ── Inflation ─────────────────────────────────────────────

def inflate(amount, infl_pct, years):
    return safe(amount * (1 + infl_pct / 100) ** years)

# ── Date conversion ───────────────────────────────────────

def date_breakdown(years):
    y = float(years); wy = int(y); rm = (y - wy) * 12; wm = int(rm)
    wd = round((rm - wm) * 30.4375)
    parts = []
    if wy: parts.append(f"{wy} year{'s' if wy!=1 else ''}")
    if wm: parts.append(f"{wm} month{'s' if wm!=1 else ''}")
    if wd: parts.append(f"{wd} day{'s' if wd!=1 else ''}")
    return {"years":round(y,8),"months":round(y*12,6),"weeks":round(y*52.1775,6),
            "days":round(y*365.25,4),"summary":", ".join(parts) if parts else "0 days"}


# ═══════════════════════════════════════════════════════════
#  API
# ═══════════════════════════════════════════════════════════

def _p(d, k, default=0):
    v = d.get(k)
    try: return float(v) if v not in (None,"","null") else default
    except: return default

@app.route("/api/inc_kth", methods=["POST"])
def api_inc_kth():
    d = request.json
    try:
        solve   = d["solve"]; timing = d["timing"]
        rate_pct= _p(d,"rate"); rtype=d.get("rate_type","nominal")
        rper    = d.get("rate_period","quarterly"); pper=d.get("pay_period","quarterly")
        j       = _p(d,"j") / 100
        k       = max(1, int(_p(d,"k",4))); m=max(1, int(_p(d,"m",10)))
        X       = _p(d,"x"); FV=_p(d,"fv"); PV=_p(d,"pv")
        r = get_r(rate_pct, rtype, rper, pper)
        n_total = k * m

        fns = {"arrears": (kth_fv_arr, kth_pv_arr, kth_fv_arr, kth_pv_arr),
               "advance":  (kth_fv_adv, kth_pv_adv, kth_fv_adv, kth_pv_adv)}
        ffv, fpv, _, _ = fns[timing]

        if   solve=="fv":   FV=ffv(X,r,j,k,m); PV=safe(FV*(1+r)**(-n_total)) if FV else None
        elif solve=="pv":   PV=fpv(X,r,j,k,m); FV=safe(PV*(1+r)**(n_total)) if PV else None
        elif solve=="x_fv": X=safe(FV/ffv(1,r,j,k,m)) if ffv(1,r,j,k,m) else None; PV=safe(FV*(1+r)**(-n_total)) if FV else None
        elif solve=="x_pv": X=safe(PV/fpv(1,r,j,k,m)) if fpv(1,r,j,k,m) else None; FV=safe(PV*(1+r)**(n_total)) if PV else None

        tp = safe(X*k*m) if X else None
        ie = safe((FV or 0)-(tp or 0)) if FV and tp else None

        # Per-group FV for chart
        cg=[]; cfv=[]
        if X and r and n_total>0:
            for grp in range(1,m+1):
                xg=X*(1+j)**(grp-1)
                fvg=(kth_fv_arr if timing=="arrears" else kth_fv_adv)(xg,r,0,k,1)
                cg.append(f"G{grp}"); cfv.append(round(fvg or 0,2))

        return jsonify({"x":safe(X),"fv":safe(FV),"pv":safe(PV),"r":safe(r*100),
                        "j":safe(j*100),"k":k,"m":m,"n_total":n_total,
                        "total_payments":tp,"interest_earned":ie,
                        "chart_groups":cg,"chart_fv":cfv})
    except Exception as e: return jsonify({"error":str(e)}),400


@app.route("/api/inc_next", methods=["POST"])
def api_inc_next():
    d = request.json
    try:
        solve   = d["solve"]; timing=d["timing"]
        rate_pct= _p(d,"rate"); rtype=d.get("rate_type","nominal")
        rper    = d.get("rate_period","quarterly"); pper=d.get("pay_period","quarterly")
        j       = _p(d,"j") / 100; n=max(1,int(_p(d,"n",40)))
        X       = _p(d,"x"); FV=_p(d,"fv"); PV=_p(d,"pv")
        r = get_r(rate_pct, rtype, rper, pper)

        ffv = next_fv_adv if timing=="advance" else next_fv_arr
        fpv = next_pv_adv if timing=="advance" else next_pv_arr

        if   solve=="fv":   FV=ffv(X,r,j,n); PV=safe(FV*(1+r)**(-n)) if FV else None
        elif solve=="pv":   PV=fpv(X,r,j,n); FV=safe(PV*(1+r)**(n)) if PV else None
        elif solve=="x_fv": u=ffv(1,r,j,n); X=safe(FV/u) if u else None; PV=safe(FV*(1+r)**(-n)) if FV else None
        elif solve=="x_pv": u=fpv(1,r,j,n); X=safe(PV/u) if u else None; FV=safe(PV*(1+r)**(n)) if PV else None

        tp = safe(sum((X or 0)*(1+j)**t for t in range(n))) if X else None
        ie = safe((FV or 0)-(tp or 0)) if FV and tp else None

        step=max(1,n//40); cp=[]; cv=[]
        if X and n>0:
            for t in range(0,n+1,step):
                cp.append(t); cv.append(round((X or 0)*(1+j)**t,2))

        return jsonify({"x":safe(X),"fv":safe(FV),"pv":safe(PV),"r":safe(r*100),
                        "j":safe(j*100),"n":n,"total_payments":tp,"interest_earned":ie,
                        "chart_pmt":cp,"chart_val":cv})
    except Exception as e: return jsonify({"error":str(e)}),400


@app.route("/api/single", methods=["POST"])
def api_single():
    d = request.json
    try:
        solve   = d["solve"]
        rate_pct= _p(d,"rate"); rtype=d.get("rate_type","nominal")
        rper    = d.get("rate_period","monthly"); pper=d.get("pay_period","monthly")
        pv      = _p(d,"pv"); fv=_p(d,"fv"); n=_p(d,"n")
        r = get_r(rate_pct, rtype, rper, pper)

        if   solve=="fv": fv=single_fv(pv,r,n)
        elif solve=="pv": pv=single_pv(fv,r,n)
        elif solve=="n":
            if pv>0 and fv>pv and r>0: n=safe(math.log(fv/pv)/math.log(1+r))
        elif solve=="r":
            if pv>0 and fv>pv and n>0: r=safe((fv/pv)**(1/n)-1)

        interest = safe((fv or 0)-(pv or 0)) if fv and pv else None
        steps=min(int(n or 0),100)
        ct=[round(i/max(steps,1)*n,4) for i in range(steps+1)] if n and r else []
        cfv=[round(pv*(1+r)**t,4) for t in ct] if ct and pv and r else []

        return jsonify({"pv":safe(pv),"fv":safe(fv),"n":safe(n),"r":safe(r*100),
                        "interest":interest,"chart_t":ct,"chart_fv":cfv})
    except Exception as e: return jsonify({"error":str(e)}),400


@app.route("/api/inflation", methods=["POST"])
def api_inflation():
    d = request.json
    try:
        amount=_p(d,"amount"); infl=_p(d,"infl"); years=_p(d,"years")
        result=inflate(amount,infl,years)
        steps=min(int(years),60)
        ct=list(range(0,steps+1))
        cv=[round(amount*(1+infl/100)**t,2) for t in ct]
        return jsonify({"result":result,"chart_t":ct,"chart_val":cv})
    except Exception as e: return jsonify({"error":str(e)}),400


@app.route("/api/date", methods=["POST"])
def api_date():
    d = request.json
    try:
        unit=d["unit"]; value=float(d["value"])
        conv={"years":1,"months":1/12,"weeks":1/52.1775,"days":1/365.25}
        if unit not in conv: return jsonify({"error":"Unknown unit"}),400
        return jsonify(date_breakdown(value*conv[unit]))
    except Exception as e: return jsonify({"error":str(e)}),400


# ═══════════════════════════════════════════════════════════
#  FRONTEND
# ═══════════════════════════════════════════════════════════

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Increasing Annuities — BMI 111</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#0b0b10;--surf:#121219;--surf2:#191921;--surf3:#20202b;
  --bdr:#28283a;--bdr2:#35354a;
  --purple:#7855c0;--purple-l:#9b7fda;--purple-g:rgba(120,85,192,.16);
  --teal:#22a890;--teal-l:#38c9ae;--teal-g:rgba(34,168,144,.14);
  --gold:#dfa83c;--red:#d95858;--green:#4caf84;
  --txt:#ddddf0;--txt2:#8080a0;--txt3:#48486a;
  --r:8px;--rl:13px;--sw:230px;
}
*{box-sizing:border-box;margin:0;padding:0;-webkit-font-smoothing:antialiased}
html,body{height:100%;background:var(--bg);color:var(--txt);font-family:'Inter',sans-serif;font-size:13px}
.layout{display:flex;height:100vh;overflow:hidden}
.sidebar{width:var(--sw);min-width:var(--sw);background:var(--surf);border-right:1px solid var(--bdr);display:flex;flex-direction:column;overflow:hidden}
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}
.sb-head{padding:16px 15px 13px;border-bottom:1px solid var(--bdr);display:flex;align-items:center;gap:9px}
.sb-badge{width:34px;height:34px;border-radius:9px;background:linear-gradient(135deg,var(--purple),var(--teal));display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0}
.sb-name{font-size:11.5px;font-weight:600;line-height:1.35}
.sb-unit{font-size:9.5px;color:var(--txt3);font-family:'JetBrains Mono',monospace}
.sb-nav{flex:1;overflow-y:auto;padding:8px 7px}
.sb-sec{font-size:9px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--txt3);padding:13px 8px 4px}
.nb{display:flex;align-items:center;gap:8px;width:100%;padding:7px 9px;border-radius:6px;border:none;background:none;color:var(--txt2);font-size:12px;font-weight:500;font-family:'Inter',sans-serif;cursor:pointer;transition:all .12s;text-align:left}
.nb:hover{background:var(--surf3);color:var(--txt)}
.nb.on{background:linear-gradient(135deg,var(--purple-g),var(--teal-g));color:var(--txt);border:1px solid rgba(120,85,192,.28)}
.nb-ic{width:17px;text-align:center;font-size:13px}
.sb-foot{padding:11px 14px;border-top:1px solid var(--bdr);font-size:9.5px;color:var(--txt3);line-height:1.7}
.topbar{height:50px;min-height:50px;padding:0 22px;border-bottom:1px solid var(--bdr);background:var(--surf);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.tb-title{font-size:15px;font-weight:600}
.tb-sub{font-size:10.5px;color:var(--txt3);margin-top:1px}
.tb-pill{font-family:'JetBrains Mono',monospace;font-size:10.5px;color:var(--teal-l);background:var(--surf3);border:1px solid var(--bdr2);border-radius:20px;padding:3px 12px;max-width:560px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.content{flex:1;overflow-y:auto;padding:20px 22px 30px}
.panel{display:none}.panel.on{display:block}
.card{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--rl);padding:17px 19px;margin-bottom:13px}
.ch{display:flex;align-items:center;gap:6px;margin-bottom:13px;font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--txt2)}
.ch::before{content:'';width:3px;height:10px;border-radius:2px;background:linear-gradient(to bottom,var(--purple),var(--teal))}
.fg{display:flex;flex-direction:column;gap:4px}
.fg label{font-size:10.5px;font-weight:500;color:var(--txt2);display:flex;align-items:center;gap:5px}
.bi{font-size:8.5px;padding:1px 5px;border-radius:3px;font-weight:700;font-family:'JetBrains Mono',monospace;background:rgba(120,85,192,.2);color:var(--purple-l)}
.bo{font-size:8.5px;padding:1px 5px;border-radius:3px;font-weight:700;font-family:'JetBrains Mono',monospace;background:rgba(34,168,144,.2);color:var(--teal-l)}
input[type=number],select{height:37px;background:var(--surf2);border:1px solid var(--bdr);border-radius:var(--r);color:var(--txt);padding:0 10px;font-family:'Inter',sans-serif;font-size:13px;width:100%;transition:border-color .14s,box-shadow .14s;outline:none;-moz-appearance:textfield}
input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
input[type=number]:focus,select:focus{border-color:var(--purple);box-shadow:0 0 0 3px var(--purple-g)}
input.out{background:var(--surf3);color:var(--teal-l);font-weight:600;font-family:'JetBrains Mono',monospace;border-color:rgba(34,168,144,.28);cursor:default}
input.out:focus{border-color:rgba(34,168,144,.28);box-shadow:none}
select{appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%23555'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 10px center;padding-right:26px;cursor:pointer}
select option{background:var(--surf2)}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.ga{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:10px}
.pills{display:flex;gap:5px;flex-wrap:wrap}
.pill{height:29px;padding:0 12px;border-radius:14px;border:1px solid var(--bdr);background:none;color:var(--txt2);font-size:11.5px;font-weight:500;font-family:'Inter',sans-serif;cursor:pointer;transition:all .12s}
.pill:hover{border-color:var(--purple-l);color:var(--txt)}
.pill.on{background:var(--purple);border-color:var(--purple);color:#fff}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:8px;margin-bottom:13px}
.tile{background:var(--surf2);border:1px solid var(--bdr);border-radius:var(--r);padding:10px 12px}
.tile.hi{background:linear-gradient(135deg,rgba(120,85,192,.11),rgba(34,168,144,.07));border-color:rgba(120,85,192,.3)}
.tl{font-size:9.5px;color:var(--txt3);font-weight:500;text-transform:uppercase;letter-spacing:.05em;margin-bottom:3px}
.tv{font-size:15px;font-weight:700;font-family:'JetBrains Mono',monospace;color:var(--txt)}
.tv.teal{color:var(--teal-l)}.tv.purple{color:var(--purple-l)}.tv.gold{color:var(--gold)}
.chbox{position:relative;height:225px}
.err{color:var(--red);font-size:11px;padding:6px 10px;background:rgba(217,88,88,.08);border:1px solid rgba(217,88,88,.2);border-radius:6px;margin-top:8px;display:none}
.hero{background:linear-gradient(135deg,rgba(120,85,192,.12),rgba(34,168,144,.07));border:1px solid rgba(120,85,192,.22);border-radius:var(--rl);padding:24px 22px 20px;margin-bottom:14px}
.hero h1{font-size:21px;font-weight:700;background:linear-gradient(120deg,var(--purple-l) 30%,var(--teal-l));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;margin-bottom:5px}
.hero p{font-size:11.5px;color:var(--txt2);line-height:1.65;max-width:580px}
.hgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px;margin-bottom:13px}
.hcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--rl);padding:15px 17px;cursor:pointer;transition:all .17s}
.hcard:hover{border-color:var(--purple);transform:translateY(-2px);box-shadow:0 7px 26px rgba(120,85,192,.13)}
.hc-icon{font-size:20px;margin-bottom:8px}.hc-title{font-size:12.5px;font-weight:600;margin-bottom:3px}
.hc-desc{font-size:11px;color:var(--txt2);line-height:1.5}
.fgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:6px}
.frow{display:flex;align-items:center;gap:9px;background:var(--surf2);border:1px solid var(--bdr);border-radius:7px;padding:7px 11px}
.flbl{font-size:10.5px;color:var(--txt2);min-width:130px;flex-shrink:0}
.feq{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--teal-l);font-weight:500}
.ibox{background:var(--surf2);border:1px solid var(--bdr2);border-radius:var(--r);padding:10px 13px;margin-bottom:13px;font-size:11px;color:var(--txt2);line-height:1.6}
.ibox strong{color:var(--teal-l)}
.dc-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:9px}
.dc-u{background:var(--surf2);border:1px solid var(--bdr);border-radius:var(--r);padding:12px 14px}
.dc-ul{font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--txt3);margin-bottom:5px}
.dc-u input{height:38px;font-size:16px;font-weight:600;font-family:'JetBrains Mono',monospace;background:transparent;border:none;border-bottom:2px solid var(--bdr2);border-radius:0;padding:0 2px;color:var(--txt);width:100%;outline:none}
.dc-u input:focus{border-bottom-color:var(--teal)}
.dc-note{font-size:10px;color:var(--txt3);margin-top:4px;font-family:'JetBrains Mono',monospace}
.dc-sum{background:linear-gradient(135deg,rgba(34,168,144,.09),rgba(120,85,192,.06));border:1px solid rgba(34,168,144,.22);border-radius:var(--rl);padding:13px 17px;margin-top:12px}
.dc-sum h4{font-size:9.5px;color:var(--teal-l);text-transform:uppercase;letter-spacing:.08em;margin-bottom:5px}
.dc-phrase{font-size:16px;font-weight:600}
.qgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:6px}
.qc{background:var(--surf2);border:1px solid var(--bdr);border-radius:6px;padding:6px 10px;cursor:pointer;transition:all .13s}
.qc:hover{border-color:var(--teal);background:rgba(34,168,144,.06)}
.qc-n{font-size:9.5px;color:var(--txt3)}.qc-v{font-size:12px;font-weight:600;font-family:'JetBrains Mono',monospace}
.srow{display:flex;gap:6px;flex-wrap:wrap;margin-top:9px}
.sb2{height:29px;padding:0 12px;border-radius:14px;border:1px solid var(--bdr2);background:none;color:var(--txt2);font-size:11.5px;font-family:'Inter',sans-serif;cursor:pointer;transition:all .13s}
.sb2:hover{border-color:var(--teal-l);color:var(--teal-l)}
.sb2.del{border-color:rgba(217,88,88,.35);color:var(--red)}
.sb2.del:hover{border-color:var(--red)}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:3px}
</style>
</head>
<body>
<div class="layout">

<aside class="sidebar">
  <div class="sb-head">
    <div class="sb-badge">📈</div>
    <div><div class="sb-name">Increasing Annuities</div><div class="sb-unit">BMI 111 · Study Unit 6</div></div>
  </div>
  <nav class="sb-nav">
    <div class="sb-sec">Overview</div>
    <button class="nb on" onclick="go('home',this)"><span class="nb-ic">⊞</span>Dashboard</button>
    <div class="sb-sec">Calculators</div>
    <button class="nb" onclick="go('kth',this)"><span class="nb-ic">🔢</span>Every k-th Period</button>
    <button class="nb" onclick="go('next',this)"><span class="nb-ic">↗</span>Every Next Payment</button>
    <button class="nb" onclick="go('single',this)"><span class="nb-ic">💰</span>Single Investment</button>
    <button class="nb" onclick="go('infl',this)"><span class="nb-ic">📊</span>Inflation Adjustment</button>
    <div class="sb-sec">Tools</div>
    <button class="nb" onclick="go('date',this)"><span class="nb-ic">📅</span>Date Converter</button>
  </nav>
  <div class="sb-foot">NWU Centre for BMI<br>Assignment 6 · 2025</div>
</aside>

<div class="main">
  <div class="topbar">
    <div><div class="tb-title" id="tbt">Dashboard</div><div class="tb-sub" id="tbs">Increasing annuities — Study Unit 6</div></div>
    <div class="tb-pill" id="tbp">FV = X·[(1+r)^k−1]/r · [(1+r)^km−(1+j)^m]/[(1+r)^k−(1+j)]</div>
  </div>
  <div class="content">

    <!-- HOME -->
    <div class="panel on" id="panel-home">
      <div class="hero">
        <h1>Increasing Annuities Calculator</h1>
        <p>BMI 111 · Assignment 6 · NWU Centre for Business Mathematics and Informatics.<br>
           All answers verified against the assignment checklist. Every rate type supported including effective periodic, nominal, force of interest, and simple interest (any period).</p>
      </div>
      <div class="hgrid">
        <div class="hcard" onclick="go('kth')"><div class="hc-icon">🔢</div><div class="hc-title">Every k-th Period</div><div class="hc-desc">Payments level for k periods, then increase by j%. Arrears or advance. Solve FV, PV, or X.</div></div>
        <div class="hcard" onclick="go('next')"><div class="hc-icon">↗</div><div class="hc-title">Every Next Payment</div><div class="hc-desc">Each payment j% larger than the previous (k=1 special case). Solve FV, PV, or X.</div></div>
        <div class="hcard" onclick="go('single')"><div class="hc-icon">💰</div><div class="hc-title">Single Investment</div><div class="hc-desc">FV = PV·(1+r)^n. Any rate type. Solve FV, PV, n or r. With growth chart.</div></div>
        <div class="hcard" onclick="go('infl')"><div class="hc-icon">📊</div><div class="hc-title">Inflation Adjustment</div><div class="hc-desc">Amount_n = Amount_0·(1+f)^n. Shows real purchasing power over time.</div></div>
        <div class="hcard" onclick="go('date')"><div class="hc-icon">📅</div><div class="hc-title">Date Converter</div><div class="hc-desc">Years ↔ months ↔ weeks ↔ days. Natural breakdown. Send to any calculator.</div></div>
      </div>
      <div class="card">
        <div class="ch">Key Formulas — Study Unit 6</div>
        <div class="fgrid">
          <div class="frow"><span class="flbl">k-th period arrears FV</span><span class="feq">X·[(1+r)^k−1]/r · [(1+r)^km−(1+j)^m]/[(1+r)^k−(1+j)]</span></div>
          <div class="frow"><span class="flbl">k-th period advance FV</span><span class="feq">FV_arrears · (1+r)</span></div>
          <div class="frow"><span class="flbl">Any annuity PV</span><span class="feq">PV = FV · (1+r)^(−km)</span></div>
          <div class="frow"><span class="flbl">Every-next arrears FV</span><span class="feq">X · [(1+r)^n − (1+j)^n] / (r − j)</span></div>
          <div class="frow"><span class="flbl">Every-next advance FV</span><span class="feq">FV_arrears · (1+r)</span></div>
          <div class="frow"><span class="flbl">r (eff. per pay period)</span><span class="feq">r = (1+i_eff)^(1/p) − 1</span></div>
          <div class="frow"><span class="flbl">Single investment</span><span class="feq">FV = PV · (1+r)^n</span></div>
          <div class="frow"><span class="flbl">Inflation adjustment</span><span class="feq">Amount_n = Amount_0 · (1+f)^n</span></div>
          <div class="frow"><span class="flbl">Special case k=1</span><span class="feq">FV = X·[(1+r)^np−(1+j)^np]/(r−j)</span></div>
        </div>
      </div>
    </div>

    <!-- K-TH PERIOD -->
    <div class="panel" id="panel-kth">
      <div class="ibox"><strong>Every k-th period:</strong> Payments are the same for <em>k</em> consecutive periods, then increase by <em>j%</em>. There are <em>m</em> groups, giving <em>k×m</em> total payments. Example: monthly payments increasing annually → k=12, m=years.</div>
      <div class="card">
        <div class="ch">Timing &amp; Solve For</div>
        <div style="display:flex;gap:14px;flex-wrap:wrap">
          <div class="fg" style="min-width:185px"><label>Payment timing</label>
            <select id="kth-timing" onchange="doKth()">
              <option value="arrears">Arrears — end of period</option>
              <option value="advance">Advance — start of period</option>
            </select>
          </div>
          <div><label style="font-size:10.5px;font-weight:500;color:var(--txt2);display:block;margin-bottom:4px">Solve for</label>
            <div class="pills">
              <button class="pill on" onclick="setKth('fv',this)">Future Value</button>
              <button class="pill"    onclick="setKth('pv',this)">Present Value</button>
              <button class="pill"    onclick="setKth('x_fv',this)">X given FV</button>
              <button class="pill"    onclick="setKth('x_pv',this)">X given PV</button>
            </div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="ch">Interest Rate</div>
        <div class="ga">
          <div class="fg"><label>Rate value (%)</label><input type="number" id="kth-rate" value="8" step="0.0001" oninput="doKth()"></div>
          <div class="fg"><label>Rate type</label>
            <select id="kth-rtype" onchange="doKth()">
              <option value="nominal" selected>Nominal i^(p)</option>
              <option value="effective_annual">Effective annual (i)</option>
              <option value="effective_periodic">Effective periodic (i_p)</option>
              <option value="force">Force of interest (δ)</option>
              <option value="simple">Simple interest</option>
            </select>
          </div>
          <div class="fg"><label>Rate compounding period</label>
            <select id="kth-rper" onchange="doKth()">
              <option value="annual">Annual</option><option value="semi_annual">Semi-annual</option>
              <option value="quarterly" selected>Quarterly</option><option value="monthly">Monthly</option>
              <option value="weekly">Weekly</option><option value="daily">Daily</option>
            </select>
          </div>
          <div class="fg"><label>Payment period</label>
            <select id="kth-pper" onchange="doKth()">
              <option value="annual">Annual</option><option value="semi_annual">Semi-annual</option>
              <option value="quarterly" selected>Quarterly</option><option value="monthly">Monthly</option>
              <option value="weekly">Weekly</option><option value="daily">Daily</option>
            </select>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="ch">Annuity Parameters</div>
        <div class="ga">
          <div class="fg"><label>First payment X (R) <span class="bi" id="kth-xb">INPUT</span></label><input type="number" id="kth-x" value="10000" oninput="doKth()"></div>
          <div class="fg"><label>Increase per group j (%)</label><input type="number" id="kth-j" value="5" step="0.0001" oninput="doKth()"></div>
          <div class="fg"><label>Periods per group k</label><input type="number" id="kth-k" value="4" min="1" step="1" oninput="doKth()"></div>
          <div class="fg"><label>Number of groups m</label><input type="number" id="kth-m" value="10" min="1" step="1" oninput="doKth()"></div>
          <div class="fg"><label>Future Value (R) <span class="bi" id="kth-fvb">INPUT</span></label><input type="number" id="kth-fv" placeholder="—" oninput="doKth()"></div>
          <div class="fg"><label>Present Value (R) <span class="bi" id="kth-pvb">INPUT</span></label><input type="number" id="kth-pv" placeholder="—" oninput="doKth()"></div>
        </div>
        <div class="err" id="kth-err"></div>
      </div>
      <div class="tiles" id="kth-tiles"></div>
      <div class="card"><div class="ch">FV contribution per level group</div>
        <div class="chbox"><canvas id="kth-chart"></canvas></div>
      </div>
    </div>

    <!-- EVERY NEXT PAYMENT -->
    <div class="panel" id="panel-next">
      <div class="ibox"><strong>Every-next-payment:</strong> Each payment is <em>j%</em> larger than the previous one. This is the special case k=1 of the k-th period formula: <em>FV = X·[(1+r)^n−(1+j)^n]/(r−j)</em>.</div>
      <div class="card">
        <div class="ch">Timing &amp; Solve For</div>
        <div style="display:flex;gap:14px;flex-wrap:wrap">
          <div class="fg" style="min-width:185px"><label>Payment timing</label>
            <select id="nxt-timing" onchange="doNext()">
              <option value="arrears">Arrears — end of period</option>
              <option value="advance">Advance — start of period</option>
            </select>
          </div>
          <div><label style="font-size:10.5px;font-weight:500;color:var(--txt2);display:block;margin-bottom:4px">Solve for</label>
            <div class="pills">
              <button class="pill on" onclick="setNext('fv',this)">Future Value</button>
              <button class="pill"    onclick="setNext('pv',this)">Present Value</button>
              <button class="pill"    onclick="setNext('x_fv',this)">X given FV</button>
              <button class="pill"    onclick="setNext('x_pv',this)">X given PV</button>
            </div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="ch">Interest Rate</div>
        <div class="ga">
          <div class="fg"><label>Rate value (%)</label><input type="number" id="nxt-rate" value="8" step="0.0001" oninput="doNext()"></div>
          <div class="fg"><label>Rate type</label>
            <select id="nxt-rtype" onchange="doNext()">
              <option value="nominal" selected>Nominal i^(p)</option>
              <option value="effective_annual">Effective annual (i)</option>
              <option value="effective_periodic">Effective periodic (i_p)</option>
              <option value="force">Force of interest (δ)</option>
              <option value="simple">Simple interest</option>
            </select>
          </div>
          <div class="fg"><label>Rate compounding period</label>
            <select id="nxt-rper" onchange="doNext()">
              <option value="annual">Annual</option><option value="semi_annual">Semi-annual</option>
              <option value="quarterly" selected>Quarterly</option><option value="monthly">Monthly</option>
              <option value="weekly">Weekly</option><option value="daily">Daily</option>
            </select>
          </div>
          <div class="fg"><label>Payment period</label>
            <select id="nxt-pper" onchange="doNext()">
              <option value="annual">Annual</option><option value="semi_annual">Semi-annual</option>
              <option value="quarterly" selected>Quarterly</option><option value="monthly">Monthly</option>
              <option value="weekly">Weekly</option><option value="daily">Daily</option>
            </select>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="ch">Annuity Parameters</div>
        <div class="ga">
          <div class="fg"><label>First payment X (R) <span class="bi" id="nxt-xb">INPUT</span></label><input type="number" id="nxt-x" value="10000" oninput="doNext()"></div>
          <div class="fg"><label>Increase per payment j (%)</label><input type="number" id="nxt-j" value="1" step="0.0001" oninput="doNext()"></div>
          <div class="fg"><label>Total payments n</label><input type="number" id="nxt-n" value="40" min="1" step="1" oninput="doNext()"></div>
          <div class="fg"><label>Future Value (R) <span class="bi" id="nxt-fvb">INPUT</span></label><input type="number" id="nxt-fv" placeholder="—" oninput="doNext()"></div>
          <div class="fg"><label>Present Value (R) <span class="bi" id="nxt-pvb">INPUT</span></label><input type="number" id="nxt-pv" placeholder="—" oninput="doNext()"></div>
        </div>
        <div class="err" id="nxt-err"></div>
      </div>
      <div class="tiles" id="nxt-tiles"></div>
      <div class="card"><div class="ch">Payment escalation over time</div>
        <div class="chbox"><canvas id="nxt-chart"></canvas></div>
      </div>
    </div>

    <!-- SINGLE INVESTMENT -->
    <div class="panel" id="panel-single">
      <div class="card">
        <div class="ch">Solve For</div>
        <div class="pills">
          <button class="pill on" onclick="setSI('fv',this)">Future Value</button>
          <button class="pill"    onclick="setSI('pv',this)">Present Value</button>
          <button class="pill"    onclick="setSI('n',this)">Periods n</button>
          <button class="pill"    onclick="setSI('r',this)">Rate r</button>
        </div>
      </div>
      <div class="card">
        <div class="ch">Inputs</div>
        <div class="ga">
          <div class="fg"><label>Rate value (%) <span class="bi" id="si-rb">INPUT</span></label><input type="number" id="si-rate" value="8" step="0.0001" oninput="doSI()"></div>
          <div class="fg"><label>Rate type</label>
            <select id="si-rtype" onchange="doSI()">
              <option value="nominal" selected>Nominal i^(p)</option>
              <option value="effective_annual">Effective annual (i)</option>
              <option value="effective_periodic">Effective periodic (i_p)</option>
              <option value="force">Force of interest (δ)</option>
              <option value="simple">Simple interest</option>
            </select>
          </div>
          <div class="fg"><label>Rate period</label>
            <select id="si-rper" onchange="doSI()">
              <option value="annual">Annual</option><option value="semi_annual">Semi-annual</option>
              <option value="quarterly">Quarterly</option><option value="monthly" selected>Monthly</option>
              <option value="weekly">Weekly</option><option value="daily">Daily</option>
            </select>
          </div>
          <div class="fg"><label>Compounding / payment period</label>
            <select id="si-pper" onchange="doSI()">
              <option value="annual">Annual</option><option value="semi_annual">Semi-annual</option>
              <option value="quarterly">Quarterly</option><option value="monthly" selected>Monthly</option>
              <option value="weekly">Weekly</option><option value="daily">Daily</option>
            </select>
          </div>
          <div class="fg"><label>Present Value (R) <span class="bi" id="si-pvb">INPUT</span></label><input type="number" id="si-pv" value="30000" oninput="doSI()"></div>
          <div class="fg"><label>Future Value (R) <span class="bi" id="si-fvb">INPUT</span></label><input type="number" id="si-fv" placeholder="—" oninput="doSI()"></div>
          <div class="fg"><label>Periods n <span class="bi" id="si-nb">INPUT</span></label><input type="number" id="si-n" value="84" oninput="doSI()"></div>
        </div>
        <div class="err" id="si-err"></div>
      </div>
      <div class="tiles" id="si-tiles"></div>
      <div class="card"><div class="ch">Investment growth — FV = PV·(1+r)^n</div>
        <div class="chbox"><canvas id="si-chart"></canvas></div>
      </div>
    </div>

    <!-- INFLATION -->
    <div class="panel" id="panel-infl">
      <div class="card">
        <div class="ch">Inflation adjustment — Amount × (1+f)^n</div>
        <div class="ga">
          <div class="fg"><label>Starting amount (R)</label><input type="number" id="infl-amt" value="8000" oninput="doInfl()"></div>
          <div class="fg"><label>Annual inflation rate (%)</label><input type="number" id="infl-f" value="6" step="0.01" oninput="doInfl()"></div>
          <div class="fg"><label>Number of years</label><input type="number" id="infl-n" value="9" step="0.5" oninput="doInfl()"></div>
        </div>
        <div class="err" id="infl-err"></div>
      </div>
      <div class="tiles" id="infl-tiles"></div>
      <div class="card"><div class="ch">Inflation-adjusted value over time</div>
        <div class="chbox"><canvas id="infl-chart"></canvas></div>
      </div>
    </div>

    <!-- DATE CONVERTER -->
    <div class="panel" id="panel-date">
      <div class="card">
        <div class="ch">Enter a value in any one unit</div>
        <div class="dc-grid">
          <div class="dc-u"><div class="dc-ul">Years</div><input type="number" id="dc-y" placeholder="e.g. 10" step="0.001" oninput="dcCalc('years',this.value)"><div class="dc-note" id="dc-yn">—</div></div>
          <div class="dc-u"><div class="dc-ul">Months</div><input type="number" id="dc-m" placeholder="e.g. 120" step="0.01" oninput="dcCalc('months',this.value)"><div class="dc-note" id="dc-mn">—</div></div>
          <div class="dc-u"><div class="dc-ul">Weeks</div><input type="number" id="dc-w" placeholder="e.g. 520" step="0.1" oninput="dcCalc('weeks',this.value)"><div class="dc-note" id="dc-wn">—</div></div>
          <div class="dc-u"><div class="dc-ul">Days</div><input type="number" id="dc-d" placeholder="e.g. 3650" step="1" oninput="dcCalc('days',this.value)"><div class="dc-note" id="dc-dn">—</div></div>
        </div>
        <div class="dc-sum" id="dc-sum" style="display:none">
          <h4>Natural breakdown</h4>
          <div class="dc-phrase" id="dc-ph">—</div>
        </div>
      </div>
      <div class="card">
        <div class="ch">Quick reference — click to load</div>
        <div class="qgrid">
          <div class="qc" onclick="dcLoad(0.25)"><div class="qc-n">1 Quarter</div><div class="qc-v">0.25 yrs</div></div>
          <div class="qc" onclick="dcLoad(0.5)"><div class="qc-n">6 Months</div><div class="qc-v">0.5 yrs</div></div>
          <div class="qc" onclick="dcLoad(1)"><div class="qc-n">1 Year</div><div class="qc-v">1 yr</div></div>
          <div class="qc" onclick="dcLoad(2)"><div class="qc-n">2 Years</div><div class="qc-v">2 yrs</div></div>
          <div class="qc" onclick="dcLoad(5)"><div class="qc-n">5 Years</div><div class="qc-v">5 yrs</div></div>
          <div class="qc" onclick="dcLoad(10)"><div class="qc-n">10 Years</div><div class="qc-v">10 yrs</div></div>
          <div class="qc" onclick="dcLoad(15)"><div class="qc-n">15 Years</div><div class="qc-v">15 yrs</div></div>
          <div class="qc" onclick="dcLoad(20)"><div class="qc-n">20 Years</div><div class="qc-v">20 yrs</div></div>
          <div class="qc" onclick="dcLoad(30)"><div class="qc-n">30 Years</div><div class="qc-v">30 yrs</div></div>
          <div class="qc" onclick="dcLoad(40)"><div class="qc-n">40 Years</div><div class="qc-v">40 yrs</div></div>
        </div>
      </div>
      <div class="card">
        <div class="ch">Send year value to calculator</div>
        <div class="srow">
          <button class="sb2" onclick="dcSend('kth-k')">→ k-th: k periods</button>
          <button class="sb2" onclick="dcSend('kth-m')">→ k-th: m groups</button>
          <button class="sb2" onclick="dcSend('nxt-n')">→ Next: n payments</button>
          <button class="sb2" onclick="dcSend('si-n')">→ Single: n periods</button>
          <button class="sb2" onclick="dcSend('infl-n')">→ Inflation: years</button>
          <button class="sb2 del" onclick="dcClear()">✕ Clear</button>
        </div>
      </div>
    </div>

  </div>
</div>
</div>

<script>
const META={
  home:  {t:'Dashboard',          s:'Increasing annuities — Study Unit 6 · BMI 111',         p:'FV = X · A · B'},
  kth:   {t:'Every k-th Period',  s:'Level for k periods then increases by j% — m groups',   p:'FV = X·[(1+r)^k−1]/r · [(1+r)^km−(1+j)^m]/[(1+r)^k−(1+j)]'},
  next:  {t:'Every Next Payment', s:'Each payment j% larger than the previous one (k=1)',    p:'FV = X·[(1+r)^n−(1+j)^n]/(r−j)'},
  single:{t:'Single Investment',  s:'FV = PV·(1+r)^n — any rate type, any period',           p:'FV = PV·(1+r)^n'},
  infl:  {t:'Inflation Adjustment',s:'Real purchasing-power adjustment over time',            p:'Amount_n = Amount_0·(1+f)^n'},
  date:  {t:'Date Converter',     s:'Years ↔ months ↔ weeks ↔ days — send to any panel',    p:'1 yr = 12 mths = 52.1775 wks = 365.25 days'},
};
function go(id,btn){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.nb').forEach(b=>b.classList.remove('on'));
  document.getElementById('panel-'+id).classList.add('on');
  if(btn) btn.classList.add('on');
  else{ const b=[...document.querySelectorAll('.nb')].find(b=>b.getAttribute('onclick')?.includes("'"+id+"'")); if(b) b.classList.add('on'); }
  const m=META[id];
  document.getElementById('tbt').textContent=m.t;
  document.getElementById('tbs').textContent=m.s;
  document.getElementById('tbp').textContent=m.p;
}
const R  = v=>v==null?'—':'R\u202f'+Number(v).toLocaleString('en-ZA',{minimumFractionDigits:2,maximumFractionDigits:2});
const Pc = (v,d=4)=>v==null?'—':Number(v).toFixed(d)+'%';
const Fx = (v,d=2)=>v==null?'—':Number(v).toFixed(d);
const $  = id=>document.getElementById(id);
async function post(url,data){const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});return r.json();}
const PU='rgba(120,85,192,1)',TL='rgba(34,168,144,1)',GD='rgba(223,168,60,1)';
const PUA='rgba(120,85,192,0.18)',TLA='rgba(34,168,144,0.14)',GDA='rgba(223,168,60,0.16)';
const GR='rgba(255,255,255,0.045)',TK='#5a5a7a';
const charts={};

function mkBar(id,labels,vals){
  if(charts[id]) charts[id].destroy();
  charts[id]=new Chart($(id),{type:'bar',data:{labels,datasets:[{label:'FV per group',data:vals,backgroundColor:PU,borderWidth:0,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{y:{ticks:{color:TK,callback:v=>'R'+Number(v).toLocaleString()},grid:{color:GR}},x:{ticks:{color:TK,maxTicksLimit:20},grid:{display:false}}}}});
}
function mkLine(id,labels,ds){
  if(charts[id]) charts[id].destroy();
  charts[id]=new Chart($(id),{type:'line',data:{labels,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:'#888',usePointStyle:true,pointStyle:'line'}}},
      scales:{y:{ticks:{color:TK,callback:v=>'R'+Number(v).toLocaleString()},grid:{color:GR}},x:{ticks:{color:TK,maxTicksLimit:10,autoSkip:true},grid:{display:false}}}}});
}

// ── K-TH PERIOD ─────────────────────────
let kthSolve='fv';
function setKth(s,btn){
  kthSolve=s;
  document.querySelectorAll('#panel-kth .pill').forEach(b=>b.classList.remove('on')); btn.classList.add('on');
  ['kth-xb','kth-fvb','kth-pvb'].forEach(id=>{$(id).textContent='INPUT';$(id).className='bi';});
  ['kth-x','kth-fv','kth-pv'].forEach(id=>{const el=$(id);el.classList.remove('out');el.readOnly=false;el.placeholder='—';});
  const outMap={fv:'kth-fv',pv:'kth-pv',x_fv:'kth-x',x_pv:'kth-x'};
  const badgeMap={fv:'kth-fvb',pv:'kth-pvb',x_fv:'kth-xb',x_pv:'kth-xb'};
  const el=$(outMap[s]); el.classList.add('out');el.readOnly=true;el.value='';el.placeholder='Calculated';
  $(badgeMap[s]).textContent='OUTPUT';$(badgeMap[s]).className='bo';
  doKth();
}
async function doKth(){
  const res=await post('/api/inc_kth',{solve:kthSolve,timing:$('kth-timing').value,
    rate:$('kth-rate').value,rate_type:$('kth-rtype').value,
    rate_period:$('kth-rper').value,pay_period:$('kth-pper').value,
    j:$('kth-j').value,k:$('kth-k').value,m:$('kth-m').value,
    x:$('kth-x').value,fv:$('kth-fv').value,pv:$('kth-pv').value});
  const err=$('kth-err');
  if(res.error){err.textContent=res.error;err.style.display='block';return;}
  err.style.display='none';
  const outMap={fv:'kth-fv',pv:'kth-pv',x_fv:'kth-x',x_pv:'kth-x'};
  const valMap={fv:res.fv,pv:res.pv,x_fv:res.x,x_pv:res.x};
  const v=valMap[kthSolve]; $(outMap[kthSolve]).value=v!=null?Number(v).toFixed(4):'';
  $('kth-tiles').innerHTML=`
    <div class="tile hi"><div class="tl">Future Value</div><div class="tv teal">${R(res.fv)}</div></div>
    <div class="tile hi"><div class="tl">Present Value</div><div class="tv purple">${R(res.pv)}</div></div>
    <div class="tile"><div class="tl">First payment X</div><div class="tv">${R(res.x)}</div></div>
    <div class="tile"><div class="tl">r per pay period</div><div class="tv">${Pc(res.r,6)}</div></div>
    <div class="tile"><div class="tl">Increase j</div><div class="tv">${Pc(res.j,4)}</div></div>
    <div class="tile"><div class="tl">Total payments k×m</div><div class="tv">${res.n_total}</div></div>
    <div class="tile"><div class="tl">Total deposited</div><div class="tv">${R(res.total_payments)}</div></div>
    <div class="tile"><div class="tl">Interest earned</div><div class="tv gold">${R(res.interest_earned)}</div></div>`;
  if(res.chart_groups?.length) mkBar('kth-chart',res.chart_groups,res.chart_fv);
}
doKth();

// ── EVERY NEXT ───────────────────────────
let nxtSolve='fv';
function setNext(s,btn){
  nxtSolve=s;
  document.querySelectorAll('#panel-next .pill').forEach(b=>b.classList.remove('on')); btn.classList.add('on');
  ['nxt-xb','nxt-fvb','nxt-pvb'].forEach(id=>{$(id).textContent='INPUT';$(id).className='bi';});
  ['nxt-x','nxt-fv','nxt-pv'].forEach(id=>{const el=$(id);el.classList.remove('out');el.readOnly=false;el.placeholder='—';});
  const outMap={fv:'nxt-fv',pv:'nxt-pv',x_fv:'nxt-x',x_pv:'nxt-x'};
  const badgeMap={fv:'nxt-fvb',pv:'nxt-pvb',x_fv:'nxt-xb',x_pv:'nxt-xb'};
  const el=$(outMap[s]); el.classList.add('out');el.readOnly=true;el.value='';el.placeholder='Calculated';
  $(badgeMap[s]).textContent='OUTPUT';$(badgeMap[s]).className='bo';
  doNext();
}
async function doNext(){
  const res=await post('/api/inc_next',{solve:nxtSolve,timing:$('nxt-timing').value,
    rate:$('nxt-rate').value,rate_type:$('nxt-rtype').value,
    rate_period:$('nxt-rper').value,pay_period:$('nxt-pper').value,
    j:$('nxt-j').value,n:$('nxt-n').value,
    x:$('nxt-x').value,fv:$('nxt-fv').value,pv:$('nxt-pv').value});
  const err=$('nxt-err');
  if(res.error){err.textContent=res.error;err.style.display='block';return;}
  err.style.display='none';
  const outMap={fv:'nxt-fv',pv:'nxt-pv',x_fv:'nxt-x',x_pv:'nxt-x'};
  const valMap={fv:res.fv,pv:res.pv,x_fv:res.x,x_pv:res.x};
  const v=valMap[nxtSolve]; $(outMap[nxtSolve]).value=v!=null?Number(v).toFixed(4):'';
  $('nxt-tiles').innerHTML=`
    <div class="tile hi"><div class="tl">Future Value</div><div class="tv teal">${R(res.fv)}</div></div>
    <div class="tile hi"><div class="tl">Present Value</div><div class="tv purple">${R(res.pv)}</div></div>
    <div class="tile"><div class="tl">First payment X</div><div class="tv">${R(res.x)}</div></div>
    <div class="tile"><div class="tl">r per pay period</div><div class="tv">${Pc(res.r,6)}</div></div>
    <div class="tile"><div class="tl">Increase j per pmt</div><div class="tv">${Pc(res.j,4)}</div></div>
    <div class="tile"><div class="tl">Total payments n</div><div class="tv">${res.n}</div></div>
    <div class="tile"><div class="tl">Total deposited</div><div class="tv">${R(res.total_payments)}</div></div>
    <div class="tile"><div class="tl">Interest earned</div><div class="tv gold">${R(res.interest_earned)}</div></div>`;
  if(res.chart_pmt?.length) mkLine('nxt-chart',res.chart_pmt,[
    {label:'Payment amount',data:res.chart_val,borderColor:TL,backgroundColor:TLA,fill:true,tension:0.3,pointRadius:0,borderWidth:2}]);
}
doNext();

// ── SINGLE INVESTMENT ────────────────────
let siSolve='fv';
const SI_OUT={fv:'si-fv',pv:'si-pv',n:'si-n',r:'si-rate'};
const SI_BADGE={fv:'si-fvb',pv:'si-pvb',n:'si-nb',r:'si-rb'};
function setSI(s,btn){
  siSolve=s;
  document.querySelectorAll('#panel-single .pill').forEach(b=>b.classList.remove('on')); btn.classList.add('on');
  Object.values(SI_BADGE).forEach(id=>{$(id).textContent='INPUT';$(id).className='bi';});
  Object.values(SI_OUT).forEach(id=>{const el=$(id);el.classList.remove('out');el.readOnly=false;el.placeholder='—';});
  const el=$(SI_OUT[s]); el.classList.add('out');el.readOnly=true;el.value='';el.placeholder='Calculated';
  $(SI_BADGE[s]).textContent='OUTPUT';$(SI_BADGE[s]).className='bo';
  doSI();
}
async function doSI(){
  const res=await post('/api/single',{solve:siSolve,
    rate:$('si-rate').value,rate_type:$('si-rtype').value,
    rate_period:$('si-rper').value,pay_period:$('si-pper').value,
    pv:$('si-pv').value,fv:$('si-fv').value,n:$('si-n').value});
  const err=$('si-err');
  if(res.error){err.textContent=res.error;err.style.display='block';return;}
  err.style.display='none';
  const vals={fv:res.fv,pv:res.pv,n:res.n,r:res.r};
  const v=vals[siSolve]; $(SI_OUT[siSolve]).value=v!=null?Number(v).toFixed(6):'';
  $('si-tiles').innerHTML=`
    <div class="tile hi"><div class="tl">Future Value</div><div class="tv teal">${R(res.fv)}</div></div>
    <div class="tile hi"><div class="tl">Present Value</div><div class="tv purple">${R(res.pv)}</div></div>
    <div class="tile"><div class="tl">Interest earned</div><div class="tv gold">${R(res.interest)}</div></div>
    <div class="tile"><div class="tl">r per period</div><div class="tv">${Pc(res.r,6)}</div></div>
    <div class="tile"><div class="tl">Periods n</div><div class="tv">${Fx(res.n,2)}</div></div>`;
  if(res.chart_t?.length) mkLine('si-chart',res.chart_t,[
    {label:'Investment value',data:res.chart_fv,borderColor:PU,backgroundColor:PUA,fill:true,tension:0.4,pointRadius:0,borderWidth:2}]);
}
doSI();

// ── INFLATION ─────────────────────────────
async function doInfl(){
  const amt=parseFloat($('infl-amt').value)||0;
  const res=await post('/api/inflation',{amount:$('infl-amt').value,infl:$('infl-f').value,years:$('infl-n').value});
  const err=$('infl-err');
  if(res.error){err.textContent=res.error;err.style.display='block';return;}
  err.style.display='none';
  const mult=res.result&&amt?res.result/amt:null;
  $('infl-tiles').innerHTML=`
    <div class="tile hi"><div class="tl">Adjusted amount</div><div class="tv teal">${R(res.result)}</div></div>
    <div class="tile"><div class="tl">Original amount</div><div class="tv">${R(amt)}</div></div>
    <div class="tile"><div class="tl">Increase</div><div class="tv gold">${R(res.result&&amt?res.result-amt:null)}</div></div>
    <div class="tile"><div class="tl">Growth multiplier</div><div class="tv">${mult!=null?mult.toFixed(4)+'×':'—'}</div></div>`;
  if(res.chart_t?.length) mkLine('infl-chart',res.chart_t,[
    {label:'Inflation-adjusted value',data:res.chart_val,borderColor:GD,backgroundColor:GDA,fill:true,tension:0.3,pointRadius:0,borderWidth:2}]);
}
doInfl();

// ── DATE CONVERTER ───────────────────────
let dcYrs=null;
async function dcCalc(unit,value){
  if(!value||isNaN(value)) return;
  const res=await post('/api/date',{unit,value});
  if(res.error) return;
  dcYrs=res.years;
  $('dc-y').value=unit==='years'  ?value:res.years.toFixed(6);
  $('dc-m').value=unit==='months' ?value:res.months.toFixed(4);
  $('dc-w').value=unit==='weeks'  ?value:res.weeks.toFixed(4);
  $('dc-d').value=unit==='days'   ?value:res.days.toFixed(2);
  $('dc-yn').textContent=res.years.toFixed(6)+' years';
  $('dc-mn').textContent=res.months.toFixed(4)+' months';
  $('dc-wn').textContent=res.weeks.toFixed(4)+' weeks';
  $('dc-dn').textContent=res.days.toFixed(2)+' days';
  $('dc-ph').textContent=res.summary;
  $('dc-sum').style.display='block';
}
function dcLoad(y){$('dc-y').value=y; dcCalc('years',y);}
function dcClear(){
  ['dc-y','dc-m','dc-w','dc-d'].forEach(id=>$(id).value='');
  ['dc-yn','dc-mn','dc-wn','dc-dn'].forEach(id=>$(id).textContent='—');
  $('dc-sum').style.display='none'; dcYrs=null;
}
function dcSend(targetId){
  if(dcYrs==null) return;
  // Use the raw years value, but for k/m/n it should be a round integer
  const v=Number(dcYrs).toFixed(4);
  $(targetId).value=v;
  if(targetId.startsWith('kth')) doKth();
  else if(targetId.startsWith('nxt')) doNext();
  else if(targetId.startsWith('si')) doSI();
  else if(targetId.startsWith('infl')) doInfl();
}
</script>
</body></html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

def _open():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:5051")

if __name__ == "__main__":
    print("\n" + "═"*56)
    print("  Increasing Annuities Calculator — BMI 111")
    print("  Assignment 6  ·  Study Unit 6  ·  NWU")
    print("═"*56)
    print("  ▶  Open browser at  http://127.0.0.1:5051")
    print("  ■  Stop server:     Ctrl+C")
    print("═"*56 + "\n")
    threading.Thread(target=_open, daemon=True).start()
    app.run(host="127.0.0.1", port=5051, debug=False)
