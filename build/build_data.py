#!/usr/bin/env python3
"""Build the drill-down revenue and expense data for the USF EM finance site.

Usage: python3 -I build_data.py <uploads_root> <revenue.csv> <umsa_monthly.csv> <out.json>
<uploads_root> holds the staged Box folder "_USF Emergency Medicine Department Folder"
and the iCloud folder "com~apple~CloudDocs". Nothing here is secret. The output JSON is
encrypted by encrypt.py before it goes into the repo.
"""
import sys, re, csv, json, datetime, warnings, collections
import openpyxl
warnings.filterwarnings("ignore")
ROOT, REV_CSV, UMSA_CSV, OUT = sys.argv[1:5]
BOX = ROOT + "/_USF Emergency Medicine Department Folder/"
ICL = ROOT + "/com~apple~CloudDocs/Emergency Medicine/"
F_MODEL = BOX + "Contracts/TeamHealth Contract/FY27 TH Contract Build/FINAL Emergency Medicine TeamHealth Contract FY2027- New Model.xlsx"
F_RUN = BOX + "Contracts/TeamHealth Contract/FY27 TH Contract Build/Running Expenses- FY 27 Budget.xlsx"
F_PERS = BOX + "Budget/FY2027/Budget Workbooks with Notes FY27/FY27 Personnel Workbook_Emer Med.xlsx"
F_TRK = {"faculty": BOX + "Payroll/Salary Tracking/EM Faculty Salary Tracking-FY2027.xlsx",
         "research": BOX + "Payroll/Salary Tracking/EM Faculty Salary Tracking - Research Team-FY2027.xlsx",
         "admin": BOX + "Payroll/Salary Tracking/EM Admin Salary Tracking-FY2027.xlsx"}
F_ASF = ICL + "ASF Monthly Pay Tracker/ASF_Monthly_Pay_FY27.xlsx"
S_MODEL = "Box, Contracts/TeamHealth Contract/FY27 TH Contract Build/FINAL Emergency Medicine TeamHealth Contract FY2027- New Model.xlsx (Dashboard tab)"
S_TRK = "Box, Payroll/Salary Tracking (FY2027 tracking workbooks, pay periods 2701 on)"
S_ASF = "iCloud, Emergency Medicine/ASF Monthly Pay Tracker/ASF_Monthly_Pay_FY27.xlsx (built from the 2600 ASF payroll files in Box)"
S_PERS = "Box, Budget/FY2027/Budget Workbooks with Notes FY27/FY27 Personnel Workbook_Emer Med.xlsx"
S_RUN = "Box, Contracts/TeamHealth Contract/FY27 TH Contract Build/Running Expenses- FY 27 Budget.xlsx"
S_ALL = "Box, Budget/FY2027/Budget Workbooks with Notes FY27/FY27 All-Source WORKBOOK_Emer Med_Notes.xlsx (SUMMARY tab)"
PERIODS = 26

def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0

def node(name, b=None, a=None, p=None, kids=None, **meta):
    n = {"n": name, "v": [b, a, p]}
    if kids: n["k"] = kids
    n.update({k: v for k, v in meta.items() if v})
    return n

def roll(n):
    """Fill missing values of a node from its children (sum), bottom up."""
    kids = n.get("k", [])
    for c in kids: roll(c)
    for i in range(3):
        if n["v"][i] is None and kids:
            vals = [c["v"][i] for c in kids if c["v"][i] is not None]
            n["v"][i] = round(sum(vals), 2) if vals else None
        elif n["v"][i] is not None:
            n["v"][i] = round(n["v"][i], 2)
    return n

def lastkey(name):
    return re.sub(r"[^a-z]", "", name.split(",")[0].lower()).replace("ubelacker", "uebelacker")

# ---------- A. TeamHealth contract model, one row per physician ----------
wb = openpyxl.load_workbook(F_MODEL, data_only=True)
ws = wb["Dashboard"]
hdr = [str(c).strip() if c is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
col = {h: i for i, h in enumerate(hdr) if h}
def mc(row, label):
    for h, i in col.items():
        if h.startswith(label): return num(row[i])
    raise KeyError(label)
model = {}
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[0] or str(r[0]).startswith("Grand Total"): break
    model[str(r[0]).strip()] = r

# ---------- B. ASF tracker, by physician and pay code ----------
wb = openpyxl.load_workbook(F_ASF, data_only=True)
status = {}
bym = list(wb["FY27 by Month"].iter_rows(values_only=True))
hrow = next(i for i, r in enumerate(bym) if r[0] == "Physician")
for j, m in enumerate(bym[hrow][1:13], 1):
    status[m] = str(bym[hrow + 1][j]).upper().startswith("ACTUAL")
months_actual = [m for m, s in status.items() if s]
asf = collections.defaultdict(lambda: collections.defaultdict(lambda: [0.0, 0.0]))  # person -> code -> [actual, all]
MON = {"Jul": "Jul 2026", "Aug": "Aug 2026", "Sep": "Sep 2026", "Oct": "Oct 2026", "Nov": "Nov 2026", "Dec": "Dec 2026",
       "Jan": "Jan 2027", "Feb": "Feb 2027", "Mar": "Mar 2027", "Apr": "Apr 2027", "May": "May 2027", "Jun": "Jun 2027"}
for s in wb.worksheets:
    m = re.match(r"^\d{4}-\d{2} (\w{3})$", s.title)
    if not m: continue
    rows = list(s.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r[0] == "Physician")
    codes = rows[hi]
    tot_i = next(i for i, c in enumerate(codes) if c and str(c).startswith("Total ASF"))
    act = status.get(MON[m.group(1)], False)
    for r in rows[hi + 1:]:
        if not r[0] or str(r[0]).strip() == "Total": break
        for i in range(1, tot_i):
            v = num(r[i])
            if v:
                asf[str(r[0]).strip()][str(codes[i])][1] += v
                if act: asf[str(r[0]).strip()][str(codes[i])][0] += v

# ---------- C. USF salary tracking by person and chartfield ----------
def fund_of(cf):
    cf = str(cf)
    f = re.search(r"(?<!\d)(\d{5})(?!\d)", cf)
    fund = f.group(1) if f else "?"
    proj = re.search(r"(\d{9,11})\s*$", cf)
    proj = proj.group(1).lstrip("0") if proj else ""
    if fund == "94533": return "UMSA 2600 convenience account (USF salary funded by the practice plan)", "TeamHealth contract and other UMSA 2600 revenue"
    if fund == "94512": return "GME fund 94512", "GME program director, coordinator and core faculty support"
    if fund == "94513": return "HPCC fund 94513 (outside EM)", "Another USF Health unit"
    if fund == "94540": return "Faculty Affairs fund 94540 (outside EM)", "Another USF Health unit"
    if fund.startswith("94"): return "Other UMSA fund " + fund, "UMSA"
    if fund == "11000": return "E&G teaching (EVU), fund 11000", "UME EVU allocation"
    if fund == "10000": return ("E&G Preeminence (HSCPRE), fund 10000" if "HSCPRE" in cf else "E&G fund 10000"), "College support, state side"
    if fund == "10009":
        if "MEDH" in cf: return "Carry forward 10009, hyperbaric program (" + re.search(r"MEDH\d+", cf).group(0) + ")", "Carry forward, state side"
        if "HSCPRE" in cf: return "Carry forward 10009, Preeminence start up", "College support, state side"
        return "Carry forward 10009, Chair letter of offer", "College support, state side"
    if fund in ("20000", "21000", "22000"):
        if proj.startswith("888001810") or proj.startswith("88801810"): return "Florida Center for EMS grant", "Florida Center for EMS (FDOH)"
        if proj.startswith("8880018700"): return "CFBHN project", "CFBHN through TGH"
        kind = "Clinical trial or contract" if fund == "22000" else "Sponsored grant"
        return kind + " project " + (proj or "?"), "Sponsored research"
    if fund == "01576": return "EM Research auxiliary 01576", "TGH research support, FOCUS and CFBHN invoices"
    if fund == "18300": return "RIA fund 18300", "F&A return"
    if fund.startswith("01"): return "Auxiliary fund " + fund, "Auxiliary"
    return "Fund " + fund, ""

track = {}  # (grp, sheet) -> {"funds": {name: [ytd, rate, src]}, "last": period index}
gmax = 0
for grp, f in F_TRK.items():
    wb = openpyxl.load_workbook(f, data_only=True)
    for s in wb.worksheets:
        if s.title == "Blank": continue
        rows = list(s.iter_rows(values_only=True))
        if not rows: continue
        seen, prow = set(), []
        for r in rows:
            if isinstance(r[0], (int, float)) and 2700 < r[0] <= 2700 + PERIODS and r[0] not in seen:
                seen.add(r[0]); prow.append(r)
        if not prow: continue
        head = rows[:3]
        funds = {}
        width = max(len(r) for r in rows)
        for c in range(3, width):
            cells = [str(h[c]) for h in head if c < len(h) and h[c] is not None]
            cf = next((x for x in cells if re.search(r"\d{5}", x) and "-" in x), None)
            if not cf or any(x.strip().lower().startswith("total") for x in cells): continue
            name, src = fund_of(cf)
            vals = [(int(r[0]) - 2700, num(r[c]) if c < len(r) else 0.0) for r in prow]
            d = funds.setdefault(name, {"src": src, "by": collections.defaultdict(float)})
            for i, v in vals: d["by"][i] += v
        allp = collections.defaultdict(float)
        for d in funds.values():
            for i, v in d["by"].items(): allp[i] += v
        paid = [i for i, v in allp.items() if abs(v) > 0.005]
        if not paid: continue
        last = max(paid); gmax = max(gmax, last)
        track[(grp, s.title)] = {"funds": funds, "last": last}

def track_nodes(t):
    """Children per fund with actual to date and a projection at the latest pay period's rate."""
    kids = []
    for name, d in sorted(t["funds"].items(), key=lambda x: -sum(x[1]["by"].values())):
        ytd = sum(d["by"].values())
        if abs(ytd) < 0.005: continue
        rate = d["by"].get(t["last"], 0.0) if t["last"] == gmax else 0.0
        kids.append(node(name, None, ytd, ytd + rate * (PERIODS - gmax), src=S_TRK, fb=d["src"],
                         basis="Actual through pay period 27%02d, then the latest pay period's amount for the remaining %d periods" % (gmax, PERIODS - gmax)))
    return kids

# ---------- D. Personnel workbook roster ----------
wb = openpyxl.load_workbook(F_PERS, data_only=True)
rows = list(wb["Rpt Home Dept Roster"].iter_rows(values_only=True))
hi = next(i for i, r in enumerate(rows) if r[0] == "Sal Plan Category")
roster = {}
for r in rows[hi + 1:]:
    if r[0] in ("Faculty", "Administration", "Staff") and r[1]:
        roster[re.sub(r"\s+,", ",", str(r[1]).strip())] = {"cat": r[0], "title": str(r[4] or "").strip(), "fte": r[5], "sal": num(r[14])}

# ---------- assemble people ----------
physicians = sorted(set(model) | set(asf), key=str.lower)
phys_by_last = collections.defaultdict(list)
for p in physicians: phys_by_last[lastkey(p)].append(p)
used_track, used_roster = set(), set()

def find_track(name, grps, first_needed=False):
    lk = lastkey(name); first = name.split(",")[1].strip().lower() if "," in name else ""
    for (g, t), v in track.items():
        if g not in grps or (g, t) in used_track: continue
        tl = lastkey(t); tf = t.split(",")[1].strip().lower() if "," in t else ""
        if tl != lk: continue
        if tf and not first.startswith(tf): continue
        if not tf and first_needed: continue
        used_track.add((g, t)); return v
    return None

def roster_for(name):
    lk = lastkey(name); first = name.split(",")[1].strip().lower()[:3] if "," in name else ""
    for rn, rv in roster.items():
        if lastkey(rn) == lk and rn.split(",")[1].strip().lower().startswith(first):
            used_roster.add(rn); return rv
    return None

ASF_FB = {"Regular ASF pay": "TeamHealth clinical hours and fixed administrative payment",
          "Consult": "TGH toxicology consult funding", "GME": "GME program support",
          "Stipend": "TeamHealth academic stipends", "Shift": "TeamHealth clinical hours (extra shifts)",
          "Bonus": "TeamHealth sign on and retention funding", "RVU": "TeamHealth RVU incentive pass through",
          "CDU": "TeamHealth CDU disposition pass through ($60 per disposition)",
          "PIP": "TeamHealth PIP pass through", "Other special": "TeamHealth incentive pass through"}
def asf_fb(code):
    return next((v for k, v in ASF_FB.items() if code.startswith(k)), "TeamHealth contract, paid out on ASF")

phys_nodes, mal_nodes, tax_nodes, oh_nodes, fixed_nodes = [], [], [], [], []
for p in physicians:
    r = model.get(p); ro = roster_for(p)
    kids = []
    if p in asf or r:
        ak = [node(code, None, v[0], v[1], src=S_ASF, fb=asf_fb(code),
                   basis="Actual for " + months_actual[0] + " to " + months_actual[-1] + ", later months at standard monthly pay")
              for code, v in sorted(asf[p].items(), key=lambda x: -x[1][1])] if p in asf else []
        b = (mc(r, "ASF Salary") + mc(r, "Historical Total incentives") + mc(r, "New Total Incentices Brandon") + mc(r, "TH Sign On Bonus")) if r else None
        kids.append(node("ASF pay (TeamHealth money paid out on ASF)", b, None if ak else None, None if ak else b, ak, src=S_MODEL if r else S_ASF,
                         note="Budget is ASF salary plus incentives and sign on bonus in the FY27 contract model" if r else ""))
    multi = len(phys_by_last[lastkey(p)]) > 1
    t = find_track(p, ("faculty",), multi) or find_track(p, ("research",), multi)
    b = mc(r, "UCS Salary") if r else (ro["sal"] if ro else None)
    if t or b:
        tk = track_nodes(t) if t else []
        kids.append(node("USF (UCS) salary", b, None if tk else None, None if tk else b, tk, src=S_MODEL if r else S_PERS))
    if r:
        fr = mc(r, "UCS Fringe") + mc(r, "ASF Fringe") + (mc(r, "ASF Fringe Expense Incentives") if "ASF Fringe Expense Incentives" in col else 0)
        if fr: kids.append(node("Fringe benefits", fr, None, fr, src=S_MODEL, basis="Budget only. Projection equals budget",
                                kids=[node("UCS fringe", mc(r, "UCS Fringe"), None, mc(r, "UCS Fringe")), node("ASF fringe", mc(r, "ASF Fringe"), None, mc(r, "ASF Fringe")),
                                      node("ASF fringe on incentives", mc(r, "ASF Fringe Expense Incentives"), None, mc(r, "ASF Fringe Expense Incentives"))]))
        for lst, lab in ((mal_nodes, "Mal+CME"), (tax_nodes, "Dean's Tax"), (oh_nodes, "Dept OH")):
            v = mc(r, lab)
            if v: lst.append(node(p, v, None, v, src=S_MODEL, basis="Budget only"))
        v = mc(r, "TeamHealth Total Administrative Cost")
        if v: fixed_nodes.append(node(p, None, None, v, src=S_MODEL, note="Share of the fixed administrative payment in the contract model (stipends, benefits, malpractice, Dean's tax, overhead)"))
    title = (ro["title"] if ro else "") or ("In the FY27 contract model as " + str(r[1]) if r else "")
    if kids: phys_nodes.append(node(p, None, None, None, kids, title=title, pay="Physician. ASF and UCS"))

def other_people(grps, cats, label_pay):
    out = []
    for (g, tname), t in sorted(track.items(), key=lambda x: x[0][1].lower()):
        if g not in grps or (g, tname) in used_track: continue
        cands = [rn for rn in roster if lastkey(rn) == lastkey(tname) and rn not in used_roster]
        ro = roster[cands[0]] if len(cands) == 1 else None
        cat = ro["cat"] if ro else ("Faculty" if g == "faculty" else "Administration")
        if cat not in cats: continue
        used_track.add((g, tname))
        if ro: used_roster.add(cands[0])
        out.append(node(cands[0] if ro else tname, ro["sal"] if ro else None, None, None, track_nodes(t), title=ro["title"] if ro else "", pay=label_pay))
    for rn, ro in sorted(roster.items()):
        if rn in used_roster or ro["cat"] not in cats: continue
        used_roster.add(rn)
        out.append(node(rn, ro["sal"], None, ro["sal"], title=ro["title"], pay=label_pay, src=S_PERS,
                        basis="Budget only. Not found in the FY27 salary tracking workbooks, so the projection equals the March 2026 budget salary"))
    return out

res_fac = other_people(("faculty", "research", "admin"), ("Faculty",), "USF (UCS) accounts only")
staff = other_people(("faculty", "research", "admin"), ("Administration", "Staff"), "USF (UCS) accounts only")

# ---------- operating lines ----------
wb = openpyxl.load_workbook(F_RUN, data_only=True)
ops = []
for r in wb["FY Expenses Summary-Sorted"].iter_rows(min_row=2, values_only=True):
    if not r[0] or str(r[0]).strip().startswith("Subtotal"): break
    v = num(r[1])
    if v and "Toxicology FTE" not in str(r[0]):
        ops.append(node(str(r[0]).strip(), v, None, v, src=S_RUN, note=str(r[4])[:160] if len(r) > 4 and r[4] else "", basis="Budget only"))
allsrc = [("Travel", 161000), ("Contractual services", 782098.40), ("Materials and supplies", 137295), ("OPS (temporary staff)", 106190), ("Other expense including overheads", 1095424)]

expenses = roll(node("All expenses", kids=[
    node("Salaries and payroll", kids=[
        node("Faculty", kids=[
            node("Physicians", kids=phys_nodes, note="Paid on ASF from TeamHealth money and on USF (UCS) accounts"),
            node("Research and other faculty", kids=res_fac, note="PhD faculty can only be paid on USF (UCS) state side accounts, not on ASF")]),
        node("Administration and staff", kids=staff)]),
    node("Operating and other expenses", kids=[
        node("Fellow salary support (GME Schedule A)", 449164, None, 449164, src=S_RUN, basis="Contracted amount"),
        node("Malpractice and CME", kids=mal_nodes),
        node("Dean's tax", kids=tax_nodes),
        node("Department operating lines (UMSA)", kids=ops),
        node("Non payroll spending on USF funds", kids=[node(n, v, None, v, src=S_ALL, basis="Budget only") for n, v in allsrc],
             note="Grants, clinical trials, auxiliaries, carry forward and Foundation, by budget category")])]))

# ---------- pay programs view (same dollars as payroll, grouped by program) ----------
prog = collections.OrderedDict()
def add_prog(group, pname, person, a, p, fb, src):
    g = prog.setdefault(group, collections.OrderedDict())
    d = g.setdefault(pname, {"fb": fb, "src": src, "k": []})
    d["k"].append(node(person, None, a, p))
for pn in phys_nodes + res_fac + staff:
    for c in pn.get("k", []):
        grp = "ASF pay programs (TeamHealth money)" if c["n"].startswith("ASF pay") else ("USF (UCS) salary by fund" if c["n"].startswith("USF") else None)
        if grp:
            for leaf in c.get("k", []): add_prog(grp, leaf["n"], pn["n"], leaf["v"][1], leaf["v"][2], leaf.get("fb", ""), leaf.get("src", ""))
        elif not c.get("k") and pn in res_fac + staff:
            add_prog("USF (UCS) salary by fund", c["n"], pn["n"], c["v"][1], c["v"][2], c.get("fb", ""), c.get("src", ""))
programs = roll(node("All pay programs", kids=[
    node(g, kids=[node(pname, kids=sorted(d["k"], key=lambda x: -(x["v"][2] or 0)), fb=d["fb"], src=d["src"]) for pname, d in ps.items()])
    for g, ps in prog.items()]))
for g in programs["k"]: g["k"].sort(key=lambda x: -(x["v"][2] or 0))

# ---------- revenue tree ----------
rev = list(csv.DictReader(open(REV_CSV)))
n_fixed_paid = len(months_actual)  # paid on the 15th of each month
inc = collections.defaultdict(lambda: [0.0, 0.0])
for p, codes in asf.items():
    for code, v in codes.items():
        if code.startswith(("RVU", "CDU", "PIP", "Other special")):
            inc[code][0] += v[0]; inc[code][1] += v[1]
streams = collections.OrderedDict()
for r in rev:
    f26, f27 = float(r["fy26"] or 0), float(r["fy27"] or 0)
    kids, a = None, None
    if r["line"].startswith("TeamHealth fixed"):
        a = f27 / 12 * n_fixed_paid; kids = sorted(fixed_nodes, key=lambda x: -x["v"][2])
        for k in kids: k["v"] = [None, None, k["v"][2]]
    elif r["line"].startswith("TeamHealth clinical"):
        inv = [("July 2026 invoice", 568303.08), ("August 2026 invoice", 577407.0), ("September 2026 invoice", 572538.60)]
        a = sum(v for _, v in inv)
        kids = [node(n, None, v, v, src="Box, Contracts/TeamHealth Contract/TH Invoices.xlsx") for n, v in inv]
        kids.append(node("October 2026 to June 2027 at the three month average", None, None, f27 - a, basis="Estimate"))
    elif r["line"].startswith("TeamHealth incentive"):
        kids = [node(code, None, v[0], None, src=S_ASF, note="Paid to physicians on ASF to date. Revenue passes through UMSA") for code, v in sorted(inc.items(), key=lambda x: -x[1][0])]
        a = sum(v[0] for v in inc.values())
    n = node(r["line"], f26, a, f27, kids, fund=r["fund"], basis=r["basis"])
    if kids: n["v"] = [f26, a, f27]
    streams.setdefault(r["stream"], []).append(n)
revenue = node("All revenue", kids=[node(s, kids=ls) for s, ls in streams.items()])
def roll_rev(n):
    kids = n.get("k", [])
    if all(x is None for x in n["v"]):
        for c in kids: roll_rev(c)
        for i in range(3):
            vals = [c["v"][i] for c in kids if c["v"][i] is not None]
            n["v"][i] = round(sum(vals), 2) if vals else None
    else:
        n["v"] = [None if x is None else round(x, 2) for x in n["v"]]
        for c in kids:
            c["v"] = [None if x is None else round(x, 2) for x in c["v"]]
    return n
roll_rev(revenue)

umsa = [{"month": r["month"], "revenue": float(r["revenue"]), "expenses": float(r["expenses"]), "balance": float(r["balance"]), "basis": r["clinical_basis"]}
        for r in csv.DictReader(open(UMSA_CSV))]
out = {"generated": datetime.date.today().isoformat(),
       "asf_actual_through": months_actual[-1], "ucs_actual_through": "pay period 27%02d" % gmax,
       "tabs": {"revenue": {"cols": ["FY26", "FY27 actual to date", "FY27 projected"], "tree": revenue},
                "expenses": {"cols": ["FY27 budget", "FY27 actual to date", "FY27 projected"], "tree": expenses},
                "programs": {"cols": ["FY27 budget", "FY27 actual to date", "FY27 projected"], "tree": programs}},
       "umsa": umsa}
json.dump(out, open(OUT, "w"))
def show(n, d=0, maxd=2):
    print("  " * d + n["n"][:50].ljust(52 - 2 * d), [None if x is None else round(x) for x in n["v"]], len(n.get("k", [])))
    if d < maxd:
        for c in n.get("k", []): show(c, d + 1, maxd)
show(revenue, 0, 1); show(expenses, 0, 3 if False else 2); show(programs, 0, 2)
print("physicians", len(phys_nodes), "research/other faculty", len(res_fac), "staff", len(staff), "unused tracking", [k for k in track if k not in used_track])
