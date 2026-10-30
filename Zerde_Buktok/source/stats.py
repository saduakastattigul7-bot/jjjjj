"""Толтырылған Zertteu_kestesi.xlsx файлынан нәтижені есептеп, natije.json жазады (формулаларға сүйенбейді).
Ешбір дерек болмаса, natije.json жазылмайды."""
import json, sys
from openpyxl import load_workbook
from zertteu import SURVEY, SOURCES, TEXT_Q, N_RESP, WEEKS, N_TRAILERS, N_BLOGGERS, N_INTERVIEWS

src = sys.argv[1] if len(sys.argv) > 1 else "../Zertteu_kestesi.xlsx"
out = sys.argv[2] if len(sys.argv) > 2 else "natije.json"
wb = load_workbook(src, data_only=True)
s = lambda v: str(v).strip() if v not in (None, "") else ""


def num(v):
    try:
        return int(float(str(v).replace(",", ".")))
    except (TypeError, ValueError):
        return None


pct = lambda a, b: round(a / b * 100, 1) if b else None
R = {}

# ---- Сауалнама ----
wsv = wb["Сауалнама"]
resp = []
for r in range(2, N_RESP + 2):
    vals = [wsv.cell(r, c).value for c in range(2, 2 + len(SURVEY) + len(SOURCES) + len(TEXT_Q))]
    if all(v in (None, "") for v in vals):
        continue
    a = {}
    for k, (code, _, _) in enumerate(SURVEY):
        a[code] = s(wsv.cell(r, 2 + k if k < 5 else 2 + len(SOURCES) + k).value)
    a["SRC"] = [num(wsv.cell(r, 7 + j).value) == 1 for j in range(len(SOURCES))]
    for t, (code, _) in enumerate(TEXT_Q):
        a[code] = s(wsv.cell(r, 7 + len(SOURCES) + len(SURVEY) - 5 + t).value)
    resp.append(a)
R["n_resp"] = len(resp)
R["survey"] = {}
for code, q, opts in SURVEY:
    base = [a[code] for a in resp if a[code]]
    if code == "FINISH":  # тек блогер ұсынған кітапты оқығандар
        base = [a[code] for a in resp if a[code] and a["TRIED"] == "Иә"]
    R["survey"][code] = {"n": len(base), **{o: pct(sum(x == o for x in base), len(base)) for o in opts}}
R["src"] = {sname: pct(sum(a["SRC"][j] for a in resp), len(resp)) for j, sname in enumerate(SOURCES)}
R["by_class"] = {c: sum(a["SYN"] == c for a in resp) for c in ("7", "8", "9")}
R["why"] = [a["WHY"] for a in resp if a["WHY"]][:40]
R["blog_names"] = [a["BLOG"] for a in resp if a["BLOG"]][:60]

# ---- Тәжірибе ----
wt = wb["Тәжірибе"]
R["exp_name"], R["exp_size"] = s(wt.cell(1, 3).value), num(wt.cell(1, 5).value)
R["ctrl_name"], R["ctrl_size"] = s(wt.cell(2, 3).value), num(wt.cell(2, 5).value)
weeks = []
for i, (w, ph) in enumerate(WEEKS):
    r = 5 + i
    weeks.append({"week": w, "phase": ph, **{k: num(wt.cell(r, c).value) for k, c in
                  (("e_take", 3), ("c_take", 4), ("e_fin", 5), ("c_fin", 6), ("e_rev", 7), ("c_rev", 8))}})
R["weeks"] = weeks
R["exp_filled"] = sum(any(w[k] is not None for k in ("e_take", "c_take")) for w in weeks)
tot = lambda k, ph: sum(w[k] or 0 for w in weeks if w["phase"] == ph)
R["exp"] = {k: {"before": tot(k, "Дейін"), "during": tot(k, "Кезінде")} for k in ("e_take", "c_take", "e_fin", "c_fin", "e_rev", "c_rev")}
g = lambda k: round(R["exp"][k]["during"] / R["exp"][k]["before"], 1) if R["exp"][k]["before"] else None
R["growth_e"], R["growth_c"] = g("e_take"), g("c_take")

# ---- Буктрейлерлер, блогерлер, сұхбат ----
wtr = wb["Буктрейлерлер"]
R["trailers"] = [{"book": s(wtr.cell(r, 2).value), "author": s(wtr.cell(r, 3).value), "date": s(wtr.cell(r, 4).value)[:10],
                  "sec": num(wtr.cell(r, 5).value), "asked": num(wtr.cell(r, 6).value)}
                 for r in range(2, N_TRAILERS + 2) if s(wtr.cell(r, 2).value)]
wbl = wb["Блогерлер"]
R["bloggers"] = [{"name": s(wbl.cell(r, 2).value), "net": s(wbl.cell(r, 3).value), "subs": s(wbl.cell(r, 4).value),
                  "topic": s(wbl.cell(r, 5).value), "kaz_lit": s(wbl.cell(r, 6).value)}
                 for r in range(2, N_BLOGGERS + 2) if s(wbl.cell(r, 2).value)]
wi = wb["Сұхбат"]
R["interviews"] = [{"code": s(wi.cell(r, 1).value), "who": s(wi.cell(r, 2).value), "idea": s(wi.cell(r, 3).value), "quote": s(wi.cell(r, 4).value)}
                   for r in range(2, N_INTERVIEWS + 2) if s(wi.cell(r, 3).value) or s(wi.cell(r, 4).value)]

if not (resp or R["exp_filled"] or R["trailers"] or R["bloggers"] or R["interviews"]):
    print("Кестеде дерек жоқ – natije.json жазылмады.")
else:
    json.dump(R, open(out, "w", encoding="utf8"), ensure_ascii=False, indent=1)
    print("written", out, "resp:", len(resp), "weeks:", R["exp_filled"], "trailers:", len(R["trailers"]),
          "bloggers:", len(R["bloggers"]), "interviews:", len(R["interviews"]))
