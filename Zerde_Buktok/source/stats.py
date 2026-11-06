"""Толтырылған Zertteu_kestesi.xlsx файлынан нәтижені есептеп, natije.json жазады (формулаларға сүйенбейді).
Ешбір дерек болмаса, natije.json жазылмайды."""
import json, sys
from openpyxl import load_workbook
from zertteu import SURVEY, SOURCES, TEXT_Q, N_RESP, N_BLOGGERS, N_INTERVIEWS

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
# «Қайдан білесің?» – тек кемінде бір нұсқа белгілеген оқушылар бойынша
src_base = [a for a in resp if any(a["SRC"])]
R["n_src"] = len(src_base)
R["src"] = {sname: pct(sum(a["SRC"][j] for a in src_base), len(src_base)) for j, sname in enumerate(SOURCES)}
R["src_n"] = {sname: sum(a["SRC"][j] for a in src_base) for j, sname in enumerate(SOURCES)}
R["by_class"] = {c: sum(a["SYN"] == c for a in resp) for c in ("7", "8", "9")}
R["why"] = [a["WHY"] for a in resp if a["WHY"]][:40]
R["blog_names"] = [a["BLOG"] for a in resp if a["BLOG"]][:60]

# ---- Блогерлер, сұхбат ----
wbl = wb["Блогерлер"]
def one(v):  # көп жолды мәтінді бір жолға: «Бұл арнада сіз:\n– А\n– Б» → «А; Б»
    ls = [" ".join(x.strip(" –-•").split()) for x in s(v).split("\n")]
    ls = [x for x in ls if x and not x.endswith(":")]
    return "; ".join(ls)
R["bloggers"] = [{"name": one(wbl.cell(r, 2).value), "net": s(wbl.cell(r, 3).value), "subs": num(wbl.cell(r, 4).value),
                  "topic": one(wbl.cell(r, 5).value), "kaz_lit": s(wbl.cell(r, 6).value)}
                 for r in range(2, N_BLOGGERS + 2) if s(wbl.cell(r, 2).value)]
wi = wb["Сұхбат"]
def unq(v):  # сыртқы тырнақшаны алып, ішкі "…" тырнақшаны „…“ етеді
    q = s(v).strip().rstrip(".").strip("«»\"“” ")
    while q.count('"') >= 2:
        q = q.replace('"', "„", 1).replace('"', "“", 1)
    return q
R["interviews"] = [{"code": s(wi.cell(r, 1).value), "who": s(wi.cell(r, 2).value), "idea": s(wi.cell(r, 3).value), "quote": unq(wi.cell(r, 4).value)}
                   for r in range(2, N_INTERVIEWS + 2) if s(wi.cell(r, 3).value) or s(wi.cell(r, 4).value)]

if not (resp or R["bloggers"] or R["interviews"]):
    print("Кестеде дерек жоқ – natije.json жазылмады.")
else:
    json.dump(R, open(out, "w", encoding="utf8"), ensure_ascii=False, indent=1)
    print("written", out, "resp:", len(resp),
          "bloggers:", len(R["bloggers"]), "interviews:", len(R["interviews"]))
