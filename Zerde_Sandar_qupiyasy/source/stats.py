"""Толтырылған Zertteu_kestesi.xlsx файлынан нәтижені есептеп, natije.json жазады.
«Сандар» парағында дерек жоқ болса, natije.json жазылмайды."""
import json, sys
from collections import Counter
from openpyxl import load_workbook
from sandar import CORPUS, TALE_TYPES, CATS, NUMBERS, SACRED, N_ROWS

src = sys.argv[1] if len(sys.argv) > 1 else "../Zertteu_kestesi.xlsx"
out = sys.argv[2] if len(sys.argv) > 2 else "natije.json"
wb = load_workbook(src, data_only=True)
s = lambda v: str(v).strip() if v not in (None, "") else ""
TYPE = dict(CORPUS)

R = {}
wc = wb["Ертегілер"]
R["tales"] = [{"name": wc.cell(r, 2).value, "type": wc.cell(r, 3).value, "edition": s(wc.cell(r, 4).value),
               "read": s(wc.cell(r, 5).value).lower() == "иә", "triad": s(wc.cell(r, 6).value).lower() == "иә",
               "triad_ex": s(wc.cell(r, 7).value)} for r in range(2, len(CORPUS) + 2)]
read = [t for t in R["tales"] if t["read"]] or R["tales"]
R["tales_read"] = sum(t["read"] for t in R["tales"])
R["triad"] = sum(t["triad"] for t in R["tales"])
R["triad_magic"] = sum(t["triad"] for t in R["tales"] if t["type"] == "Қиял-ғажайып")
R["magic_read"] = sum(1 for t in read if t["type"] == "Қиял-ғажайып")

wn = wb["Сандар"]
rows = []
for r in range(2, N_ROWS + 2):
    tale, num = s(wn.cell(r, 2).value), wn.cell(r, 4).value
    try:
        num = int(num)
    except (TypeError, ValueError):
        continue
    if not tale:
        continue
    rows.append({"tale": tale, "type": TYPE.get(tale, ""), "num": num, "phrase": s(wn.cell(r, 5).value),
                 "cat": s(wn.cell(r, 6).value), "func": s(wn.cell(r, 7).value).upper(), "page": s(wn.cell(r, 8).value)})
R["rows"] = rows
R["total"] = len(rows)
cnt = Counter(x["num"] for x in rows)
R["by_num"] = {str(n): {"all": cnt.get(n, 0), "Н": sum(x["num"] == n and x["func"] == "Н" for x in rows),
                        "Ф": sum(x["num"] == n and x["func"] == "Ф" for x in rows)} for n in NUMBERS}
R["other"] = sum(v for k, v in cnt.items() if k not in NUMBERS)
R["other_list"] = sorted(k for k in cnt if k not in NUMBERS)
R["distinct"] = len(cnt)
R["top"] = cnt.most_common(5)
R["sacred"] = sum(cnt.get(n, 0) for n in SACRED)
R["sacred_pct"] = round(R["sacred"] / len(rows) * 100, 1) if rows else None
R["F"] = sum(x["func"] == "Ф" for x in rows)
R["F_pct"] = round(R["F"] / len(rows) * 100, 1) if rows else None
sac = [x for x in rows if x["num"] in SACRED]
R["sacred_F_pct"] = round(sum(x["func"] == "Ф" for x in sac) / len(sac) * 100, 1) if sac else None
non = [x for x in rows if x["num"] not in SACRED]
R["nonsacred_F_pct"] = round(sum(x["func"] == "Ф" for x in non) / len(non) * 100, 1) if non else None
R["by_cat"] = {c: sum(x["cat"] == c for x in rows) for c in CATS}
R["by_type"] = {t: {"n": sum(x["type"] == t for x in rows), "F": sum(x["type"] == t and x["func"] == "Ф" for x in rows),
                    "tales": sum(1 for x in read if x["type"] == t)} for t in TALE_TYPES}
R["per_tale"] = round(len(rows) / max(1, len(read)), 1)
R["by_tale"] = dict(Counter(x["tale"] for x in rows))
# сан мен сөз тіркесінің жиі қайталанатын жұптары (формулаларды табу үшін)
R["phrases"] = Counter(x["phrase"].lower() for x in rows if x["phrase"]).most_common(10)

if not rows:
    print("«Сандар» парағында дерек жоқ – natije.json жазылмады.")
else:
    json.dump(R, open(out, "w", encoding="utf8"), ensure_ascii=False, indent=1)
    print("written", out, "numbers:", len(rows))
