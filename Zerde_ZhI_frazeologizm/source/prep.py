"""build.js үшін барлық деректерді бір data.json файлына жинайды."""
import json, os
import tazhiribe as T
from taldau import res as taldau

sozdik = [l.split(" – ", 1) for l in open("sozdik.txt", encoding="utf8").read().split("\n") if l.strip()]
toptar = []
for blk in open("toptar.txt", encoding="utf8").read().split("# ")[1:]:
    head, items = blk.strip().split("\n", 1)
    name, gloss = head.split(" | ")
    toptar.append({"name": name, "gloss": gloss, "items": [x.strip() for x in items.split(";") if x.strip()]})

ALPH = "аәбвгғдеёжзийкқлмнңоөпрстуұүфхһцчшщъыіьэюя"
kaz = lambda w: [ALPH.index(c) if c in ALPH else (-1 if c in " -" else 100 + ord(c)) for c in w.lower()]

natije = json.load(open("natije.json", encoding="utf8")) if os.path.exists("natije.json") else {}
data = {
    "groups": T.GROUPS, "e1": T.E1, "e2": T.E2, "e3": T.E3, "e3_n": T.E3_N,
    "score": T.SCORE, "err": T.ERR, "e3_codes": T.E3_CODES,
    "prompts": [T.PROMPT1, T.PROMPT2, T.PROMPT3],
    "sozdik": sorted(sozdik, key=lambda p: kaz(p[0])), "toptar": toptar, "taldau": taldau, "natije": natije,
}
json.dump(data, open("data.json", "w", encoding="utf8"), ensure_ascii=False, indent=1)
print("data.json: sozdik", len(sozdik), "toptar", len(toptar), "natije", "yes" if natije else "no")
