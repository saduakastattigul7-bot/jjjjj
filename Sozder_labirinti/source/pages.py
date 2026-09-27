# Мазмұнындағы бет нөмірлерін PDF-тен анықтау (екі өтімді құрастыру)
import json, sys, re, subprocess
from pypdf import PdfReader
pdf = sys.argv[1]
toc = json.loads(subprocess.check_output(["node", "-e", "console.log(JSON.stringify(require('./paper').TOC.map(x=>x[0])))"]))
pages = [re.sub(r"\s+", " ", p.extract_text() or "") for p in PdfReader(pdf).pages]
start = next(i for i, t in enumerate(pages) if "МАЗМҰНЫ" in t) + 1
res = {}
for h in toc:
    for i in range(start, len(pages)):
        if h in pages[i]:
            res[h] = i + 1; start = i; break
json.dump(res, open("pages.json", "w"), ensure_ascii=False, indent=1)
print(res)
print("Негізгі мәтін (кіріспеден дереккөздерге дейін):", res["ПАЙДАЛАНЫЛҒАН ДЕРЕККӨЗДЕР"] - res["КІРІСПЕ"], "бет; барлығы", len(pages))
