"""1-өтуден кейін PDF-тен тараулардың бет нөмірін тауып, toc_pages.json жазады."""
import pymupdf, json, re, sys
d = pymupdf.open(sys.argv[1] if len(sys.argv) > 1 else "out.pdf")
heads = ["АННОТАЦИЯ"] + [re.sub(r"^#\*? |^## ", "", l) for l in open("content.txt", encoding="utf8").read().split("\n") if re.match(r"^#\*? |^## ", l)]
heads += [f"{a} ҚОСЫМШАСЫ" for a in "АӘБВ"]
norm = lambda s: re.sub(r"\s+", " ", s)
res, start = {}, 1
for h in heads:
    key = norm(h)[:40]
    for i in range(start, d.page_count):
        txt = norm(d[i].get_text())
        if key in txt and (h != "АННОТАЦИЯ" or i < 3) and "МАЗМҰНЫ" not in txt:
            res[h] = i + 1; start = i; break
    else:
        print("NOT FOUND", h)
json.dump(res, open("toc_pages.json", "w"), ensure_ascii=False)
print("pages", d.page_count, res)
