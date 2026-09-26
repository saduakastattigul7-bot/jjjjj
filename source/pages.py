import pymupdf as fitz, json, re
d = fitz.open("out.pdf"); print("pages", d.page_count)
heads = [re.sub(r"^#\*? |^## ", "", l) for l in open("content.txt", encoding="utf8").read().split("\n") if re.match(r"^#\*? |^## ", l)]
res = {}; start = 3
norm = lambda s: re.sub(r"\s+", " ", s)
for h in heads:
    key = norm(h)[:35]
    for i in range(start, d.page_count):
        if key in norm(d[i].get_text()):
            res[h] = i + 1; start = i; break
    else: print("NOT FOUND", h)
for k,v in res.items(): print(v, k)
json.dump(res, open("toc_pages.json", "w"), ensure_ascii=False)
