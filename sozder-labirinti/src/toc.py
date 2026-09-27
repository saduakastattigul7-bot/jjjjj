# PDF-тен тараулардың бет нөмірін табады (мазмұнына арналған)
import pypdfium2 as pdfium, json, re, sys
pdf = pdfium.PdfDocument(sys.argv[1])
heads = json.loads(sys.argv[2])
norm = lambda s: re.sub(r"\s+", " ", s).strip()
texts = [norm(pdf[i].get_textpage().get_text_range()) for i in range(len(pdf))]
res = {}; start = 2
for h in heads:
    key = norm(h)
    for i in range(start, len(texts)):
        if key in texts[i] or key.upper() in texts[i]:
            res[h] = i + 1; start = i; break
    else: print("NOT FOUND", h)
print(len(texts), "pages"); print(res)
json.dump(res, open("toc_pages.json", "w"), ensure_ascii=False)
