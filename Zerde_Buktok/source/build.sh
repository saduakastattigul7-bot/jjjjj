#!/bin/sh
# Толық құрастыру: ./build.sh  (Zertteu_kestesi.xlsx толтырылса, нәтижелер өздігінен енеді)
set -e
cd "$(dirname "$0")"
OUT=..
[ -f "$OUT/Zertteu_kestesi.xlsx" ] || python3 make_xlsx.py "$OUT/Zertteu_kestesi.xlsx"
rm -f natije.json; python3 stats.py "$OUT/Zertteu_kestesi.xlsx" natije.json
python3 prep.py && python3 figs.py
rm -f toc_pages.json
node build.js out.docx && soffice --headless --convert-to pdf out.docx >/dev/null 2>&1
python3 pages.py out.pdf
for V in "" "--jury"; do
  NAME=Zertteu_zhumysy; [ -n "$V" ] && NAME=Zertteu_zhumysy_qazylar_danasy
  node build.js "$OUT/$NAME.docx" $V
  python3 post.py "$OUT/$NAME.docx"
  soffice --headless --convert-to pdf --outdir "$OUT" "$OUT/$NAME.docx" >/dev/null 2>&1
done
rm -f out.docx out.pdf
if [ -f docs.js ]; then
  node docs.js "$OUT"
  for N in Zertteu_kundeligi ZhI_deklaratsiyasy Etika_paragy Zhetekshi_pikiri; do
    python3 post.py "$OUT/$N.docx"
    soffice --headless --convert-to pdf --outdir "$OUT" "$OUT/$N.docx" >/dev/null 2>&1
  done
fi
