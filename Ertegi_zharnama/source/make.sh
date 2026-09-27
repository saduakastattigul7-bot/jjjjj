#!/bin/sh
# Builds every deliverable into the parent folder.
# Needs: node (docx), python3 (matplotlib, openpyxl, pymupdf), LibreOffice, Playwright Chromium.
set -e
cd "$(dirname "$0")"
OUT=..
TMP=$(mktemp -d)
python3 figs.py
# Main work: two-pass build so the TOC gets real page numbers
for mode in org jury; do
  flag=""; [ "$mode" = jury ] && flag="--jury"
  rm -f toc_pages.json
  for pass in 1 2; do
    node build.js "$TMP/zhoba_$mode.docx" $flag
    python3 post.py "$TMP/zhoba_$mode.docx"
    soffice --headless --convert-to pdf --outdir "$TMP" "$TMP/zhoba_$mode.docx" >/dev/null 2>&1
    python3 pages.py "$TMP/zhoba_$mode.pdf" > /dev/null
  done
done
rm -f toc_pages.json
cp "$TMP/zhoba_org.docx" "$OUT/01_Zhoba_uiymdastyrushyga.docx"
cp "$TMP/zhoba_org.pdf"  "$OUT/01_Zhoba_uiymdastyrushyga.pdf"
cp "$TMP/zhoba_jury.docx" "$OUT/01_Zhoba_qazylarga.docx"
cp "$TMP/zhoba_jury.pdf"  "$OUT/01_Zhoba_qazylarga.pdf"
# Forms (appendices 3-6 of the rules) and supervisor review
node forms.js "$OUT"
node forms.js "$OUT" --jury
for f in "$OUT"/0[2-6]_*.docx; do
  python3 post.py "$f"
  soffice --headless --convert-to pdf --outdir "$OUT" "$f" >/dev/null 2>&1
done
rm -f "$OUT"/*_qazylarga.docx.tmp
rm -f "$OUT"/0[23]_*_qazylarga.docx
# Stand layout and data workbook
node stend.js "$OUT"
python3 xlsx.py "$OUT/08_Derekter.xlsx"
rm -rf "$TMP"
ls -1 "$OUT"
