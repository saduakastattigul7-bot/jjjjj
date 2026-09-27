#!/usr/bin/env bash
# Барлық құжаттарды қайта жинау: bash src/build.sh   (QR үшін: STAND_URL=https://... bash src/build.sh)
set -euo pipefail
cd "$(dirname "$0")"
ROOT=..
TMP=$(mktemp -d)
node figs.js
# Негізгі жұмыс: мазмұнындағы бет нөмірлері үшін екі рет жинаймыз
rm -f toc_pages.json
node zhoba.js "$TMP/z.docx"
(cd "$TMP" && soffice --headless --convert-to pdf z.docx >/dev/null 2>&1)
HEADS=$(node -e 'const s=require("fs").readFileSync("zhoba.js","utf8");console.log(JSON.stringify(eval(s.match(/const TOC = (\[[\s\S]*?\]);/)[1]).map(x=>x[0])))')
python3 toc.py "$TMP/z.pdf" "$HEADS"
node zhoba.js "$ROOT/Sozder_labirinti_gylymi_zhoba.docx"
node qujattar.js "$ROOT"
node deck.js "$ROOT/Sozder_labirinti_prezentatsiya.pptx"
node stend.js "$ROOT/Stend_maketi.pdf"
# Ережелердің 2.10-т.: әр құжат жеке PDF
cd "$ROOT"
for f in Sozder_labirinti_gylymi_zhoba Annotatsiya Zertteu_kundeligi ZhI_deklaratsiya Avtorlyq_ules_paragy Antiplagiat_anyqtama Etika_qauipsizdik_paragy Qorgau_sozi; do
  soffice --headless --convert-to pdf --outdir . "$f.docx" >/dev/null 2>&1
done
rm -rf "$TMP"
ls -la *.docx *.pdf *.pptx
