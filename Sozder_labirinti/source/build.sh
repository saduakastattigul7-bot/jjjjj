#!/bin/bash
# Барлық құжатты құрастыру: DOCX → PDF (LibreOffice), мазмұнындағы беттер екі өтіммен анықталады
set -e
cd "$(dirname "$0")"
rm -f pages.json
node paper.js
soffice --headless --convert-to pdf --outdir .. ../01_Zhoba_zhumysy.docx >/dev/null 2>&1
python3 pages.py ../01_Zhoba_zhumysy.pdf
node paper.js && node paper.js --jury && node docs.js
cd ..
for f in 0*.docx 1*.docx; do soffice --headless --convert-to pdf "$f" >/dev/null 2>&1; done
ls -la *.pdf
