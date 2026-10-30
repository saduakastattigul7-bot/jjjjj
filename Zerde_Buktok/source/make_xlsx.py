"""Оқушы толтыратын зерттеу кестесі: Zertteu_kestesi.xlsx
python3 make_xlsx.py шығыс.xlsx [толтырылған_ескі.xlsx] [google_forms_жауаптары.xlsx]"""
import json, os, sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from zertteu import (SURVEY, SRC_Q, SOURCES, TEXT_Q, N_RESP, WEEKS, EXP_DEFAULT, CTRL_DEFAULT,
                     N_TRAILERS, N_BLOGGERS, N_INTERVIEWS, NETS)

OUT = sys.argv[1] if len(sys.argv) > 1 else "Zertteu_kestesi.xlsx"
thin = Side(style="thin", color="999999")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD = PatternFill("solid", fgColor="DCE6F1")
FILL = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
BOLD = Font(bold=True)


def header(ws, row, titles, widths, height=48):
    for j, (t, w) in enumerate(zip(titles, widths), 1):
        c = ws.cell(row, j, t)
        c.fill, c.font, c.alignment, c.border = HEAD, BOLD, CEN, BOX
        if w:
            ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[row].height = height


def put(ws, r, j, v=None, fill=False, center=False):
    c = ws.cell(r, j, v)
    c.border = BOX
    c.alignment = CEN if center else WRAP
    if fill:
        c.fill = FILL
    return c


def dv_list(ws, items):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True, showErrorMessage=True,
                        errorTitle="Қате мән", error="Тізімнен таңдаңыз: " + ", ".join(items))
    ws.add_data_validation(dv)
    return dv


def dv_int(ws, lo=0, hi=500):
    dv = DataValidation(type="whole", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True,
                        showErrorMessage=True, errorTitle="Қате мән", error=f"{lo}-ден {hi}-ге дейінгі бүтін сан жазыңыз")
    ws.add_data_validation(dv)
    return dv


wb = Workbook()
ws = wb.active
ws.title = "Нұсқау"
ws.column_dimensions["A"].width = 120
lines = [
    ("ЗЕРТТЕУ КЕСТЕСІН ТОЛТЫРУ НҰСҚАУЫ", True),
    ("Сары ұяшықтарды ғана толтырыңыз. «Қорытынды» парағы өздігінен есептеледі. Парақтарды, бағандарды өшірмеңіз.", False),
    ("", False),
    ("1. Сауалнама («Сауалнама» парағы)", True),
    ("Google Forms жауаптарын Файл → Жүктеп алу → .xlsx арқылы жүктеп жіберсеңіз, кестеге өздігінен аударылады. Қолмен толтырсаңыз: бір оқушы – бір жол, есім жазылмайды.", False),
    ("«Қайдан білесің?» сұрағында оқушы таңдаған әр нұсқаға 1 қойыңыз, таңдамағанына 0.", False),
    ("", False),
    ("2. «Буктрейлер» тәжірибесі («Тәжірибе» парағы)", True),
    ("Кітапханашыдан формулярлар бойынша әр аптада сыныптың кітапханадан алған кітап санын, соңына дейін оқып қайтарған кітап санын сұраңыз.", False),
    ("«Дейін» – тәжірибеге дейінгі 4 апта (салыстыру үшін), «Кезінде» – буктрейлер көрсетілген 4 апта. Екі сынып та бірдей есептеледі.", False),
    ("«Пікір / бейне» – сол аптада оқыған кітабы туралы пікір жазған немесе бейне түсірген оқушылар саны.", False),
    ("", False),
    ("3. Буктрейлерлер, блогерлер, сұхбат", True),
    ("«Буктрейлерлер» – өзіңіз түсірген әр бейне: кітап, автор, көрсетілген күні, ұзақтығы және сол кітапты эксперименттік сыныптан неше оқушы сұрағаны.", False),
    ("«Блогерлер» – TikTok, Instagram, YouTube-тен тапқан қазақ тіліндегі кітап блогерлері (кемінде 10).", False),
    ("«Сұхбат» – кітапханашы мен оқушылардың негізгі ойы; есім жазылмайды, тек код (К01, О01 …).", False),
]
for i, (t, b) in enumerate(lines, 1):
    c = ws.cell(i, 1, t)
    c.alignment = WRAP
    c.font = Font(bold=b, size=12 if b else 11)

# ---- Сауалнама ----
wsv = wb.create_sheet("Сауалнама")
cols = ["Код"] + [q for _, q, _ in SURVEY[:5]] + [f"Қайдан: {s}" for s in SOURCES] + [q for _, q, _ in SURVEY[5:]] + [q for _, q in TEXT_Q]
header(wsv, 1, cols, [7] + [14] * 5 + [10] * len(SOURCES) + [14] * (len(SURVEY) - 5) + [30, 30], height=90)
wsv.freeze_panes = "B2"
dvs = [dv_list(wsv, opts) for _, _, opts in SURVEY]
dv01 = dv_list(wsv, ["0", "1"])
for i in range(1, N_RESP + 1):
    r = i + 1
    put(wsv, r, 1, f"С{i:03d}", center=True)
    c = 2
    for k in range(5):
        dvs[k].add(put(wsv, r, c, fill=True, center=True)); c += 1
    for _ in SOURCES:
        dv01.add(put(wsv, r, c, fill=True, center=True)); c += 1
    for k in range(5, len(SURVEY)):
        dvs[k].add(put(wsv, r, c, fill=True, center=True)); c += 1
    for _ in TEXT_Q:
        put(wsv, r, c, fill=True); c += 1
SV_LAST = N_RESP + 1
COL = {code: get_column_letter(2 + k if k < 5 else 2 + len(SOURCES) + k) for k, (code, _, _) in enumerate(SURVEY)}
SRC_COL = [get_column_letter(7 + j) for j in range(len(SOURCES))]

# ---- Тәжірибе ----
wt = wb.create_sheet("Тәжірибе")
for col, w in zip("ABCDEFGH", [10, 12, 14, 14, 14, 14, 14, 14]):
    wt.column_dimensions[col].width = w
put(wt, 1, 1, "Эксперименттік сынып").font = BOLD; put(wt, 1, 3, EXP_DEFAULT[0], fill=True, center=True); put(wt, 1, 4, "оқушы саны"); put(wt, 1, 5, EXP_DEFAULT[1], fill=True, center=True)
put(wt, 2, 1, "Бақылау сыныбы").font = BOLD; put(wt, 2, 3, CTRL_DEFAULT[0], fill=True, center=True); put(wt, 2, 4, "оқушы саны"); put(wt, 2, 5, CTRL_DEFAULT[1], fill=True, center=True)
header(wt, 4, ["Апта", "Кезең", "Эксп.: алынған кітап", "Бақ.: алынған кітап", "Эксп.: соңына дейін оқылған", "Бақ.: соңына дейін оқылған", "Эксп.: пікір / бейне", "Бақ.: пікір / бейне"], [None] * 8, height=60)
dvn = dv_int(wt)
for i, (w, ph) in enumerate(WEEKS):
    r = 5 + i
    put(wt, r, 1, w, center=True); put(wt, r, 2, ph, center=True)
    for j in range(3, 9):
        dvn.add(put(wt, r, j, fill=True, center=True))
EX_FIRST, EX_LAST = 5, 4 + len(WEEKS)

# ---- Буктрейлерлер ----
wtr = wb.create_sheet("Буктрейлерлер")
header(wtr, 1, ["№", "Кітап", "Автор", "Көрсетілген күні", "Ұзақтығы, секунд", "Эксп. сыныптан сұрағандар саны"], [5, 30, 24, 14, 12, 16])
dvs2 = dv_int(wtr, 0, 600)
for i in range(1, N_TRAILERS + 1):
    put(wtr, i + 1, 1, i, center=True)
    for j in (2, 3, 4):
        put(wtr, i + 1, j, fill=True)
    dvs2.add(put(wtr, i + 1, 5, fill=True, center=True)); dvs2.add(put(wtr, i + 1, 6, fill=True, center=True))

# ---- Блогерлер ----
wbl = wb.create_sheet("Блогерлер")
header(wbl, 1, ["№", "Блогер (парақша)", "Желі", "Жазылушы саны", "Негізгі тақырып / жанр", "Ұлттық әдебиет бар ма?"], [5, 28, 14, 14, 40, 14])
dvnet, dvyn = dv_list(wbl, NETS), dv_list(wbl, ["Иә", "Жоқ"])
for i in range(1, N_BLOGGERS + 1):
    put(wbl, i + 1, 1, i, center=True); put(wbl, i + 1, 2, fill=True)
    dvnet.add(put(wbl, i + 1, 3, fill=True, center=True)); put(wbl, i + 1, 4, fill=True, center=True)
    put(wbl, i + 1, 5, fill=True); dvyn.add(put(wbl, i + 1, 6, fill=True, center=True))

# ---- Сұхбат ----
wi = wb.create_sheet("Сұхбат")
header(wi, 1, ["Код", "Кім", "Негізгі ойы (қысқаша)", "Дәйексөз (өз сөзі)"], [7, 16, 50, 50])
dvwho = dv_list(wi, ["Кітапханашы", "Оқушы", "Мұғалім"])
for i in range(1, N_INTERVIEWS + 1):
    put(wi, i + 1, 1, "К01" if i == 1 else f"О{i - 1:02d}", center=True)
    dvwho.add(put(wi, i + 1, 2, fill=True)); put(wi, i + 1, 3, fill=True); put(wi, i + 1, 4, fill=True)

# ---- Қорытынды ----
wq = wb.create_sheet("Қорытынды")
for col, w in zip("ABCDE", [46, 14, 14, 14, 14]):
    wq.column_dimensions[col].width = w
row = 1
wq.cell(row, 1, "Сауалнама").font = BOLD
put(wq, row, 2, f"=COUNTA(Сауалнама!$B$2:$B${SV_LAST})", center=True); put(wq, row, 3, "оқушы")
row += 1
header(wq, row, ["Жаңа кітап туралы қайдан біледі", "Оқушылар, %"], [None] * 2)
s_first = row
for s, cl in zip(SOURCES, SRC_COL):
    row += 1
    put(wq, row, 1, s)
    put(wq, row, 2, f'=IFERROR(ROUND(COUNTIF(Сауалнама!${cl}$2:${cl}${SV_LAST},1)/COUNT(Сауалнама!${cl}$2:${cl}${SV_LAST})*100,1),"")', center=True)
ch = BarChart(); ch.type = "bar"; ch.title = "Жаңа кітап туралы қайдан біледі, %"
ch.add_data(Reference(wq, min_col=2, min_row=s_first, max_row=row), titles_from_data=True)
ch.set_categories(Reference(wq, min_col=1, min_row=s_first + 1, max_row=row)); ch.height = 7; ch.width = 16; ch.legend = None
wq.add_chart(ch, "G1")
row += 2
header(wq, row, ["Сұрақ / жауап", "Оқушылар, %"], [None] * 2)
for code, q, opts in SURVEY[1:]:
    cl = COL[code]
    for o in opts:
        row += 1
        put(wq, row, 1, f"{q} – {o}")
        put(wq, row, 2, f'=IFERROR(ROUND(COUNTIF(Сауалнама!${cl}$2:${cl}${SV_LAST},"{o}")/COUNTA(Сауалнама!${cl}$2:${cl}${SV_LAST})*100,1),"")', center=True)
row += 2
wq.cell(row, 1, "Тәжірибе: 4 апталық қосынды").font = BOLD
row += 1
header(wq, row, ["Көрсеткіш", "Эксп. дейін", "Эксп. кезінде", "Бақ. дейін", "Бақ. кезінде"], [None] * 5)
e_first = row
for name, ce, cc in [("Алынған кітап", "C", "D"), ("Соңына дейін оқылған", "E", "F"), ("Пікір / бейне", "G", "H")]:
    row += 1
    put(wq, row, 1, name)
    for j, (cl, ph) in enumerate([(ce, "Дейін"), (ce, "Кезінде"), (cc, "Дейін"), (cc, "Кезінде")], 2):
        put(wq, row, j, f'=SUMIF(Тәжірибе!$B${EX_FIRST}:$B${EX_LAST},"{ph}",Тәжірибе!${cl}${EX_FIRST}:${cl}${EX_LAST})', center=True)
ch = BarChart(); ch.type = "col"; ch.title = "Тәжірибе: алынған кітап"
ch.add_data(Reference(wq, min_col=2, max_col=5, min_row=e_first, max_row=e_first + 1), titles_from_data=True)
ch.set_categories(Reference(wq, min_col=1, min_row=e_first + 1, max_row=e_first + 1)); ch.height = 7; ch.width = 16
wq.add_chart(ch, "G18")

# ---- Ескі кестеден көшіру ----
if len(sys.argv) > 2 and os.path.exists(sys.argv[2]):
    src = load_workbook(sys.argv[2], data_only=True)
    for sh in ["Сауалнама", "Тәжірибе", "Буктрейлерлер", "Блогерлер", "Сұхбат"]:
        if sh not in src.sheetnames:
            continue
        for r in range(1, wb[sh].max_row + 1):
            for c in range(1, wb[sh].max_column + 1):
                if wb[sh].cell(r, c).fill == FILL:
                    v = src[sh].cell(r, c).value
                    wb[sh].cell(r, c).value = v.strip() if isinstance(v, str) else v

# ---- Google Forms жауаптарын көшіру ----
if len(sys.argv) > 3:
    fr = load_workbook(sys.argv[3], data_only=True).worksheets[0]
    head = [str(fr.cell(1, c).value or "").strip() for c in range(1, fr.max_column + 1)]
    find = lambda title: next((i + 1 for i, h in enumerate(head) if h.startswith(title)), None)
    n = 0
    c_cons = find("Ата-анаң")
    for r in range(2, fr.max_row + 1):
        if all(fr.cell(r, c).value in (None, "") for c in range(2, fr.max_column + 1)):
            continue
        if c_cons and str(fr.cell(r, c_cons).value or "").strip() != "Иә":  # келісім жоқ – алынбайды
            continue
        n += 1
        rr = n + 1
        c = 2
        for k, (code, q, opts) in enumerate(SURVEY):
            j = find(q)
            v = str(fr.cell(r, j).value).strip() if j and fr.cell(r, j).value not in (None, "") else None
            col = 2 + k if k < 5 else 2 + len(SOURCES) + k
            wsv.cell(rr, col).value = v if v in opts else v
        j = find(SRC_Q.split(" (")[0])
        picked = str(fr.cell(r, j).value or "") if j else ""
        for t, s in enumerate(SOURCES):
            wsv.cell(rr, 7 + t).value = 1 if s in picked else 0
        for t, (code, q) in enumerate(TEXT_Q):
            j = find(q.split("?")[0])
            wsv.cell(rr, 7 + len(SOURCES) + len(SURVEY) - 5 + t).value = fr.cell(r, j).value if j else None
    print("survey rows from form:", n)

for sh in wb.worksheets:
    sh.page_setup.orientation = "landscape"
    sh.page_setup.paperSize = sh.PAPERSIZE_A4
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
wb.save(OUT)
print("written", OUT)
