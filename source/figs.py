import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle
import os

OUT = os.path.join(os.path.dirname(__file__), "figs")
os.makedirs(OUT, exist_ok=True)
plt.rcParams["font.family"] = "DejaVu Sans"

INK = "#1f2937"
BLUE = "#2f5d8a"
GOLD = "#b7791f"
GREEN = "#2f7d5b"
LIGHT_B = "#e6eef6"
LIGHT_G = "#fbf1de"
LIGHT_GR = "#e3f1ea"


def box(ax, x, y, w, h, text, fc, ec, fs=10, bold=False):
    p = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.06",
                       fc=fc, ec=ec, lw=1.4)
    ax.add_patch(p)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs,
            color=INK, fontweight="bold" if bold else "normal", wrap=True)


def arrow(ax, x1, y1, x2, y2):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", color="#555", lw=1.4))


# 1-сурет. Зерттеу кезеңдері
fig, ax = plt.subplots(figsize=(10, 3.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 3.2); ax.axis("off")
steps = [
    ("1-кезең\nМәселені анықтау\nжәне шолу", LIGHT_B, BLUE),
    ("2-кезең\nКорпус құру\n(22 мәтін бірлігі)", LIGHT_B, BLUE),
    ("3-кезең\nКодтау және\nсалыстыру", LIGHT_G, GOLD),
    ("4-кезең\nМодель мен\n«ТОҚТА» алгоритмі", LIGHT_GR, GREEN),
    ("5-кезең\nПилоттық\nапробация", LIGHT_GR, GREEN),
]
w, gap = 1.7, 0.3
for i, (t, fc, ec) in enumerate(steps):
    x = 0.15 + i * (w + gap)
    box(ax, x, 0.9, w, 1.5, t, fc, ec, fs=10)
    if i < len(steps) - 1:
        arrow(ax, x + w + 0.02, 1.65, x + w + gap - 0.02, 1.65)
ax.text(5, 0.35, "Теориялық бөлім  →  әдістеме  →  нәтиже  →  практикаға енгізу",
        ha="center", fontsize=9.5, color="#555", style="italic")
fig.tight_layout(); fig.savefig(f"{OUT}/fig1.png", dpi=200); plt.close(fig)

# 2-сурет. Тақырыптық кодтар жиілігі
themes = ["Білім,\nақыл", "Ар,\nадалдық", "Мейірім,\nрақым", "Еңбек,\nқайрат",
          "Өзіне\nесеп беру", "Достық,\nел, қоғам", "Жаман\nәдеттен\nсақтану"]
abai = [9, 2, 2, 6, 3, 3, 3]
push = [1, 6, 5, 3, 4, 4, 6]
fig, ax = plt.subplots(figsize=(10, 4.6))
import numpy as np
x = np.arange(len(themes)); bw = 0.38
b1 = ax.bar(x - bw / 2, abai, bw, color=BLUE, label="Абай (11 мәтін бірлігі)")
b2 = ax.bar(x + bw / 2, push, bw, color=GOLD, label="Пушкин (11 мәтін бірлігі)")
for bars in (b1, b2):
    for b in bars:
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.15, str(int(b.get_height())),
                ha="center", fontsize=9.5, color=INK)
ax.set_xticks(x); ax.set_xticklabels(themes, fontsize=9.5)
ax.set_ylabel("Код кездескен мәтін бірліктерінің саны", fontsize=9.5)
ax.set_ylim(0, 10.5)
ax.spines[["top", "right"]].set_visible(False)
ax.yaxis.grid(True, color="#ddd"); ax.set_axisbelow(True)
ax.legend(frameon=False, fontsize=9.5, loc="upper right")
fig.tight_layout(); fig.savefig(f"{OUT}/fig2.png", dpi=200); plt.close(fig)

# 3-сурет. Ортақ және дара ұстанымдар
fig, ax = plt.subplots(figsize=(9, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.6); ax.axis("off"); ax.set_aspect("equal")
ax.add_patch(Circle((3.6, 2.7), 2.5, fc=BLUE, alpha=0.13, ec=BLUE, lw=1.6))
ax.add_patch(Circle((6.4, 2.7), 2.5, fc=GOLD, alpha=0.15, ec=GOLD, lw=1.6))
ax.text(2.3, 5.0, "АБАЙ", fontsize=13, fontweight="bold", color=BLUE, ha="center")
ax.text(7.7, 5.0, "ПУШКИН", fontsize=13, fontweight="bold", color=GOLD, ha="center")
ax.text(2.35, 2.7, "• «Толық адам»:\n  ақыл–қайрат–жүрек\n• Бес дұшпан / бес\n  асыл іс тізімі\n• Тура үндеу,\n  ақыл-кеңес\n• Білім — ұлт\n  болашағының кілті",
        fontsize=10.5, va="center", ha="center", color=INK)
ax.text(7.65, 2.7, "• «Береги честь\n  смолоду»\n• Кейіпкердің\n  таңдауы арқылы үлгі\n• «Милость к падшим»\n• Лицейлік достық,\n  ерік пен ар",
        fontsize=10.5, va="center", ha="center", color=INK)
ax.text(5.0, 2.7, "ОРТАҚ\n\nАдамгершілік\nбілімнен жоғары\n\nЖастық —\nмінез қалыптасар\nшешуші кезең\n\nАдалдық, еңбек,\nмейірім, өзіне\nесеп беру",
        fontsize=10, va="center", ha="center", color=INK, fontweight="bold")
fig.tight_layout(); fig.savefig(f"{OUT}/fig3.png", dpi=200); plt.close(fig)

# 4-сурет. «Ар – Ақыл – Жүрек» моделі
fig, ax = plt.subplots(figsize=(10, 5.6))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.8); ax.axis("off")
box(ax, 0.3, 4.3, 9.4, 1.2,
    "1-деңгей. ҚҰНДЫЛЫҚ ӨЗЕГІ\nАбай: ақыл · қайрат · жүрек;  бес асыл іс        Пушкин: ар (честь) · мейірім · достық",
    LIGHT_B, BLUE, fs=10)
box(ax, 0.3, 2.35, 9.4, 1.45,
    "2-деңгей. ШЕШІМ АЛГОРИТМІ «ТОҚТА»\nТ — Тоқта   ·   О — Ойлан: қай құндылық сынға түсті?   ·   Қ — Қайтарымын көр\n"
    "Т — Таңда да, істе   ·   А — Артынан өзіңнен есеп ал",
    LIGHT_G, GOLD, fs=10)
box(ax, 0.3, 0.3, 2.9, 1.5, "3-деңгей. ЖАҒДАЯТ\nКөшіріп жазу\n(академиялық адалдық)", LIGHT_GR, GREEN, fs=9.5)
box(ax, 3.55, 0.3, 2.9, 1.5, "3-деңгей. ЖАҒДАЯТ\nЖелідегі қорлау\n(кибербуллинг)", LIGHT_GR, GREEN, fs=9.5)
box(ax, 6.8, 0.3, 2.9, 1.5, "3-деңгей. ЖАҒДАЯТ\nТоп қысымы\n(«бәрі істеп жүр»)", LIGHT_GR, GREEN, fs=9.5)
arrow(ax, 5, 4.28, 5, 3.83)
for cx in (1.75, 5.0, 8.25):
    arrow(ax, 5 if cx == 5.0 else cx, 2.33, cx, 1.83)
fig.tight_layout(); fig.savefig(f"{OUT}/fig4.png", dpi=200); plt.close(fig)

# 5-сурет. «ТОҚТА» алгоритмінің блок-сызбасы
fig, ax = plt.subplots(figsize=(8.2, 7.6))
ax.set_xlim(0, 8.2); ax.set_ylim(0.3, 8.3); ax.axis("off"); ax.set_aspect("equal")
rows = [
    ("Т", "ТОҚТА", "Бірден жауап берме. 10 рет дем ал.\n«Ақырын жүріп, анық бас» (Абай)", LIGHT_B, BLUE),
    ("О", "ОЙЛАН", "Бұл жерде қай құндылық сынға түсіп тұр:\nадалдық, мейірім, ар, достық па?", LIGHT_B, BLUE),
    ("Қ", "ҚАЙТАРЫМЫН КӨР", "Салдары: өзіме — өзгеге — ертеңіме.\nГринёвтің не Швабриннің жолы ма?", LIGHT_G, GOLD),
    ("Т", "ТАҢДА ДА, ІСТЕ", "Ең болмаса бір нақты қадам жаса:\n«Егер ..., онда мен ... істеймін»", LIGHT_GR, GREEN),
    ("А", "АРТЫНАН ЕСЕП АЛ", "Кешке 3 сұрақ: не істедім? неге?\nкелесіде не өзгертемін? (15-қара сөз)", LIGHT_GR, GREEN),
]
top = 6.95
for i, (ltr, name, desc, fc, ec) in enumerate(rows):
    y = top - i * 1.6
    ax.add_patch(Circle((0.8, y + 0.55), 0.5, fc=ec, ec=ec))
    ax.text(0.8, y + 0.55, ltr, ha="center", va="center", fontsize=20, color="white", fontweight="bold")
    box(ax, 1.55, y, 6.4, 1.1, f"{name}\n{desc}", fc, ec, fs=9.5)
    if i < len(rows) - 1:
        arrow(ax, 0.8, y + 0.04, 0.8, y - 0.44)
fig.tight_layout(); fig.savefig(f"{OUT}/fig5.png", dpi=200); plt.close(fig)
print("ok")
