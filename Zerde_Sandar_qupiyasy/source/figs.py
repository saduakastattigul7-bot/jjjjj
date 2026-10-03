"""Жұмыстағы суреттер (figs/*.png). Деректер жоқ болса, диаграммалардың орнына бос жақтау шығады."""
import json, os, textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle

D = json.load(open("data.json", encoding="utf8"))
N = D["natije"]
OUT = "figs"
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#8a8984", "xtick.color": "#52514e", "ytick.color": "#52514e"})
INK, GRID = "#0b0b0b", "#e4e3de"
BLUE, ORANGE, GOLD = "#2a78d6", "#eb6834", "#c88a00"
NAVY, NAVY_L, GOLD_L = "#1f3b73", "#e7edf8", "#fdf3dc"


def save(fig, name):
    fig.savefig(os.path.join(OUT, name + ".png"), dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def box(ax, x, y, w, h, text, fc, ec, fs=11, bold=False, color=INK):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08", fc=fc, ec=ec, lw=1.4))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=color, fontweight="bold" if bold else "normal", linespacing=1.35)


def placeholder(name, text):
    fig, ax = plt.subplots(figsize=(9, 3.4))
    ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.2, 0.2), 9.6, 3.6, boxstyle="round,pad=0.02,rounding_size=0.1", fc="#fff8dc", ec="#b08900", lw=1.5, ls="--"))
    ax.text(5, 2, text, ha="center", va="center", fontsize=13, color="#6b5200", linespacing=1.5)
    save(fig, name)


# 1-сурет: киелі сандар
KIELI = [("3", "Үш жүз\nүш ағайынды"), ("7", "жеті ата\nжеті қазына\nЖеті қарақшы"), ("9", "тоғыздап\nсыйлау\nтоғызқұмалақ"),
         ("12", "мүшел –\n12 жылдық\nжыл санау"), ("40", "қырқынан\nшығару\nқырқы")]
fig, ax = plt.subplots(figsize=(11, 3.6))
ax.set_xlim(0, 11); ax.set_ylim(0, 3.6); ax.axis("off")
for i, (n, t) in enumerate(KIELI):
    cx = 1.1 + i * 2.2
    ax.add_patch(Circle((cx, 2.6), 0.62, color=NAVY))
    ax.text(cx, 2.6, n, ha="center", va="center", fontsize=26, fontweight="bold", color="white")
    box(ax, cx - 0.95, 0.1, 1.9, 1.6, t, GOLD_L, GOLD, fs=10.5)
save(fig, "kieli")

# 2-сурет: кезеңдер
fig, ax = plt.subplots(figsize=(10.5, 2.6))
ax.set_xlim(0, 10.5); ax.set_ylim(0, 2.6); ax.axis("off")
steps = ["1-кезең\nТақырып,\nәдебиет", "2-кезең\nКиелі сандар,\nкодтау жүйесі", "3-кезең\n14 ертегіні\nоқып санау",
         "4-кезең\nЕсептеу,\nдиаграмма", "5-кезең\nТалдау,\nертегі-есептер"]
for i, t in enumerate(steps):
    box(ax, 0.1 + i * 2.1, 0.35, 1.8, 1.9, t, NAVY_L if i < 2 else (GOLD_L if i < 4 else "#e3f3ea"), NAVY if i < 2 else (GOLD if i < 4 else "#1a7f4b"), fs=10.5)
    if i < 4:
        ax.annotate("", xy=(2.2 + i * 2.1, 1.3), xytext=(1.9 + i * 2.1, 1.3), arrowprops=dict(arrowstyle="-|>", color="#6b6a66", lw=1.5))
save(fig, "kezender")

if N:
    nums = [str(n) for n in D["numbers"]]
    nv = [N["by_num"][n]["Н"] for n in nums] + [None]
    fv = [N["by_num"][n]["Ф"] for n in nums] + [None]
    labels = nums + ["басқа"]
    other = N["other"]
    other_f = sum(1 for x in N["rows"] if x["num"] not in D["numbers"] and x["func"] == "Ф")
    nv[-1], fv[-1] = other - other_f, other_f
    fig, ax = plt.subplots(figsize=(11.5, 4.6))
    xs = range(len(labels))
    cols_n = [BLUE] * len(labels); cols_f = [ORANGE] * len(labels)
    ax.bar(xs, nv, color=BLUE, label="Нақты сан (Н)", edgecolor="white", linewidth=1.5)
    ax.bar(xs, fv, bottom=nv, color=ORANGE, label="Формула, символ (Ф)", edgecolor="white", linewidth=1.5)
    for x, a, b in zip(xs, nv, fv):
        if a + b:
            ax.text(x, a + b + 0.3, str(a + b), ha="center", fontsize=10, color=INK)
    ax.set_xticks(list(xs)); ax.set_xticklabels(labels)
    for t, l in zip(ax.get_xticklabels(), labels):
        if l.isdigit() and int(l) in D["sacred"]:
            t.set_fontweight("bold"); t.set_color(NAVY)
    ax.set_xlabel("Сан («киелі» сандар қалың әріппен)"); ax.set_ylabel("Кездесу саны")
    ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True); ax.legend(frameon=False)
    save(fig, "e_freq")
    cats = sorted(D["cats"], key=lambda c: N["by_cat"][c])
    fig, ax = plt.subplots(figsize=(9, 4.2))
    vals = [N["by_cat"][c] for c in cats]
    ax.barh(cats, vals, color=BLUE, height=0.6)
    for y, v in enumerate(vals):
        ax.text(v + 0.3, y, str(v), va="center", fontsize=10.5)
    ax.set_xlabel("Сандар саны"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    save(fig, "e_cat")
else:
    placeholder("e_freq", "3-сурет осы жерде болады.\nZertteu_kestesi.xlsx толтырылғаннан кейін\nсандардың кездесуі көрсетіледі.")
    placeholder("e_cat", "4-сурет осы жерде болады.\nZertteu_kestesi.xlsx толтырылғаннан кейін\nсандардың мағынасы көрсетіледі.")

# 5-сурет: есеп үлгісі
fig, ax = plt.subplots(figsize=(8, 4.4))
ax.set_xlim(0, 8); ax.set_ylim(0, 4.4); ax.axis("off")
ax.add_patch(FancyBboxPatch((0.1, 0.1), 7.8, 4.2, boxstyle="round,pad=0.02,rounding_size=0.25", fc="white", ec=GOLD, lw=3))
ax.add_patch(Circle((0.95, 3.45), 0.55, color=NAVY))
ax.text(0.95, 3.45, "№1", ha="center", va="center", fontsize=20, fontweight="bold", color="white")
ax.text(1.75, 3.62, "ЕРТЕГІ-ЕСЕП", fontsize=18, fontweight="bold", color=NAVY, va="center")
ax.text(1.75, 3.2, "Тақырыбы: бүтін сандарды бөлу", fontsize=11, color="#52514e", va="center")
txt = ("Ертегілер көбіне «отыз күн ойын, қырық күн тойын жасапты» деп аяқталады. "
       "Ойын мен той қатар емес, бірінен соң бірі өтті делік. Той-думан барлығы неше аптаға созылды?")
ax.text(0.5, 2.55, "\n".join(textwrap.wrap(txt, 60)), fontsize=11.5, color=INK, va="top", linespacing=1.45)
ax.text(0.5, 0.55, "Жауабы: (30 + 40) : 7 = 10 апта", fontsize=12, color=GOLD, fontweight="bold")
save(fig, "esep")
print("figs done", "with data" if N else "placeholders")
