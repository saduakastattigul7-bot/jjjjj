"""Жұмыстағы суреттер (figs/*.png). Деректер жоқ болса, диаграммалардың орнына бос жақтау шығады."""
import json, math, os, textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

D = json.load(open("data.json", encoding="utf8"))
N = D["natije"]
OUT = "figs"
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#8a8984", "xtick.color": "#52514e", "ytick.color": "#52514e"})
INK, MUTED, GRID = "#0b0b0b", "#52514e", "#e4e3de"
S = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
TEAL, GOLD, TEAL_L, GOLD_L, GREEN_L = "#0b5c6b", "#c88a00", "#e3f1f3", "#fdf3dc", "#e3f3ea"


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


# Зерттеудің кезеңдері
fig, ax = plt.subplots(figsize=(11, 2.6))
ax.set_xlim(0, 11.2); ax.set_ylim(0, 2.6); ax.axis("off")
steps = ["1-кезең\nТақырып,\nтеория", "2-кезең\nСауалнама", "3-кезең\nБуктрейлер\nтүсіру", "4-кезең\n4 апталық\nтәжірибе", "5-кезең\nСұхбат,\nталдау"]
for i, t in enumerate(steps):
    fc, ec = (TEAL_L, TEAL) if i < 1 else ((GOLD_L, GOLD) if i < 4 else (GREEN_L, "#1a7f4b"))
    box(ax, 0.1 + i * 2.22, 0.3, 1.9, 2.0, t, fc, ec, fs=10.5)
    if i < 4:
        ax.annotate("", xy=(2.30 + i * 2.22, 1.3), xytext=(2.02 + i * 2.22, 1.3), arrowprops=dict(arrowstyle="-|>", color="#6b6a66", lw=1.5))
save(fig, "kezender")

# «Оқу шынжыры» моделі
fig, ax = plt.subplots(figsize=(8, 6.4))
ax.set_xlim(-4.2, 4.2); ax.set_ylim(-3.6, 3.6); ax.set_aspect("equal"); ax.axis("off")
nodes = ["Бейнешолу\nкөру", "Кітапты\nалу", "Оқып\nшығу", "Өз пікірін\nайту", "Жаңа\nбейнешолу\nтүсіру"]
cols = [S[0], S[3], S[2], S[1], S[4]]
Rr = 2.5
pts = [(Rr * math.sin(2 * math.pi * k / 5), Rr * math.cos(2 * math.pi * k / 5)) for k in range(5)]
for k in range(5):
    (x1, y1), (x2, y2) = pts[k], pts[(k + 1) % 5]
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), connectionstyle="arc3,rad=-0.25", arrowstyle="-|>", mutation_scale=18,
                                 color="#8a8984", lw=1.8, shrinkA=48, shrinkB=48))
for (x, y), t, c in zip(pts, nodes, cols):
    ax.add_patch(plt.Circle((x, y), 0.95, fc="white", ec=c, lw=3))
    ax.text(x, y, t, ha="center", va="center", fontsize=11, fontweight="bold", color=INK, linespacing=1.25)
ax.text(0, 0.15, "Оқу\nшынжыры", ha="center", va="center", fontsize=15, fontweight="bold", color=TEAL, linespacing=1.2)
ax.text(0, -0.75, "әр оқырман келесі\nоқырманды тудырады", ha="center", va="center", fontsize=9.5, color=MUTED, style="italic")
save(fig, "shynzhyr")

# Сауалнама
if N.get("n_resp"):
    src = sorted(N["src"].items(), key=lambda kv: kv[1] or 0)
    fig, ax = plt.subplots(figsize=(9, 3.6))
    ax.barh([k for k, _ in src], [v or 0 for _, v in src], color=[S[1] if k == "Әлеуметтік желі" else S[0] for k, _ in src], height=0.6)
    for y, (_, v) in enumerate(src):
        ax.text((v or 0) + 1, y, f"{(v or 0):.0f}%", va="center", fontsize=10.5)
    ax.set_xlim(0, 100); ax.set_xlabel("Оқушылар үлесі, % (бірнеше жауап)"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    save(fig, "e_src")
    sv = N["survey"]
    items = [("Бағдарламадан тыс кітап оқыған", sv["READ"]["Иә"]), ("«Буктокер» сөзін естіген", sv["HEARD"]["Иә"]),
             ("Блогер ұсынған кітапты оқыған", sv["TRIED"]["Иә"]), ("Оны соңына дейін оқыған*", sv["FINISH"]["Иә"]),
             ("Қазақша кітап блогерін көреді", sv["KAZ"]["Иә"]), ("Өзі бейне түсірер еді", sv["VIDEO"]["Иә"])]
    fig, ax = plt.subplots(figsize=(9, 3.8))
    ax.barh([a for a, _ in items][::-1], [b or 0 for _, b in items][::-1], color=S[2], height=0.6)
    for y, (_, v) in enumerate(items[::-1]):
        ax.text((v or 0) + 1, y, f"{(v or 0):.0f}%", va="center", fontsize=10.5)
    ax.set_xlim(0, 100); ax.set_xlabel("«Иә» деп жауап бергендер, %"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    save(fig, "e_survey")
else:
    placeholder("e_src", "Сурет осы жерде болады.\nСауалнама нәтижесі енгізілгеннен кейін\nдиаграмма салынады.")
    placeholder("e_survey", "Сурет осы жерде болады.\nСауалнама нәтижесі енгізілгеннен кейін\nдиаграмма салынады.")

# Тәжірибе
if N.get("exp_filled"):
    E = N["exp"]
    fig, ax = plt.subplots(figsize=(9, 4))
    labels = ["Алынған кітап", "Соңына дейін оқылған", "Пікір / бейне"]
    keys = [("e_take", "c_take"), ("e_fin", "c_fin"), ("e_rev", "c_rev")]
    w = 0.19
    for k, (lab, col) in enumerate([(f"{N['exp_name']}: дейін", "#a9c8f0"), (f"{N['exp_name']}: кезінде", S[0]),
                                     (f"{N['ctrl_name']}: дейін", "#f6c3a6"), (f"{N['ctrl_name']}: кезінде", S[1])]):
        vals = [E[ke if k < 2 else kc]["before" if k % 2 == 0 else "during"] for ke, kc in keys]
        bs = ax.bar([i + (k - 1.5) * (w + 0.01) for i in range(3)], vals, width=w, color=col, label=lab)
        for b, v in zip(bs, vals):
            ax.text(b.get_x() + w / 2, v + 0.3, str(v), ha="center", fontsize=9.5)
    ax.set_xticks(range(3)); ax.set_xticklabels(labels); ax.set_ylabel("Саны (4 апта)"); ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.legend(frameon=False, ncol=2, fontsize=9.5, loc="upper right")
    save(fig, "e_exp")
else:
    placeholder("e_exp", "Сурет осы жерде болады.\nТәжірибе деректері енгізілгеннен кейін\nдиаграмма салынады.")
print("figs done", "with data" if N else "placeholders")
