"""Жұмыстағы суреттерді салады (figs/*.png). Тәжірибе деректері жоқ болса, 4- және 5-суреттің орнына бос жақтау шығады."""
import json, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

D = json.load(open("data.json", encoding="utf8"))
N = D["natije"]
OUT = "figs"
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.edgecolor": "#8a8984", "axes.labelcolor": "#2b2b29",
                     "xtick.color": "#52514e", "ytick.color": "#52514e"})
INK, MUTED, GRID = "#0b0b0b", "#52514e", "#e4e3de"
S = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]  # категориялық слоттар (ретімен)
BLUE_L, GREEN, GREEN_L, RED, RED_L = "#e6effa", "#1a7f4b", "#e3f3ea", "#c53030", "#fbe7e6"
AIS = [f"ЖИ-{k}" for k in (1, 2, 3)]
fmt = lambda v: f"{v:.1f}".replace(".", ",").replace(",0", "")


def save(fig, name):
    fig.savefig(os.path.join(OUT, name + ".png"), dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def box(ax, x, y, w, h, text, fc, ec, fs=11, bold=False, color=INK):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08", fc=fc, ec=ec, lw=1.4))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=color,
            fontweight="bold" if bold else "normal", linespacing=1.35)


def arrow(ax, x1, y1, x2, y2, c="#6b6a66"):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle="-|>", color=c, lw=1.5))


def placeholder(name, text):
    fig, ax = plt.subplots(figsize=(9, 3.6))
    ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.2, 0.2), 9.6, 3.6, boxstyle="round,pad=0.02,rounding_size=0.1",
                                fc="#fff8dc", ec="#b08900", lw=1.5, ls="--"))
    ax.text(5, 2, text, ha="center", va="center", fontsize=13, color="#6b5200", linespacing=1.5)
    save(fig, name)


# 1-сурет: сөзбе-сөз және бейнелі түсіну
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.4); ax.axis("off")
box(ax, 3.4, 4.45, 3.2, 0.75, "«Ит өлген жер»", BLUE_L, S[0], fs=15, bold=True)
arrow(ax, 4.2, 4.45, 2.5, 3.65); arrow(ax, 5.8, 4.45, 7.5, 3.65)
ax.text(2.05, 4.05, "сөзбе-сөз жол", ha="center", fontsize=10.5, color=MUTED, style="italic")
ax.text(7.95, 4.05, "тұтас тіркес ретінде", ha="center", fontsize=10.5, color=MUTED, style="italic")
box(ax, 0.3, 2.55, 4.4, 1.05, "ит  +  өлген  +  жер\n= «иттің өлген жері»", RED_L, RED, fs=12)
box(ax, 5.3, 2.55, 4.4, 1.05, "бейнелі мағына\n= «өте алыс, шалғай жер»", GREEN_L, GREEN, fs=12)
ax.text(2.5, 2.2, "✗  мағына бұзылды", ha="center", fontsize=11.5, color=RED, fontweight="bold")
ax.text(7.5, 2.2, "✓  сөздіктегі мағына", ha="center", fontsize=11.5, color=GREEN, fontweight="bold")
reasons = ["1. Мағынасы сөздердің\nқосындысына тең емес", "2. Мәтіндерде сирек\nкездеседі", "3. ЖИ білмегенін\nойдан шығаруы мүмкін"]
for i, r in enumerate(reasons):
    box(ax, 0.3 + i * 3.2, 0.25, 3.0, 1.1, r, "#f4f3ef", "#b9b8b2", fs=10.5)
ax.text(5, 1.6, "ЖИ-дің қателесу себептері", ha="center", fontsize=11, color=INK, fontweight="bold")
save(fig, "sozbesoz")

# 2-сурет: зерттеу кезеңдері
fig, ax = plt.subplots(figsize=(10.5, 2.6))
ax.set_xlim(0, 10.5); ax.set_ylim(0, 2.6); ax.axis("off")
steps = ["1-кезең\nТақырып және\nәдебиет", f"2-кезең\nСөздік тізбесі\n({D['taldau']['total']} тіркес)",
         "3-кезең\n30 тіркес,\nбағалау\nшкаласы", "4-кезең\nТәжірибе:\n3 ЖИ,\n183 жауап", "5-кезең\nТалдау,\nжадынама,\nқорытынды"]
fills = [BLUE_L, BLUE_L, "#fdf0e0", "#fdf0e0", GREEN_L]
edges = [S[0], S[0], "#c77d1a", "#c77d1a", GREEN]
for i, (t, f, e) in enumerate(zip(steps, fills, edges)):
    box(ax, 0.1 + i * 2.1, 0.35, 1.8, 1.9, t, f, e, fs=10)
    if i < 4:
        arrow(ax, 1.9 + i * 2.1, 1.3, 2.2 + i * 2.1, 1.3)
save(fig, "kezender")

# 3-сурет: дене мүшелері
body = sorted(D["taldau"]["body"].items(), key=lambda kv: kv[1])
fig, ax = plt.subplots(figsize=(9, 5))
names, vals = [k for k, _ in body], [v for _, v in body]
bars = ax.barh(names, vals, color=S[0], height=0.62)
for b, v in zip(bars, vals):
    ax.text(v + 0.4, b.get_y() + b.get_height() / 2, str(v), va="center", fontsize=10.5, color=INK)
ax.set_xlabel(f"Тіркес саны (барлығы {D['taldau']['total']} тіркес)")
ax.xaxis.grid(True, color=GRID, lw=0.8); ax.set_axisbelow(True)
ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
save(fig, "dene")

# 4-сурет: 1-тәжірибе
if N:
    labels = ["А тобы\nмөлдір бейнелі", "Ә тобы\nастарлы бейнелі", "Б тобы\nұлттық-мәдени"]
    fig, ax = plt.subplots(figsize=(9, 4.6))
    w = 0.26
    for k in range(3):
        vals = [N.get(f"e1_{g}_{k + 1}") or 0 for g in "АӘБ"]
        xs = [i + (k - 1) * (w + 0.02) for i in range(3)]
        bs = ax.bar(xs, vals, width=w, color=S[k], label=f"ЖИ-{k + 1}" + ("" if str(N["ai_names"][k]).startswith("ЖИ-") else f": {N['ai_names'][k]}"))
        for b, v in zip(bs, vals):
            ax.text(b.get_x() + w / 2, v + 1.5, fmt(v), ha="center", fontsize=9.5, color=INK)
    ax.set_xticks(range(3)); ax.set_xticklabels(labels)
    ax.set_ylim(0, 110); ax.set_ylabel("Сапа көрсеткіші, %")
    ax.yaxis.grid(True, color=GRID, lw=0.8); ax.set_axisbelow(True)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=3, fontsize=9.5)
    save(fig, "e1")
else:
    placeholder("e1", "4-сурет осы жерде болады.\nTazhiribe_kestesi.xlsx толтырылғаннан кейін\nтоптар бойынша сапа көрсеткішінің диаграммасы салынады.")

# 5-сурет: 3-тәжірибе
if N:
    codes = D["e3_codes"]
    fig, ax = plt.subplots(figsize=(9, 3.8))
    left = [0, 0, 0]
    for ci, (code, name, _) in enumerate(codes):
        vals = [N.get(f"e3_{code}_{k}", 0) for k in (1, 2, 3)]
        ax.barh(AIS, vals, left=left, color=S[ci], height=0.55, label=f"{code} – {name}", edgecolor="white", linewidth=2)
        for y, (l, v) in enumerate(zip(left, vals)):
            if v >= 2:
                ax.text(l + v / 2, y, str(v), ha="center", va="center", fontsize=10, color="white", fontweight="bold")
        left = [l + v for l, v in zip(left, vals)]
    ax.set_xlim(0, D["e3_n"] * len(D["e3"])); ax.invert_yaxis()
    ax.set_xlabel("ЖИ атаған тіркестер саны (әр ЖИ-ге 25)")
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=2, fontsize=9.5)
    save(fig, "e3")
else:
    placeholder("e3", "5-сурет осы жерде болады.\nTazhiribe_kestesi.xlsx толтырылғаннан кейін\nЖИ атаған тіркестердің құрамы көрсетіледі.")

# 6-сурет: жадынама
fig, ax = plt.subplots(figsize=(9, 10.5))
ax.set_xlim(0, 9); ax.set_ylim(0, 10.5); ax.axis("off")
ax.add_patch(FancyBboxPatch((0.15, 0.15), 8.7, 10.2, boxstyle="round,pad=0.02,rounding_size=0.2", fc="white", ec=S[0], lw=2.5))
ax.text(4.5, 9.75, "ЖИ-ГЕ СЕН, БІРАҚ ТЕКСЕР!", ha="center", fontsize=21, fontweight="bold", color=S[0])
ax.text(4.5, 9.2, "ЖИ тұрақты тіркесті түсіндірсе, 4 қадаммен тексер", ha="center", fontsize=12.5, color=MUTED)
items = [
    ("1", "Сөздікпен салыстыр", "ЖИ берген мағынаны фразеологиялық\nнемесе түсіндірме сөздіктен тап.\nСөздікте жоқ тіркеске сенбе."),
    ("2", "Сөзбе-сөз емес пе?", "«Ит өлген жер» = «иттің өлген жері» ме?\nТіркесті сөздердің тура мағынасымен\nтүсіндірсе – бұл қатенің белгісі."),
    ("3", "Мысалды тексер", "ЖИ келтірген сөйлемде тіркес\nшын мәнінде сол мағынада тұр ма?"),
    ("4", "Үлкендерден сұра", "Күмәнің қалса, мұғаліміңнен,\nата-әжеңнен сұра: тұрақты тіркестер\nауызекі тілде сақталған."),
]
cols = [S[0], S[1], S[2], "#4a3aa7"]
for i, (n, h, t) in enumerate(items):
    y = 6.95 - i * 2.05
    ax.add_patch(FancyBboxPatch((0.6, y), 7.8, 1.8, boxstyle="round,pad=0.02,rounding_size=0.15", fc="#f7f7f5", ec="#d6d5cf", lw=1.2))
    ax.add_patch(plt.Circle((1.45, y + 0.9), 0.52, color=cols[i]))
    ax.text(1.45, y + 0.9, n, ha="center", va="center", fontsize=22, fontweight="bold", color="white")
    ax.text(2.3, y + 1.38, h, ha="left", va="center", fontsize=15, fontweight="bold", color=INK)
    ax.text(2.3, y + 0.62, t, ha="left", va="center", fontsize=11.2, color="#2b2b29", linespacing=1.3)
ax.text(4.5, 0.42, "Шешімді адам қабылдайды – ЖИ тек көмектеседі", ha="center", fontsize=12, style="italic", color=MUTED)
save(fig, "zhadynama")
print("figs done", "with data" if N else "placeholders for e1/e3")
