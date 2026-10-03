"""Жұмыстағы суреттер (figs/*.png). Деректер жоқ болса, диаграммалардың орнына бос жақтау шығады."""
import json, os, textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

D = json.load(open("data.json", encoding="utf8"))
N = D["natije"]
OUT = "figs"
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#8a8984", "xtick.color": "#52514e", "ytick.color": "#52514e"})
INK, MUTED, GRID = "#0b0b0b", "#52514e", "#e4e3de"
S = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
TEAL, GOLD, TEAL_L, GOLD_L = "#0b5c6b", "#c88a00", "#e3f1f3", "#fdf3dc"
TYPE_NAME = dict(D["types"])


def save(fig, name):
    fig.savefig(os.path.join(OUT, name + ".png"), dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def box(ax, x, y, w, h, text, fc, ec, fs=11, bold=False, color=INK, ha="center"):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08", fc=fc, ec=ec, lw=1.4))
    tx = x + w / 2 if ha == "center" else x + 0.15
    ax.text(tx, y + h / 2, text, ha=ha, va="center", fontsize=fs, color=color, fontweight="bold" if bold else "normal", linespacing=1.35)


def placeholder(name, text):
    fig, ax = plt.subplots(figsize=(9, 3.4))
    ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.2, 0.2), 9.6, 3.6, boxstyle="round,pad=0.02,rounding_size=0.1", fc="#fff8dc", ec="#b08900", lw=1.5, ls="--"))
    ax.text(5, 2, text, ha="center", va="center", fontsize=13, color="#6b5200", linespacing=1.5)
    save(fig, name)


# 1-сурет: ойын түрлері
fig, ax = plt.subplots(figsize=(11, 5.2))
ax.set_xlim(0, 11); ax.set_ylim(0, 5.4); ax.axis("off")
box(ax, 3.3, 4.4, 4.4, 0.8, "Қазақтың ұлттық ойындары", TEAL, TEAL, fs=14, bold=True, color="white")
cols = [S[0], S[1], S[2], S[3], S[4]]
for i, (code, tname) in enumerate(D["types"]):
    x = 0.1 + i * 2.18
    ax.annotate("", xy=(x + 1.0, 3.55), xytext=(5.5, 4.4), arrowprops=dict(arrowstyle="-|>", color="#8a8984", lw=1.2))
    box(ax, x, 2.75, 2.0, 0.8, "\n".join(textwrap.wrap(tname, 12)), "white", cols[i], fs=10.5, bold=True)
    games = [g for g, t, _ in D["games"] if t == code]
    box(ax, x, 0.15, 2.0, 2.4, "\n".join(games), "#f7f7f5", "#d6d5cf", fs=11)
save(fig, "turleri")

# 2-сурет: кезеңдер
fig, ax = plt.subplots(figsize=(10.5, 2.6))
ax.set_xlim(0, 10.5); ax.set_ylim(0, 2.6); ax.axis("off")
steps = ["1-кезең\nТақырып,\nәдебиет", "2-кезең\nОйындарды\nтоптастыру", "3-кезең\n15 шығарманы\nоқып талдау",
         "4-кезең\nСауалнама", "5-кезең\nНәтиже,\nкарточкалар"]
for i, t in enumerate(steps):
    box(ax, 0.1 + i * 2.1, 0.35, 1.8, 1.9, t, TEAL_L if i < 2 else (GOLD_L if i < 4 else "#e3f3ea"), TEAL if i < 2 else (GOLD if i < 4 else "#1a7f4b"), fs=10.5)
    if i < 4:
        ax.annotate("", xy=(2.2 + i * 2.1, 1.3), xytext=(1.9 + i * 2.1, 1.3), arrowprops=dict(arrowstyle="-|>", color="#6b6a66", lw=1.5))
save(fig, "kezender")

games = [g for g, _, _ in D["games"]]
if N:
    # 3-сурет: ойындардың кездесуі
    rows = sorted(games, key=lambda g: N["by_game"][g]["Ертегі"] + N["by_game"][g]["Жыр"])
    fig, ax = plt.subplots(figsize=(9, 5.2))
    e = [N["by_game"][g]["Ертегі"] for g in rows]; j = [N["by_game"][g]["Жыр"] for g in rows]
    ax.barh(rows, e, color=S[0], height=0.62, label="Ертегіде", edgecolor="white", linewidth=1.5)
    ax.barh(rows, j, left=e, color=S[1], height=0.62, label="Жырда", edgecolor="white", linewidth=1.5)
    for y, (a, b) in enumerate(zip(e, j)):
        if a + b:
            ax.text(a + b + 0.15, y, str(a + b), va="center", fontsize=10.5, color=INK)
    ax.set_xlabel("Эпизод саны"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.legend(frameon=False, loc="lower right")
    save(fig, "e_oyn")
    # 4-сурет: қызметі
    fig, ax = plt.subplots(figsize=(9, 4.2))
    labels = [f"{c} – {n}" for c, n, _ in D["funcs"]]
    xs = range(len(labels)); w = 0.38
    for k, (gname, col) in enumerate((("Ертегі", S[0]), ("Жыр", S[1]))):
        vals = [N["by_func"][c][gname] for c, _, _ in D["funcs"]]
        bs = ax.bar([x + (k - 0.5) * (w + 0.02) for x in xs], vals, width=w, color=col, label="Ертегіде" if k == 0 else "Жырда")
        for b, v in zip(bs, vals):
            ax.text(b.get_x() + w / 2, v + 0.15, str(v), ha="center", fontsize=10, color=INK)
    ax.set_xticks(list(xs)); ax.set_xticklabels(["\n".join(textwrap.wrap(l, 14)) for l in labels], fontsize=10)
    ax.set_ylabel("Эпизод саны"); ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.legend(frameon=False)
    save(fig, "e_func")
else:
    placeholder("e_oyn", "3-сурет осы жерде болады.\nZertteu_kestesi.xlsx толтырылғаннан кейін\nойындардың кездесуі көрсетіледі.")
    placeholder("e_func", "4-сурет осы жерде болады.\nZertteu_kestesi.xlsx толтырылғаннан кейін\nойынның қызметі көрсетіледі.")
if N and N.get("n_resp"):
    sv = N["survey"]
    rows = sorted(games, key=lambda g: sv[g]["know"] or 0)
    fig, ax = plt.subplots(figsize=(9, 6))
    h = 0.27
    for k, (key, lab, col) in enumerate((("know", "Біледі", S[0]), ("played", "Ойнаған", S[2]), ("folk", "Ертегіден оқыған", S[3]))):
        ax.barh([y + (1 - k) * h for y in range(len(rows))], [sv[g][key] or 0 for g in rows], height=h, color=col, label=lab)
    ax.set_yticks(range(len(rows))); ax.set_yticklabels(rows)
    ax.set_xlim(0, 100); ax.set_xlabel("Оқушылар үлесі, %"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3)
    save(fig, "e_survey")
else:
    placeholder("e_survey", "5-сурет осы жерде болады.\nСауалнама нәтижелері енгізілгеннен кейін\nдиаграмма салынады.")

# 6-сурет: карточка үлгісі (асық)
rule = next(r for g, _, r in D["games"] if g == "Асық")
ep = next((e for e in N.get("episodes", []) if e["game"] == "Асық"), None) if N else None
folk = f"«{ep['work']}»: {ep['episode']}" if ep else "[Ертегі немесе жырдан эпизод – зерттеу кестесінен алынады]"
fig, ax = plt.subplots(figsize=(8, 5))
ax.set_xlim(0, 8); ax.set_ylim(0, 5); ax.axis("off")
ax.add_patch(FancyBboxPatch((0.1, 0.1), 7.8, 4.8, boxstyle="round,pad=0.02,rounding_size=0.25", fc="white", ec=GOLD, lw=3))
ax.add_patch(FancyBboxPatch((0.1, 3.85), 7.8, 1.05, boxstyle="round,pad=0.02,rounding_size=0.25", fc=TEAL, ec=TEAL, lw=0))
ax.text(0.5, 4.38, "АСЫҚ", fontsize=24, fontweight="bold", color="white", va="center")
ax.text(7.5, 4.38, "Балалар ойыны", fontsize=12, color="#ffe9a8", va="center", ha="right")
ax.text(0.5, 3.45, "Ережесі", fontsize=12.5, fontweight="bold", color=TEAL)
ax.text(0.5, 3.15, "\n".join(textwrap.wrap(rule, 62)), fontsize=10.5, color=INK, va="top", linespacing=1.4)
ax.text(0.5, 1.75, "Ертегі мен жырда", fontsize=12.5, fontweight="bold", color=GOLD)
ax.text(0.5, 1.45, "\n".join(textwrap.wrap(folk, 62)[:4]), fontsize=10.5, color=INK, va="top", style="italic", linespacing=1.4)
save(fig, "kartochka")
print("figs done", "with data" if N else "placeholders")
