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
TEAL, GOLD, TEAL_L, GOLD_L, GREEN_L = "#0b5c6b", "#c88a00", "#e3f1f3", "#fdf3dc", "#e3f3ea"
ST_COL = {"Ж": S[2], "А": S[0], "Т": S[3], "Ұ": "#a3a29c"}
GEN_COL = {"О": S[0], "Ә": S[1], "Ү": S[2]}
GNAME = dict(D["groups"])


def save(fig, name):
    fig.savefig(os.path.join(OUT, name + ".png"), dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def box(ax, x, y, w, h, text, fc, ec, fs=11, bold=False, color=INK, ha="center"):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08", fc=fc, ec=ec, lw=1.4))
    tx = x + w / 2 if ha == "center" else x + 0.15
    ax.text(tx, y + h / 2, text, ha=ha, va="center", fontsize=fs, color=color, fontweight="bold" if bold else "normal", linespacing=1.35)


def arrow(ax, a, b, col="#6b6a66"):
    ax.annotate("", xy=b, xytext=a, arrowprops=dict(arrowstyle="-|>", color=col, lw=1.4))


def placeholder(name, text):
    fig, ax = plt.subplots(figsize=(9, 3.4))
    ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.2, 0.2), 9.6, 3.6, boxstyle="round,pad=0.02,rounding_size=0.1", fc="#fff8dc", ec="#b08900", lw=1.5, ls="--"))
    ax.text(5, 2, text, ha="center", va="center", fontsize=13, color="#6b5200", linespacing=1.5)
    save(fig, name)


# 1-сурет: көнерген сөздердің түрлері
fig, ax = plt.subplots(figsize=(11, 4.6))
ax.set_xlim(0, 11); ax.set_ylim(0, 4.8); ax.axis("off")
box(ax, 3.5, 3.9, 4.0, 0.75, "Көнерген сөздер", TEAL, TEAL, fs=14, bold=True, color="white")
box(ax, 0.3, 2.4, 4.6, 0.9, "Тарихи сөздер\nзаты да, атауы да жоқ", "white", S[3], fs=11.5, bold=True)
box(ax, 6.1, 2.4, 4.6, 0.9, "Архаизмдер\nзаты бар, атауы ескірген", "white", S[0], fs=11.5, bold=True)
arrow(ax, (5.0, 3.9), (2.6, 3.32)); arrow(ax, (6.0, 3.9), (8.4, 3.32))
box(ax, 0.3, 0.9, 4.6, 1.2, "сауыт, дулыға, болыс,\nауылнай, батпан, сәукеле", "#f7f7f5", "#d6d5cf", fs=11.5)
box(ax, 6.1, 0.9, 4.6, 1.2, "ләшкер → әскер\nтамұқ → тозақ\nмейман → қонақ", "#f7f7f5", "#d6d5cf", fs=11.5)
box(ax, 2.6, 0.05, 5.8, 0.6, "Кейбірі қайта жанданады: әкім, теңге, сарбаз, құрылтай", GREEN_L, S[2], fs=11, bold=True)
save(fig, "turleri")

# 2-сурет: кезеңдер
fig, ax = plt.subplots(figsize=(11, 2.6))
ax.set_xlim(0, 11.2); ax.set_ylim(0, 2.6); ax.axis("off")
steps = ["1-кезең\nТақырып,\nәдебиет", "2-кезең\n36 сөз,\nсөздікпен\nтексеру", "3-кезең\nБАҚ\nмониторингі",
         "4-кезең\nҮш буын\nтесті", "5-кезең\nТалдау,\nмини-сөздік"]
for i, t in enumerate(steps):
    fc, ec = (TEAL_L, TEAL) if i < 2 else ((GOLD_L, GOLD) if i < 4 else (GREEN_L, "#1a7f4b"))
    box(ax, 0.1 + i * 2.22, 0.3, 1.9, 2.0, t, fc, ec, fs=10.5)
    if i < 4:
        arrow(ax, (2.02 + i * 2.22, 1.3), (2.30 + i * 2.22, 1.3))
save(fig, "kezender")

# 3-сурет: мәртебе ережесі
RARE = D["rare"]
fig, ax = plt.subplots(figsize=(11, 4.4))
ax.set_xlim(0, 11); ax.set_ylim(0, 4.6); ax.axis("off")
q = [(f"egemen.kz-те соңғы\nжылы {RARE}-тен аз ба?", "Ұ", "Ұмыт болып\nбарады"),
     ("Мысал жаңа мағынада\n(«Ж») ма?", "Ж", "Қайта\nжанданған"),
     ("Мысал атау ішінде\n(«А») ме?", "А", "Атауда\nсақталған")]
for i, (t, c, n) in enumerate(q):
    x = 0.1 + i * 3.0
    box(ax, x, 2.75, 2.55, 1.3, t, "white", TEAL, fs=10.5)
    box(ax, x + 0.35, 0.55, 1.85, 1.1, f"{c} – {n}", "white", ST_COL[c], fs=10.5, bold=True)
    arrow(ax, (x + 1.27, 2.75), (x + 1.27, 1.67)); ax.text(x + 1.4, 2.15, "иә", fontsize=10, color=MUTED)
    arrow(ax, (x + 2.57, 3.4), (x + 3.08, 3.4)); ax.text(x + 2.62, 3.55, "жоқ", fontsize=10, color=MUTED)
box(ax, 9.15, 2.85, 1.75, 1.1, "Т – Тарихи\nмәтінде ғана", "white", ST_COL["Т"], fs=10.5, bold=True)
save(fig, "mertebe")

W = D["words"]
if N:
    words = N["words"]
    # 4-сурет: мәртебе топтар бойынша
    fig, ax = plt.subplots(figsize=(9.5, 4.4))
    groups = [g for g, _ in D["groups"]][::-1]
    left = [0] * len(groups)
    for c, n, _ in D["status"]:
        vals = [N["by_group_status"][g][c] for g in groups]
        ax.barh([GNAME[g] for g in groups], vals, left=left, color=ST_COL[c], label=f"{c} – {n}", height=0.62, edgecolor="white", linewidth=1.5)
        for y, (l, v) in enumerate(zip(left, vals)):
            if v:
                ax.text(l + v / 2, y, str(v), ha="center", va="center", color="white", fontsize=10.5, fontweight="bold")
        left = [a + b for a, b in zip(left, vals)]
    ax.set_xlabel("Сөз саны"); ax.set_xlim(0, 6); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.45, -0.13), ncol=4, fontsize=10)
    save(fig, "e_status")
    # 5-сурет: БАҚ-тағы жиілік (ең жиі 15 сөз)
    top = sorted([x for x in words if x["h1"] is not None], key=lambda x: x["hits"])[-15:]
    fig, ax = plt.subplots(figsize=(9, 0.9 + 0.36 * len(top)))
    ax.barh([x["word"] for x in top], [x["hits"] for x in top], color=[ST_COL.get(x["status"], S[0]) for x in top], height=0.62)
    for y, x in enumerate(top):
        ax.text(x["hits"] + max(t["hits"] for t in top) * 0.01 + 0.1, y, str(x["hits"]), va="center", fontsize=10, color=INK)
    ax.set_xlabel("egemen.kz: соңғы жылдағы іздеу нәтижесі (шамамен)"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    save(fig, "e_hits")
    # Сөздіктегі белгілер
    fig, ax = plt.subplots(figsize=(9, 2.8))
    marks = ["көн.", "тар.", "белгі жоқ", "сөздікте жоқ"]; mc = [S[3], S[1], S[0], "#a3a29c"]
    kinds = [("А", "Архаизмдер"), ("Т", "Тарихи сөздер")]
    left = [0, 0]
    for m, col in zip(marks, mc):
        vals = [N["marks_by_kind"][k].get(m, 0) for k, _ in kinds]
        ax.barh([n for _, n in kinds], vals, left=left, color=col, label=m, height=0.55, edgecolor="white", linewidth=1.5)
        for y, (l, v) in enumerate(zip(left, vals)):
            if v:
                ax.text(l + v / 2, y, str(v), ha="center", va="center", color="white", fontsize=11, fontweight="bold")
        left = [a + b for a, b in zip(left, vals)]
    ax.set_xlabel("Сөз саны"); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.3), ncol=4, fontsize=10, title="Сөздіктегі белгі", title_fontsize=10)
    save(fig, "e_marks")
else:
    placeholder("e_status", "Сурет осы жерде болады.\nБАҚ мониторингі енгізілгеннен кейін\nсөздердің мәртебесі көрсетіледі.")
    placeholder("e_hits", "Сурет осы жерде болады.\nБАҚ мониторингі енгізілгеннен кейін\nсөздердің жиілігі көрсетіледі.")
    placeholder("e_marks", "Сурет осы жерде болады.\nСөздікпен тексеру нәтижесі енгізілгеннен кейін\nдиаграмма салынады.")

if N and N.get("n_people"):
    fig, ax = plt.subplots(figsize=(10, 4.4))
    tw = D["test_words"]; w = 0.27
    for k, (g, gname) in enumerate(D["gens"]):
        vals = [N["test"][x][g]["ok"] or 0 for x in tw]
        ax.bar([i + (k - 1) * w for i in range(len(tw))], vals, width=w, color=GEN_COL[g], label=gname)
    ax.set_xticks(range(len(tw))); ax.set_xticklabels(tw, rotation=30, ha="right")
    ax.set_ylim(0, 100); ax.set_ylabel("Дұрыс түсіндіргендер, %"); ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=3, fontsize=10)
    save(fig, "e_test")
else:
    placeholder("e_test", "Сурет осы жерде болады.\nТест нәтижелері енгізілгеннен кейін\nүш буын салыстырылады.")

# 8-сурет: «Сөз паспорты» үлгісі
pick = "сарбаз"
info = next(x for x in W if x[0] == pick)
wd = next((x for x in N.get("words", []) if x["word"] == pick), None) if N else None
ST_NAME = {c: n for c, n, _ in D["status"]}
fig, ax = plt.subplots(figsize=(8, 5))
ax.set_xlim(0, 8); ax.set_ylim(0, 5); ax.axis("off")
ax.add_patch(FancyBboxPatch((0.1, 0.1), 7.8, 4.8, boxstyle="round,pad=0.02,rounding_size=0.25", fc="white", ec=GOLD, lw=3))
ax.add_patch(FancyBboxPatch((0.1, 3.85), 7.8, 1.05, boxstyle="round,pad=0.02,rounding_size=0.25", fc=TEAL, ec=TEAL, lw=0))
ax.text(0.5, 4.38, pick.upper(), fontsize=24, fontweight="bold", color="white", va="center")
ax.text(7.5, 4.38, "Сөз паспорты", fontsize=12, color="#ffe9a8", va="center", ha="right")
rows = [("Мағынасы", info[3]), ("Түрі", "архаизм" if info[2] == "А" else "тарихи сөз"),
        ("Бүгінгі мәртебесі", f"{wd['status']} – {ST_NAME[wd['status']]}" if wd and wd["status"] else "[мониторингтен кейін]"),
        ("Бүгінгі мысал", (wd["example"] if wd and wd["example"] else "[БАҚ-тан алынған сөйлем]"))]
y = 3.45
for k, v in rows:
    ax.text(0.5, y, k, fontsize=11.5, fontweight="bold", color=TEAL, va="top")
    lines = textwrap.wrap(v, 46)[:3]
    ax.text(2.75, y, "\n".join(lines), fontsize=11, color=INK, va="top", linespacing=1.35, style="italic" if k == "Бүгінгі мысал" else "normal")
    y -= 0.42 + 0.33 * len(lines)
save(fig, "pasport")
print("figs done", "with data" if N else "placeholders")
