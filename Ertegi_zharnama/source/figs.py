"""Figures for the project: AIDA scheme, poster sketch, and placeholders."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle

plt.rcParams["font.family"] = "DejaVu Sans"
OUT = os.path.join(os.path.dirname(__file__), "figs")
os.makedirs(OUT, exist_ok=True)


def box(ax, x, y, w, h, text, fc, ec="#333333", fs=12, bold=False, ls="-", r=0.02):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                                fc=fc, ec=ec, lw=1.6, ls=ls))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs,
            fontweight="bold" if bold else "normal", wrap=True)


def aida():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=200)
    ax.set_xlim(0, 10); ax.set_ylim(0, 4.2); ax.axis("off")
    steps = [
        ("A", "Назар аудару", "«Сөйлейтін атты\nкөрдің бе?»", "#F6C85F"),
        ("I", "Қызығушылық", "ғажайып серіктер,\nжер асты елі", "#9DD866"),
        ("D", "Қалау", "«Сен де батыр\nболғың келе ме?»", "#6FB1E0"),
        ("A", "Әрекет", "«Кітапханадан\nал да, оқы!»", "#F08A8A"),
    ]
    w, gap = 2.0, 0.55
    for i, (L, name, ex, c) in enumerate(steps):
        x = 0.3 + i * (w + gap)
        ax.add_patch(FancyBboxPatch((x, 1.9), w, 1.9, boxstyle="round,pad=0,rounding_size=0.15",
                                    fc=c, ec="#333333", lw=1.6))
        ax.text(x + w / 2, 3.25, L, ha="center", va="center", fontsize=30, fontweight="bold")
        ax.text(x + w / 2, 2.35, name, ha="center", va="center", fontsize=11.5, fontweight="bold")
        ax.text(x + w / 2, 1.2, ex, ha="center", va="center", fontsize=11, style="italic")
        if i < 3:
            ax.add_patch(FancyArrowPatch((x + w + 0.05, 2.85), (x + w + gap - 0.05, 2.85),
                                         arrowstyle="-|>", mutation_scale=22, lw=2, color="#333333"))
    ax.text(5, 0.35, "Жоғарыда – қадамның атауы, төменде – менің жарнамамдағы мысал",
            ha="center", fontsize=10, color="#555555")
    fig.savefig(os.path.join(OUT, "aida.png"), bbox_inches="tight", facecolor="white")
    plt.close(fig)


def poster():
    fig, ax = plt.subplots(figsize=(6, 8.4), dpi=200)
    ax.set_xlim(0, 6); ax.set_ylim(0, 8.4); ax.axis("off")
    ax.add_patch(Rectangle((0.2, 0.2), 5.6, 8.0, fc="#FFF8E6", ec="#333333", lw=2))
    box(ax, 0.5, 6.9, 5.0, 1.0, "1. СҰРАҚ-ТАҚЫРЫП (үлкен әріппен)\n«Сөйлейтін атты көрдің бе?»", "#F6C85F", fs=11, bold=True)
    box(ax, 4.3, 5.4, 1.2, 1.2, "Самұрық\nқұс\n(сурет)", "#FFFFFF", fs=9, ls="--")
    box(ax, 0.5, 2.9, 3.6, 3.7, "2. БАСТЫ СУРЕТ\n\nТөстік\nШалқұйрыққа мініп\nшауып келеді", "#CFE8FB", fs=11, bold=True)
    box(ax, 4.3, 2.9, 1.2, 2.3, "Жер\nасты\nелі:\n«Онда\nне бар?»", "#E3D7F5", fs=9, ls="--")
    box(ax, 0.5, 1.6, 5.0, 1.0, "3. СЛОГАН\n«Ер Төстік – оқы да, батыр бол!»", "#9DD866", fs=11, bold=True)
    box(ax, 0.5, 0.45, 5.0, 0.9, "4. ШАҚЫРУ\n«Мектеп кітапханасынан ал да, оқы!»", "#F08A8A", fs=10, bold=True)
    fig.savefig(os.path.join(OUT, "poster.png"), bbox_inches="tight", facecolor="white")
    plt.close(fig)


def placeholder(name, text, ratio=0.6):
    fig, ax = plt.subplots(figsize=(8, 8 * ratio), dpi=150)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    ax.add_patch(Rectangle((0.01, 0.01), 0.98, 0.98, fc="#F2F2F2", ec="#888888", lw=2, ls="--"))
    ax.text(0.5, 0.5, text, ha="center", va="center", fontsize=15, color="#555555", wrap=True)
    fig.savefig(os.path.join(OUT, name + ".png"), bbox_inches="tight", facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    aida()
    poster()
    placeholder("photo", "ОСЫ ЖЕРГЕ\nөзің қолмен салған Б плакатыңның\nфотосуретін қой", 0.9)
    placeholder("chart1", "ОСЫ ЖЕРГЕ 1-диаграмманы қой\n\n(Derekter.xlsx → «Диаграммалар» парағы;\nсауалнаманың 3-сұрағы бойынша)")
    placeholder("chart2", "ОСЫ ЖЕРГЕ 2-диаграмманы қой\n\n(Derekter.xlsx → «Диаграммалар» парағы;\nжарнамаға дейін және кейін)")
    print("ok")
