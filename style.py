import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib import rcParams


def apply_report_style():
    sns.set_theme(
        context="talk",
        style="whitegrid",
        font="AppleGothic",
        rc={
            "axes.facecolor": "#FAFAFA",
            "figure.facecolor": "white",
            "axes.edgecolor": "#DDDDDD",
            "grid.color": "#EAEAEA",
            "grid.linewidth": 0.8,
            "axes.titleweight": "bold",
            "axes.titlesize": 18,
            "axes.labelsize": 12,
            "xtick.labelsize": 10,
            "ytick.labelsize": 10,
            "legend.frameon": False,
            "figure.dpi": 160,
            "savefig.dpi": 300,
            "savefig.bbox": "tight",
        }
    )

    rcParams["axes.unicode_minus"] = False


def polish_axes(ax, title=None, subtitle=None, xlabel=None, ylabel=None, source=None):
    if title:
        ax.set_title(title, loc="left", pad=22, fontsize=18, fontweight="bold")

    if subtitle:
        ax.text(
            0, 1.02, subtitle,
            transform=ax.transAxes,
            ha="left",
            va="bottom",
            fontsize=10,
            color="#666666"
        )

    ax.set_xlabel(xlabel or "", labelpad=10)
    ax.set_ylabel(ylabel or "", labelpad=10)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#DDDDDD")
    ax.spines["bottom"].set_color("#DDDDDD")

    ax.grid(axis="y", alpha=0.45)
    ax.grid(axis="x", visible=False)

    if source:
        ax.text(
            0,
            -0.16,
            source,
            transform=ax.transAxes,
            ha="left",
            va="top",
            fontsize=8,
            color="#888888"
        )