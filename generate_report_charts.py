"""
Generate publication-quality charts for the Staff Standards FA report.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
import os

OUTPUT_DIR = os.path.expanduser("~/nhs-staff-standards/output")
DATA_DIR = os.path.expanduser("~/nhs-staff-standards/data")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# DHSC / GOV.UK colour palette
DHSC_TEAL = "#00A990"
DHSC_DARK_TEAL = "#008674"
DHSC_DEEPEST = "#005A4E"
GOVUK_RED = "#D4351C"
GOVUK_ORANGE = "#F47738"
GOVUK_PURPLE = "#4C2C92"
GOVUK_GREEN = "#00703C"
GOVUK_BLUE = "#1D70B8"
GOVUK_DARK_GREY = "#505A5F"
GOVUK_MID_GREY = "#B1B4B6"
GOVUK_LIGHT_GREY = "#F3F2F1"

STD_COLOURS = {
    "Violence": GOVUK_RED,
    "Racism": GOVUK_ORANGE,
    "SexSafety": GOVUK_PURPLE,
    "FlexWorking": GOVUK_GREEN,
    "LineMgmt": GOVUK_BLUE,
    "SupportEnv": DHSC_TEAL,
}

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Helvetica Neue", "Arial", "Helvetica"],
    "font.size": 10,
    "axes.titlesize": 12,
    "axes.labelsize": 10,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "axes.grid": False,
})


def chart_factor_correlations():
    """Heatmap of 6-factor CFA latent correlations."""
    labels = ["Violence", "Racism", "SexSafety", "FlexWorking", "LineMgmt", "SupportEnv"]
    display_labels = ["Violence", "Racism", "Sexual\nSafety", "Flexible\nWorking", "Line\nMgmt", "Supportive\nEnv"]

    corr = np.array([
        [ 1.000, -0.695,  0.955, -0.689, -0.675, -0.605],
        [-0.695,  1.000, -0.582,  0.775,  0.823,  0.865],
        [ 0.955, -0.582,  1.000, -0.675, -0.644, -0.543],
        [-0.689,  0.775, -0.675,  1.000,  0.976,  0.914],
        [-0.675,  0.823, -0.644,  0.976,  1.000,  0.989],
        [-0.605,  0.865, -0.543,  0.914,  0.989,  1.000],
    ])

    fig, ax = plt.subplots(figsize=(7, 6))

    # Use absolute values for colour intensity, with sign shown in text
    abs_corr = np.abs(corr)
    cmap = LinearSegmentedColormap.from_list("dhsc", ["#FFFFFF", DHSC_TEAL])
    im = ax.imshow(abs_corr, cmap=cmap, vmin=0, vmax=1, aspect="equal")

    ax.set_xticks(range(6))
    ax.set_yticks(range(6))
    ax.set_xticklabels(display_labels, fontsize=9)
    ax.set_yticklabels(display_labels, fontsize=9)

    for i in range(6):
        for j in range(6):
            val = corr[i, j]
            if i == j:
                text = "1.00"
                color = "white"
                weight = "normal"
            else:
                text = f"{val:+.2f}"
                color = "white" if abs(val) > 0.70 else GOVUK_DARK_GREY
                weight = "bold" if abs(val) > 0.90 else "normal"
            ax.text(j, i, text, ha="center", va="center",
                    fontsize=9, color=color, fontweight=weight)

    # Draw boxes around redundant clusters
    # FlexWorking/LineMgmt/SupportEnv (indices 3,4,5)
    rect1 = mpatches.FancyBboxPatch((2.5, 2.5), 3, 3, linewidth=2.5,
                                     edgecolor=GOVUK_RED, facecolor="none",
                                     boxstyle="round,pad=0.05")
    ax.add_patch(rect1)

    # Violence/SexSafety (indices 0,2)
    rect2 = mpatches.Rectangle((-0.5, -0.5), 1, 1, linewidth=0, fill=False)
    # Draw lines connecting 0,2
    for (r, c) in [(0, 2), (2, 0)]:
        circle = plt.Circle((c, r), 0.42, linewidth=2, edgecolor=GOVUK_RED,
                           facecolor="none", linestyle="--")
        ax.add_patch(circle)

    plt.colorbar(im, ax=ax, label="Absolute correlation", shrink=0.8)
    ax.set_title("Latent Factor Correlations (6-Factor CFA)\nBold = r > 0.90 (empirically redundant)",
                 fontsize=11, fontweight="bold", pad=15)

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "report_factor_correlations.png"), dpi=200, bbox_inches="tight")
    plt.close()
    print("  Saved: report_factor_correlations.png")


def chart_variance_decomposition():
    """Pie/donut chart showing 64% shared vs 36% specific variance."""
    fig, ax = plt.subplots(figsize=(5, 5))

    sizes = [64, 36]
    colours = [GOVUK_MID_GREY, DHSC_TEAL]
    labels = ["General 'good trust'\nfactor (64%)", "Standard-specific\nvariance (36%)"]

    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colours, autopct="%1.0f%%",
        startangle=90, pctdistance=0.55, labeldistance=1.15,
        wedgeprops=dict(width=0.5, edgecolor="white", linewidth=2),
        textprops=dict(fontsize=10),
    )
    for t in autotexts:
        t.set_fontsize(14)
        t.set_fontweight("bold")
        t.set_color("white")

    ax.set_title("Variance Decomposition\nAcross NHS Staff Standards",
                 fontsize=12, fontweight="bold", pad=20)

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "report_variance_decomposition.png"), dpi=200, bbox_inches="tight")
    plt.close()
    print("  Saved: report_variance_decomposition.png")


def chart_trust_profiles():
    """Bar chart showing example trust profiles (z-scores)."""
    means = pd.read_csv(os.path.join(DATA_DIR, "expanded_trust_means.csv"))

    # Compute composites
    mgr_items = ["mgr_encourages", "mgr_clear_feedback", "mgr_asks_opinion",
                 "mgr_interest_wellbeing", "mgr_values_work", "mgr_understanding_problems",
                 "mgr_listens", "mgr_cares_concerns", "mgr_effective_action"]
    means["mgr_composite"] = means[[c for c in mgr_items if c in means.columns]].mean(axis=1)

    composites = {
        "Violence": {"items": ["violence_patients", "harassment_patients",
                               "harassment_managers", "harassment_colleagues"], "reverse": True},
        "Racism": {"items": ["fair_career_progression", "discrim_patients",
                             "discrim_colleagues", "org_respects_differences"],
                   "reverse_items": ["discrim_patients", "discrim_colleagues"]},
        "FlexWorking": {"items": ["satisfaction_flexible_working",
                                  "good_wl_balance"], "reverse": False},
        "LineMgmt": {"items": ["mgr_composite", "had_appraisal",
                               "access_learning_dev"], "reverse": False},
        "SupportEnv": {"items": ["adequate_materials_equipment", "enough_staff",
                                 "org_positive_wellbeing"], "reverse": False},
    }

    scores = pd.DataFrame(index=means.index)
    scores["org_id"] = means["org_id"]

    for std, cfg in composites.items():
        available = [i for i in cfg["items"] if i in means.columns]
        vals = means[available].copy()
        if cfg.get("reverse"):
            vals = 1 - vals
        elif cfg.get("reverse_items"):
            for ri in cfg["reverse_items"]:
                if ri in vals.columns:
                    vals[ri] = 1 - vals[ri]
        scores[std] = vals.mean(axis=1)

    std_cols = list(composites.keys())
    for col in std_cols:
        scores[f"{col}_z"] = (scores[col] - scores[col].mean()) / scores[col].std()

    z_cols = [f"{c}_z" for c in std_cols]
    scores["profile_sd"] = scores[z_cols].std(axis=1)

    # Pick 4 interesting trusts
    # Most differentiated
    top = scores.nlargest(3, "profile_sd")
    # One near-average
    scores["abs_mean_z"] = scores[z_cols].mean(axis=1).abs()
    avg_trust = scores[(scores["profile_sd"] > 0.3) & (scores["abs_mean_z"] < 0.3)].head(1)
    example_trusts = pd.concat([top, avg_trust])

    fig, axes = plt.subplots(1, len(example_trusts), figsize=(14, 4.5), sharey=True)

    display_names = {
        "Violence": "Violence", "Racism": "Racism",
        "FlexWorking": "Flex\nWorking", "LineMgmt": "Line\nMgmt",
        "SupportEnv": "Supportive\nEnv"
    }

    for idx, (_, trust) in enumerate(example_trusts.iterrows()):
        ax = axes[idx]
        z_values = [trust[f"{s}_z"] for s in std_cols]
        colours = [STD_COLOURS[s] for s in std_cols]

        bars = ax.bar(range(len(std_cols)), z_values, color=colours,
                      edgecolor="white", linewidth=0.5, width=0.7)

        ax.axhline(y=0, color=GOVUK_DARK_GREY, linewidth=0.8, linestyle="-")
        ax.set_xticks(range(len(std_cols)))
        ax.set_xticklabels([display_names[s] for s in std_cols], fontsize=8)
        ax.set_title(f"Trust {trust['org_id']}", fontsize=10, fontweight="bold")

        for bar, val in zip(bars, z_values):
            y_pos = val + 0.08 if val >= 0 else val - 0.15
            ax.text(bar.get_x() + bar.get_width()/2, y_pos,
                    f"{val:+.1f}", ha="center", va="bottom" if val >= 0 else "top",
                    fontsize=8, fontweight="bold")

        ax.set_ylim(-3.5, 2.5)
        if idx == 0:
            ax.set_ylabel("Z-score (0 = average trust)", fontsize=9)

    fig.suptitle("Example Trust Profiles: Relative Strengths Across Standards",
                 fontsize=12, fontweight="bold", y=1.02)

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "report_trust_profiles.png"), dpi=200, bbox_inches="tight")
    plt.close()
    print("  Saved: report_trust_profiles.png")


def chart_cfa_comparison():
    """Bar chart of CFI across competing models."""
    models = ["1-factor", "3-factor\n(merged)", "6-factor\n(theory)"]
    cfi = [0.517, 0.569, 0.586]
    chi2_df = [19.2, 17.5, 18.1]

    fig, ax = plt.subplots(figsize=(6, 4))

    bars = ax.bar(models, cfi, color=[GOVUK_MID_GREY, DHSC_DARK_TEAL, DHSC_TEAL],
                  edgecolor="white", linewidth=1, width=0.5)

    ax.axhline(y=0.90, color=GOVUK_RED, linewidth=1.5, linestyle="--", alpha=0.7)
    ax.text(2.35, 0.91, "Acceptable\nfit (0.90)", fontsize=8, color=GOVUK_RED, va="bottom")

    for bar, val in zip(bars, cfi):
        ax.text(bar.get_x() + bar.get_width()/2, val + 0.01,
                f"{val:.3f}", ha="center", va="bottom", fontsize=10, fontweight="bold")

    ax.set_ylabel("Comparative Fit Index (CFI)")
    ax.set_ylim(0, 1.05)
    ax.set_title("CFA Model Comparison\nMore factors = better fit (but none reaches 'acceptable')",
                 fontsize=11, fontweight="bold")

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "report_cfa_comparison.png"), dpi=200, bbox_inches="tight")
    plt.close()
    print("  Saved: report_cfa_comparison.png")


def chart_standardised_loadings():
    """Horizontal bar chart of CFA standardised loadings per standard."""
    data = {
        "Violence": [
            ("q14a: Harassment from patients", 0.983),
            ("q13a: Violence from patients", 0.903),
            ("q14b: Harassment from managers", 0.614),
            ("q14c: Harassment from colleagues", 0.468),
        ],
        "Racism": [
            ("q15: Fair career progression", 0.953),
            ("q21: Respects differences", 0.883),
            ("q16b: Discrim. from colleagues", 0.866),
            ("q16a: Discrim. from patients", 0.723),
        ],
        "SexSafety": [
            ("q17a: Sexual behaviour, patients", 0.867),
            ("q17b: Sexual behaviour, staff", 0.849),
        ],
        "FlexWorking": [
            ("q6c: Good work-life balance", 0.972),
            ("q4d: Flex working satisfaction", 0.927),
        ],
        "LineMgmt": [
            ("q9a-i: Manager quality (composite)", 0.886),
            ("q24e: Access to L&D", 0.875),
            ("q23a: Had appraisal", 0.574),
            ("q23b: Appraisal improved job", 0.276),
        ],
        "SupportEnv": [
            ("q11a: Org acts on wellbeing", 0.925),
            ("q8b-d: Interpersonal culture", 0.831),
            ("q3h: Adequate materials", 0.804),
            ("q3i: Enough staff", 0.699),
        ],
    }

    fig, axes = plt.subplots(2, 3, figsize=(14, 8))
    axes = axes.flatten()

    display_names = {
        "Violence": "Reducing Violence",
        "Racism": "Tackling Racism",
        "SexSafety": "Sexual Safety",
        "FlexWorking": "Flexible Working",
        "LineMgmt": "Line Management",
        "SupportEnv": "Supportive Env.",
    }

    for idx, (std, items) in enumerate(data.items()):
        ax = axes[idx]
        labels = [it[0] for it in items]
        loadings = [it[1] for it in items]

        y_pos = range(len(items))
        colour = STD_COLOURS[std]

        bars = ax.barh(y_pos, loadings, color=colour, edgecolor="white",
                       height=0.6, alpha=0.85)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(labels, fontsize=8)
        ax.set_xlim(0, 1.15)
        ax.axvline(x=0.40, color=GOVUK_MID_GREY, linewidth=0.8, linestyle=":", alpha=0.5)
        ax.set_xlabel("Standardised loading", fontsize=8)
        ax.set_title(display_names[std], fontsize=10, fontweight="bold", color=colour)
        ax.invert_yaxis()

        for bar, val in zip(bars, loadings):
            ax.text(val + 0.02, bar.get_y() + bar.get_height()/2,
                    f"{val:.3f}", va="center", fontsize=9, fontweight="bold")

    fig.suptitle("CFA Standardised Loadings by Standard",
                 fontsize=13, fontweight="bold", y=1.01)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "report_loadings.png"), dpi=200, bbox_inches="tight")
    plt.close()
    print("  Saved: report_loadings.png")


def chart_empirical_structure():
    """Diagram showing theorised 6 vs empirical 3 structure."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Left: Theorised 6 standards
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 10)
    ax1.set_aspect("equal")
    ax1.axis("off")
    ax1.set_title("Theorised Structure\n(6 independent standards)", fontsize=11,
                  fontweight="bold", pad=15)

    std_names = ["Reducing\nViolence", "Tackling\nRacism", "Sexual\nSafety",
                 "Flexible\nWorking", "Line\nMgmt", "Supportive\nEnv"]
    std_keys = ["Violence", "Racism", "SexSafety", "FlexWorking", "LineMgmt", "SupportEnv"]
    positions = [(2, 7.5), (5, 7.5), (8, 7.5), (2, 3.5), (5, 3.5), (8, 3.5)]

    for (x, y), name, key in zip(positions, std_names, std_keys):
        circle = plt.Circle((x, y), 1.2, color=STD_COLOURS[key], alpha=0.3)
        ax1.add_patch(circle)
        circle_border = plt.Circle((x, y), 1.2, fill=False, edgecolor=STD_COLOURS[key], linewidth=2)
        ax1.add_patch(circle_border)
        ax1.text(x, y, name, ha="center", va="center", fontsize=8, fontweight="bold")

    # Right: Empirical 3 clusters
    ax2.set_xlim(0, 10)
    ax2.set_ylim(0, 10)
    ax2.set_aspect("equal")
    ax2.axis("off")
    ax2.set_title("Empirical Structure\n(3 clusters from factor analysis)", fontsize=11,
                  fontweight="bold", pad=15)

    # Cluster 1: Staff Safety (Violence + SexSafety)
    ellipse1 = mpatches.FancyBboxPatch((0.3, 6), 4, 3.2, boxstyle="round,pad=0.3",
                                        facecolor=GOVUK_RED, alpha=0.15, edgecolor=GOVUK_RED, linewidth=2)
    ax2.add_patch(ellipse1)
    ax2.text(2.3, 8.5, "Staff Safety", fontsize=10, fontweight="bold", color=GOVUK_RED, ha="center")

    c1a = plt.Circle((1.5, 7.2), 0.8, color=STD_COLOURS["Violence"], alpha=0.4)
    ax2.add_patch(c1a)
    ax2.text(1.5, 7.2, "Violence", ha="center", va="center", fontsize=8, fontweight="bold")

    c1b = plt.Circle((3.2, 7.2), 0.8, color=STD_COLOURS["SexSafety"], alpha=0.4)
    ax2.add_patch(c1b)
    ax2.text(3.2, 7.2, "Sexual\nSafety", ha="center", va="center", fontsize=7, fontweight="bold")

    ax2.text(2.3, 6.2, "r = 0.96", ha="center", fontsize=8, style="italic", color=GOVUK_DARK_GREY)

    # Cluster 2: Management & Culture (Flex + LineMgmt + SupportEnv)
    ellipse2 = mpatches.FancyBboxPatch((4.8, 1.5), 4.8, 5, boxstyle="round,pad=0.3",
                                        facecolor=GOVUK_BLUE, alpha=0.1, edgecolor=GOVUK_BLUE, linewidth=2)
    ax2.add_patch(ellipse2)
    ax2.text(7.2, 6, "Management\n& Culture", fontsize=10, fontweight="bold", color=DHSC_DEEPEST, ha="center")

    c2a = plt.Circle((6.2, 4.5), 0.8, color=STD_COLOURS["FlexWorking"], alpha=0.4)
    ax2.add_patch(c2a)
    ax2.text(6.2, 4.5, "Flex\nWorking", ha="center", va="center", fontsize=7, fontweight="bold")

    c2b = plt.Circle((8.2, 4.5), 0.8, color=STD_COLOURS["LineMgmt"], alpha=0.4)
    ax2.add_patch(c2b)
    ax2.text(8.2, 4.5, "Line\nMgmt", ha="center", va="center", fontsize=7, fontweight="bold")

    c2c = plt.Circle((7.2, 2.7), 0.8, color=STD_COLOURS["SupportEnv"], alpha=0.4)
    ax2.add_patch(c2c)
    ax2.text(7.2, 2.7, "Support\nEnv", ha="center", va="center", fontsize=7, fontweight="bold")

    ax2.text(7.2, 1.8, "r = 0.91-0.99", ha="center", fontsize=8, style="italic", color=GOVUK_DARK_GREY)

    # Cluster 3: Fairness (Racism - standalone)
    ellipse3 = mpatches.FancyBboxPatch((0.3, 1), 3.5, 3.5, boxstyle="round,pad=0.3",
                                        facecolor=GOVUK_ORANGE, alpha=0.15,
                                        edgecolor=GOVUK_ORANGE, linewidth=2)
    ax2.add_patch(ellipse3)
    ax2.text(2, 3.8, "Organisational\nFairness", fontsize=10, fontweight="bold",
             color="#B04B00", ha="center")

    c3 = plt.Circle((2, 2.5), 0.8, color=STD_COLOURS["Racism"], alpha=0.4)
    ax2.add_patch(c3)
    ax2.text(2, 2.5, "Racism", ha="center", va="center", fontsize=8, fontweight="bold")

    ax2.text(2, 1.3, "most distinct", ha="center", fontsize=8, style="italic", color=GOVUK_DARK_GREY)

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "report_structure_comparison.png"), dpi=200, bbox_inches="tight")
    plt.close()
    print("  Saved: report_structure_comparison.png")


if __name__ == "__main__":
    print("Generating report charts...")
    chart_factor_correlations()
    chart_variance_decomposition()
    chart_trust_profiles()
    chart_cfa_comparison()
    chart_standardised_loadings()
    chart_empirical_structure()
    print("Done.")
