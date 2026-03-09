"""
Q6b vs Q6c: Testing question strength for the Flexible Working theme.

Policy want Q6b ("My organisation is committed to helping me balance my work
and home life") in the Promoting Flexible Working standard.
Paper recommended Q6c ("I achieve a good balance between my work life and my
home life").

This script tests both using the SAME methodology as expanded_fa.py:
  - EFA: minres extraction, promax rotation
  - Same expanded variable pool
  - Trust-level means (expanded_trust_means.csv)

Outputs:
  1. EFA loadings for Q6b and Q6c side by side
  2. Cronbach's alpha for FlexWorking theme with Q6b vs Q6c
  3. Corrected item-total correlations
  4. CFA-style fit comparison (single-factor congeneric model for FlexWorking)
"""

import pandas as pd
import numpy as np
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_kmo
from scipy import stats
import warnings
warnings.filterwarnings("ignore")

DATA_FILE = "expanded_trust_means.csv"

# ── Same reduced variable set used in the CFA scripts ──────────────
# (compositing the q9 items into mgr_composite, q8b/c/d into people_culture,
#  matching cfa_lavaan.R / cfa_lavaan_v2.R)
MGR_ITEMS = [
    "mgr_encourages", "mgr_clear_feedback", "mgr_asks_opinion",
    "mgr_interest_wellbeing", "mgr_values_work", "mgr_understanding_problems",
    "mgr_listens", "mgr_cares_concerns", "mgr_effective_action",
]
PEOPLE_ITEMS = [
    "people_kind_to_each_other", "people_polite_respectful", "people_show_appreciation",
]


def load_data():
    df = pd.read_csv(DATA_FILE)
    # Create composites (same as CFA scripts)
    df["mgr_composite"] = df[MGR_ITEMS].mean(axis=1)
    df["people_culture"] = df[PEOPLE_ITEMS].mean(axis=1)
    return df


def run_efa_comparison(df):
    """Run the same EFA as expanded_fa.py but on the CFA variable set,
    once with Q6c and once with Q6b, to compare loadings."""

    print("=" * 80)
    print("  PART 1: EFA LOADINGS — Q6b vs Q6c")
    print("  (minres extraction, promax rotation — same as expanded_fa.py)")
    print("=" * 80)

    # Base items (same as cfa_lavaan_v2.R 6-factor model, minus the flex item)
    base_items = [
        "violence_patients", "harassment_patients", "harassment_managers",
        "harassment_colleagues",
        "fair_career_progression", "discrim_patients", "discrim_colleagues",
        "org_respects_differences",
        "sexual_behaviour_patients", "sexual_behaviour_staff",
        "satisfaction_flexible_working",
        # Q6c or Q6b goes here
        "mgr_composite", "had_appraisal", "appraisal_improved_job",
        "access_learning_dev",
        "adequate_materials_equipment", "enough_staff", "people_culture",
        "org_positive_wellbeing",
    ]

    for label, flex_item in [("Q6c (good_wl_balance)", "good_wl_balance"),
                              ("Q6b (org_committed_wlb)", "org_committed_wlb")]:
        items = base_items.copy()
        # Insert the flex item after satisfaction_flexible_working
        idx = items.index("satisfaction_flexible_working") + 1
        items.insert(idx, flex_item)

        dat = df[items].dropna()
        n_obs = len(dat)

        print(f"\n  --- EFA with {label} (N={n_obs}, {len(items)} items, 6 factors) ---")

        fa = FactorAnalyzer(n_factors=6, rotation="promax", method="minres")
        fa.fit(dat)

        loadings = pd.DataFrame(
            fa.loadings_, index=items,
            columns=[f"F{i+1}" for i in range(6)]
        )

        # Variance explained
        _, _, cum_var = fa.get_factor_variance()
        print(f"  Total variance explained: {cum_var[-1]:.1%}")

        # Print loadings (highlight > 0.30)
        print(f"\n  {'Item':40s}", end="")
        for c in loadings.columns:
            print(f" {c:>7s}", end="")
        print(f"  {'Primary':>8s}")
        print("  " + "-" * 95)

        for item, row in loadings.iterrows():
            primary = row.abs().idxmax()
            primary_val = row[primary]
            marker = " ***" if item in [flex_item, "satisfaction_flexible_working"] else ""
            print(f"  {item:40s}", end="")
            for v in row:
                if abs(v) >= 0.30:
                    print(f" {v:7.3f}", end="")
                else:
                    print(f"    .   ", end="")
            print(f"  {primary:>4s} ({primary_val:+.3f}){marker}")

        # Factor correlations
        if fa.phi_ is not None:
            phi = pd.DataFrame(
                fa.phi_,
                index=[f"F{i+1}" for i in range(6)],
                columns=[f"F{i+1}" for i in range(6)]
            )
            print(f"\n  Factor correlations:")
            print("  " + phi.round(3).to_string().replace("\n", "\n  "))

    # Also run the FULL expanded EFA (all items including BOTH Q6b and Q6c)
    print(f"\n\n  --- FULL EFA with BOTH Q6b AND Q6c (to see where each lands) ---")
    full_items = base_items.copy()
    idx = full_items.index("satisfaction_flexible_working") + 1
    full_items.insert(idx, "good_wl_balance")
    full_items.insert(idx + 1, "org_committed_wlb")
    # Also add manager_flexible_working (Q6d) for completeness
    if "manager_flexible_working" in df.columns:
        full_items.insert(idx + 2, "manager_flexible_working")

    dat_full = df[full_items].dropna()
    print(f"  N={len(dat_full)}, {len(full_items)} items, 6 factors")

    fa_full = FactorAnalyzer(n_factors=6, rotation="promax", method="minres")
    fa_full.fit(dat_full)

    loadings_full = pd.DataFrame(
        fa_full.loadings_, index=full_items,
        columns=[f"F{i+1}" for i in range(6)]
    )

    flex_questions = ["satisfaction_flexible_working", "good_wl_balance",
                      "org_committed_wlb", "manager_flexible_working"]

    print(f"\n  Loadings for flexible working items only:")
    print(f"  {'Item':40s}", end="")
    for c in loadings_full.columns:
        print(f" {c:>7s}", end="")
    print(f"  {'Primary':>8s}")
    print("  " + "-" * 95)

    for item in flex_questions:
        if item in loadings_full.index:
            row = loadings_full.loc[item]
            primary = row.abs().idxmax()
            primary_val = row[primary]
            print(f"  {item:40s}", end="")
            for v in row:
                if abs(v) >= 0.30:
                    print(f" {v:7.3f}", end="")
                else:
                    print(f"    .   ", end="")
            print(f"  {primary:>4s} ({primary_val:+.3f})")

    return loadings_full


def cronbach_alpha(items_df):
    """Compute Cronbach's alpha for a set of items."""
    k = items_df.shape[1]
    if k < 2:
        return np.nan
    item_vars = items_df.var(axis=0, ddof=1)
    total_var = items_df.sum(axis=1).var(ddof=1)
    alpha = (k / (k - 1)) * (1 - item_vars.sum() / total_var)
    return alpha


def item_total_correlations(items_df):
    """Corrected item-total correlation for each item."""
    total = items_df.sum(axis=1)
    results = {}
    for col in items_df.columns:
        corrected_total = total - items_df[col]
        r, p = stats.pearsonr(items_df[col], corrected_total)
        results[col] = r
    return results


def run_reliability_comparison(df):
    """Compare Cronbach's alpha and item-total correlations for
    FlexWorking theme with Q6b vs Q6c."""

    print("\n\n" + "=" * 80)
    print("  PART 2: RELIABILITY ANALYSIS — FlexWorking with Q6b vs Q6c")
    print("=" * 80)

    # The FlexWorking factor in the CFA had: satisfaction_flexible_working + one of Q6b/Q6c
    # But the full theoretical assignment from expanded_fa.py also included Q6d
    # Test both the minimal (2-item) and full (3-item) versions

    configs = {
        "2-item with Q6c": ["satisfaction_flexible_working", "good_wl_balance"],
        "2-item with Q6b": ["satisfaction_flexible_working", "org_committed_wlb"],
        "3-item with Q6c": ["satisfaction_flexible_working", "good_wl_balance",
                            "manager_flexible_working"],
        "3-item with Q6b": ["satisfaction_flexible_working", "org_committed_wlb",
                            "manager_flexible_working"],
        "All 4 flex items": ["satisfaction_flexible_working", "good_wl_balance",
                             "org_committed_wlb", "manager_flexible_working"],
    }

    for label, items in configs.items():
        available = [i for i in items if i in df.columns]
        dat = df[available].dropna()
        n = len(dat)

        alpha = cronbach_alpha(dat)
        itc = item_total_correlations(dat)

        print(f"\n  --- {label} (N={n}) ---")
        print(f"  Cronbach's alpha: {alpha:.3f}")
        print(f"  Inter-item correlation (mean): {dat.corr().where(np.triu(np.ones(dat.corr().shape), k=1).astype(bool)).stack().mean():.3f}")
        print(f"  Corrected item-total correlations:")
        for item, r in itc.items():
            print(f"    {item:40s}  r = {r:.3f}")


def run_cfa_like_comparison(df):
    """Compare single-factor congeneric models for FlexWorking with Q6b vs Q6c.
    Since semopy won't install, use a maximum-likelihood factor analysis
    (1-factor CFA is equivalent to 1-factor FA with ML estimation)."""

    print("\n\n" + "=" * 80)
    print("  PART 3: SINGLE-FACTOR MODEL FIT — FlexWorking with Q6b vs Q6c")
    print("  (1-factor ML extraction = congeneric measurement model)")
    print("=" * 80)

    configs = {
        "FlexWorking with Q6c (paper)": [
            "satisfaction_flexible_working", "good_wl_balance",
            "manager_flexible_working"],
        "FlexWorking with Q6b (policy)": [
            "satisfaction_flexible_working", "org_committed_wlb",
            "manager_flexible_working"],
        "FlexWorking with both Q6b + Q6c": [
            "satisfaction_flexible_working", "good_wl_balance",
            "org_committed_wlb", "manager_flexible_working"],
    }

    for label, items in configs.items():
        available = [i for i in items if i in df.columns]
        dat = df[available].dropna()
        n = len(dat)
        p = len(available)

        print(f"\n  --- {label} (N={n}, {p} items) ---")

        # 1-factor ML
        try:
            fa = FactorAnalyzer(n_factors=1, rotation=None, method="ml")
            fa.fit(dat)

            loadings = pd.Series(fa.loadings_.flatten(), index=available)
            communalities = pd.Series(fa.get_communalities(), index=available)
            ev, _ = fa.get_eigenvalues()
            _, prop_var, cum_var = fa.get_factor_variance()

            print(f"  Variance explained: {cum_var[0]:.1%}")
            print(f"  Eigenvalue: {ev[0]:.3f}")
            print(f"\n  {'Item':40s} {'Loading':>8s} {'Communality':>12s}")
            print("  " + "-" * 62)
            for item in available:
                print(f"  {item:40s} {loadings[item]:8.3f} {communalities[item]:12.3f}")

            # Goodness of fit (chi-square test from ML factor analysis)
            # factor_analyzer doesn't expose this directly, so compute from
            # the reproduced correlation matrix
            R = dat.corr().values
            L = fa.loadings_
            psi = np.diag(fa.get_uniquenesses())
            R_hat = L @ L.T + psi

            # Residual matrix
            residuals = R - R_hat
            off_diag = residuals[np.triu_indices_from(residuals, k=1)]
            rmsr = np.sqrt(np.mean(off_diag ** 2))

            print(f"\n  RMSR (root mean square residual): {rmsr:.4f}")
            print(f"  Max absolute residual: {np.max(np.abs(off_diag)):.4f}")

        except Exception as e:
            print(f"  FAILED: {e}")


def run_correlation_comparison(df):
    """Show how Q6b and Q6c correlate with all other items in the model."""

    print("\n\n" + "=" * 80)
    print("  PART 4: CORRELATION WITH OTHER STANDARDS' ITEMS")
    print("  (Does Q6b or Q6c discriminate better from other themes?)")
    print("=" * 80)

    other_items = {
        "Violence": ["violence_patients", "harassment_patients"],
        "Racism": ["fair_career_progression", "org_respects_differences"],
        "SexSafety": ["sexual_behaviour_patients", "sexual_behaviour_staff"],
        "LineMgmt": ["mgr_composite", "had_appraisal", "appraisal_improved_job"],
        "SupportEnv": ["adequate_materials_equipment", "enough_staff",
                        "people_culture", "org_positive_wellbeing"],
    }

    # Create composites
    df_work = df.copy()
    df_work["mgr_composite"] = df[MGR_ITEMS].mean(axis=1)
    df_work["people_culture"] = df[PEOPLE_ITEMS].mean(axis=1)

    flex_items = ["satisfaction_flexible_working", "good_wl_balance",
                  "org_committed_wlb", "manager_flexible_working"]

    # Compute flex composites
    df_work["flex_with_q6c"] = df_work[["satisfaction_flexible_working",
                                         "good_wl_balance"]].mean(axis=1)
    df_work["flex_with_q6b"] = df_work[["satisfaction_flexible_working",
                                         "org_committed_wlb"]].mean(axis=1)

    print(f"\n  Correlations of FlexWorking composite (with Q6c vs Q6b) with other standards:")
    print(f"\n  {'Standard':15s} {'Item':40s} {'r(Q6c ver)':>11s} {'r(Q6b ver)':>11s} {'Diff':>7s}")
    print("  " + "-" * 88)

    for std, items in other_items.items():
        for item in items:
            if item in df_work.columns:
                valid = df_work[[item, "flex_with_q6c", "flex_with_q6b"]].dropna()
                r_c = valid[item].corr(valid["flex_with_q6c"])
                r_b = valid[item].corr(valid["flex_with_q6b"])
                diff = abs(r_b) - abs(r_c)
                print(f"  {std:15s} {item:40s} {r_c:+11.3f} {r_b:+11.3f} {diff:+7.3f}")

    # Direct comparison: Q6b vs Q6c correlation with each other and with Q4d, Q6d
    print(f"\n\n  Inter-correlations among flexible working items:")
    corr = df_work[flex_items].corr()
    print("  " + corr.round(3).to_string().replace("\n", "\n  "))

    # Key comparison
    r_bc = df_work["good_wl_balance"].corr(df_work["org_committed_wlb"])
    print(f"\n  Q6b <-> Q6c correlation: r = {r_bc:.3f}")
    print(f"  (High correlation means they're measuring similar things;")
    print(f"   the question is which loads more cleanly on FlexWorking vs other factors)")


def main():
    print("=" * 80)
    print("  Q6b vs Q6c: WHICH FITS BETTER IN THE FLEXIBLE WORKING STANDARD?")
    print("  Q6b: 'My org is committed to helping me balance work and home life'")
    print("  Q6c: 'I achieve a good balance between my work life and my home life'")
    print("=" * 80)

    df = load_data()
    print(f"\n  Data: {len(df)} trusts, {df.shape[1]} variables")

    run_efa_comparison(df)
    run_reliability_comparison(df)
    run_cfa_like_comparison(df)
    run_correlation_comparison(df)

    # Summary
    print("\n\n" + "=" * 80)
    print("  SUMMARY")
    print("=" * 80)
    print("""
  Compare across the four tests above:
  1. EFA LOADINGS: Which item loads more strongly/cleanly on the FlexWorking factor?
  2. RELIABILITY: Which gives better Cronbach's alpha and item-total correlations?
  3. MODEL FIT: Which gives better fit in a single-factor measurement model?
  4. DISCRIMINANT VALIDITY: Which correlates less with OTHER standards' items?
     (Lower cross-correlations = better discriminant validity = more distinct theme)
""")


if __name__ == "__main__":
    main()
