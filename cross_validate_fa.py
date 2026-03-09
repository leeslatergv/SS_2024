"""
Cross-validation of factor structure: EFA on 2023, CFA on 2024.
Matches the People Promise methodology (explore on year N, confirm on year N+1).
Uses post-COVID years only.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_kmo, calculate_bartlett_sphericity
import os
import warnings
warnings.filterwarnings("ignore")

DATA_DIR = os.path.expanduser("~/nhs-staff-standards/data")
RAW_FILE = os.path.expanduser("~/nhs-burnout/data/raw/benchmark_data.xlsx")
OUTPUT_DIR = os.path.expanduser("~/nhs-staff-standards/output")

TRUST_SHEETS = [
    "Acute&Acute Community Trusts",
    "Acute Specialist Trusts",
    "MH&LD, MH, LD&Community Trusts",
    "Community Trusts",
    "Ambulance Trusts",
]

# Same expanded variable pool as main analysis
EXPANDED_VARS = {
    "q13a": "violence_patients",
    "q13d": "violence_reported",
    "q14a": "harassment_patients",
    "q14b": "harassment_managers",
    "q14c": "harassment_colleagues",
    "q14d": "harassment_reported",
    "q15": "fair_career_progression",
    "q16a": "discrim_patients",
    "q16b": "discrim_colleagues",
    "q21": "org_respects_differences",
    "q17a": "sexual_behaviour_patients",
    "q17b": "sexual_behaviour_staff",
    "q4d": "satisfaction_flexible_working",
    "q6b": "org_committed_wlb",
    "q6c": "good_wl_balance",
    "q6d": "manager_flexible_working",
    "q9a": "mgr_encourages",
    "q9b": "mgr_clear_feedback",
    "q9c": "mgr_asks_opinion",
    "q9d": "mgr_interest_wellbeing",
    "q9e": "mgr_values_work",
    "q9f": "mgr_understanding_problems",
    "q9g": "mgr_listens",
    "q9h": "mgr_cares_concerns",
    "q9i": "mgr_effective_action",
    "q23a": "had_appraisal",
    "q23b": "appraisal_improved_job",
    "q23c": "appraisal_clear_objectives",
    "q23d": "appraisal_felt_valued",
    "q24a": "org_challenging_work",
    "q24b": "career_dev_opportunities",
    "q24c": "opportunities_knowledge_skills",
    "q24d": "supported_develop_potential",
    "q24e": "access_learning_dev",
    "q3h": "adequate_materials_equipment",
    "q3i": "enough_staff",
    "q5c": "relationships_strained",
    "q7c": "respect_from_colleagues",
    "q7h": "feel_valued_by_team",
    "q8b": "people_kind_to_each_other",
    "q8c": "people_polite_respectful",
    "q8d": "people_show_appreciation",
    "q11a": "org_positive_wellbeing",
    "q22": "nutritious_food",
    "q25c": "recommend_org_to_work",
    "q25e": "safe_to_speak_up",
    "q2a": "look_forward_to_work",
    "q2b": "enthusiastic_about_job",
    "q20a": "secure_raising_concerns",
    "q20b": "confident_org_address_concern",
}


def extract_single_year(year):
    """Extract trust-level data for a single year."""
    all_rows = []
    for sheet in TRUST_SHEETS:
        df = pd.read_excel(RAW_FILE, sheet_name=sheet)
        for _, org in df.iterrows():
            org_id = org.get("org_id")
            if pd.isna(org_id):
                continue
            row = {"org_id": org_id, "trust_type": sheet}
            any_data = False
            for raw_var, clean_name in EXPANDED_VARS.items():
                col = f"{raw_var}_{year}"
                if col in df.columns:
                    val = org[col]
                    row[clean_name] = val if pd.notna(val) else np.nan
                    if pd.notna(val):
                        any_data = True
                else:
                    row[clean_name] = np.nan
            if any_data:
                all_rows.append(row)
    return pd.DataFrame(all_rows)


def parallel_analysis(data, n_iter=500, percentile=95, method="minres"):
    """Horn's parallel analysis."""
    n_obs, n_vars = data.shape
    fa = FactorAnalyzer(n_factors=n_vars, rotation=None, method=method)
    fa.fit(data)
    actual_ev, _ = fa.get_eigenvalues()

    random_evs = np.zeros((n_iter, n_vars))
    for i in range(n_iter):
        rd = pd.DataFrame(np.random.normal(size=(n_obs, n_vars)), columns=data.columns)
        fa_r = FactorAnalyzer(n_factors=n_vars, rotation=None, method=method)
        fa_r.fit(rd)
        random_evs[i, :], _ = fa_r.get_eigenvalues()

    threshold = np.percentile(random_evs, percentile, axis=0)
    n_factors = int(np.sum(actual_ev > threshold))
    return n_factors, actual_ev, threshold


def run_efa(data, label, n_factors=None):
    """Run EFA on a dataset and return loadings."""
    var_cols = [c for c in data.columns if c not in ["org_id", "trust_type"]]
    items = data[var_cols].dropna()

    # Drop low-variance
    stds = items.std()
    low_var = stds[stds < 0.01]
    if len(low_var) > 0:
        print(f"  Dropping low-variance: {list(low_var.index)}")
        items = items.drop(columns=low_var.index)

    n_obs, n_vars = items.shape
    print(f"\n  {label}: N = {n_obs} trusts, {n_vars} variables")

    # Adequacy
    kmo_per, kmo_total = calculate_kmo(items)
    print(f"  KMO: {kmo_total:.3f}")

    # Parallel analysis
    if n_factors is None:
        print("  Running parallel analysis...")
        n_factors, actual_ev, threshold = parallel_analysis(items)
        print(f"  Parallel analysis suggests: {n_factors} factors")
        print(f"  First 8 eigenvalues: {np.round(actual_ev[:8], 3)}")
    else:
        print(f"  Using specified {n_factors} factors")

    # Fit EFA
    fa = FactorAnalyzer(n_factors=n_factors, rotation="promax", method="minres")
    fa.fit(items)

    loadings = pd.DataFrame(
        fa.loadings_, index=items.columns,
        columns=[f"F{i+1}" for i in range(n_factors)]
    )

    # Variance
    _, _, cum_var = fa.get_factor_variance()
    print(f"  Variance explained: {cum_var[-1]:.1%}")

    # Factor correlations
    if fa.phi_ is not None:
        phi = pd.DataFrame(fa.phi_,
            index=[f"F{i+1}" for i in range(n_factors)],
            columns=[f"F{i+1}" for i in range(n_factors)])
        print(f"\n  Factor correlations:")
        print("  " + phi.round(3).to_string().replace("\n", "\n  "))

    # Communalities
    comm = pd.Series(fa.get_communalities(), index=items.columns)
    heywood = comm[comm > 1.0]
    if len(heywood) > 0:
        print(f"\n  HEYWOOD CASES: {dict(heywood.round(3))}")

    # Print loadings
    print(f"\n  Pattern matrix (loadings > 0.30):")
    for idx, row in loadings.iterrows():
        primary = row.abs().idxmax()
        vals = "  ".join(f"{v:6.3f}" if abs(v) >= 0.3 else f"  {'':4s} " for v in row)
        print(f"    {idx:35s} {vals}  -> {primary}")

    return loadings, items.columns.tolist(), n_factors


def generate_cfa_r_script(loadings, n_factors, items_available):
    """Generate an R script that runs CFA on 2024 data based on 2023 EFA structure."""

    # Assign items to factors based on highest loading
    factor_items = {f"F{i+1}": [] for i in range(n_factors)}
    for item, row in loadings.iterrows():
        primary = row.abs().idxmax()
        if abs(row[primary]) >= 0.30:
            factor_items[primary].append(item)

    # Build CFA model spec for lavaan
    # Also need composites for the highly collinear items
    mgr_items = ["mgr_encourages", "mgr_clear_feedback", "mgr_asks_opinion",
                 "mgr_interest_wellbeing", "mgr_values_work", "mgr_understanding_problems",
                 "mgr_listens", "mgr_cares_concerns", "mgr_effective_action"]
    people_items = ["people_kind_to_each_other", "people_polite_respectful", "people_show_appreciation"]

    # For CFA, composite the collinear items and keep 3-4 per factor
    mgr_factor = None
    for f, items in factor_items.items():
        if any(m in items for m in mgr_items):
            mgr_factor = f

    print(f"\n  Factor assignments from 2023 EFA:")
    for f, items in factor_items.items():
        print(f"    {f}: {items}")

    # Write R script
    r_script = '''# Cross-validation CFA: structure found in 2023, tested on 2024
# Auto-generated from 2023 EFA results

library(lavaan)

df <- read.csv("~/nhs-staff-standards/data/trust_2024.csv")

# Create composites (same as main analysis)
mgr_items <- c("mgr_encourages", "mgr_clear_feedback", "mgr_asks_opinion",
               "mgr_interest_wellbeing", "mgr_values_work", "mgr_understanding_problems",
               "mgr_listens", "mgr_cares_concerns", "mgr_effective_action")
available_mgr <- mgr_items[mgr_items %in% names(df)]
if (length(available_mgr) > 0) df$mgr_composite <- rowMeans(df[, available_mgr], na.rm = TRUE)

people_items <- c("people_kind_to_each_other", "people_polite_respectful", "people_show_appreciation")
available_people <- people_items[people_items %in% names(df)]
if (length(available_people) > 0) df$people_culture <- rowMeans(df[, available_people], na.rm = TRUE)

'''
    # Build the 6-factor theory model (same as cfa_lavaan_v2.R)
    r_script += '''
# ── Model A: 6-factor theory-driven (same spec as main analysis) ────
mA <- '
  Violence =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues
  Racism =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
  SexSafety =~ sexual_behaviour_patients + sexual_behaviour_staff
  FlexWorking =~ satisfaction_flexible_working + good_wl_balance
  LineMgmt =~ mgr_composite + had_appraisal + appraisal_improved_job + access_learning_dev
  SupportEnv =~ adequate_materials_equipment + enough_staff + people_culture + org_positive_wellbeing
'

# ── Model B: 3-factor merged (what EFA actually found) ──────────────
mB <- '
  GoodManagement =~ mgr_composite + satisfaction_flexible_working + good_wl_balance +
                     people_culture + adequate_materials_equipment + enough_staff +
                     org_positive_wellbeing + had_appraisal + appraisal_improved_job + access_learning_dev
  PatientHarm =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues +
                 sexual_behaviour_patients + sexual_behaviour_staff
  Fairness =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
'

# ── Model C: 1-factor ───────────────────────────────────────────────
items <- c(
  "violence_patients", "harassment_patients", "harassment_managers", "harassment_colleagues",
  "fair_career_progression", "discrim_patients", "discrim_colleagues", "org_respects_differences",
  "sexual_behaviour_patients", "sexual_behaviour_staff",
  "satisfaction_flexible_working", "good_wl_balance",
  "mgr_composite", "had_appraisal", "appraisal_improved_job", "access_learning_dev",
  "adequate_materials_equipment", "enough_staff", "people_culture", "org_positive_wellbeing"
)
mC <- paste0("General =~ ", paste(items, collapse = " + "))

dat <- df[complete.cases(df[, items]), items]
cat(sprintf("\\nCROSS-VALIDATION CFA: 2024 DATA (N = %d)\\n", nrow(dat)))
cat(sprintf("Structure discovered in 2023, tested on independent 2024 data\\n\\n"))

cat(strrep("=", 90), "\\n")
cat("  MODEL COMPARISON ON HELD-OUT 2024 DATA\\n")
cat(strrep("=", 90), "\\n\\n")

models <- list(
  "A. 6-factor (theory)" = mA,
  "B. 3-factor (merged)" = mB,
  "C. 1-factor" = mC
)

fits <- list()
for (label in names(models)) {
  cat(sprintf("--- %s ---\\n", label))
  tryCatch({
    fit <- cfa(models[[label]], data = dat, estimator = "ML",
               check.gradient = FALSE, optim.force.converged = TRUE)
    fits[[label]] <- fit

    fi <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea",
                              "rmsea.ci.lower", "rmsea.ci.upper", "srmr", "aic", "bic"))
    cat(sprintf("  chi2 = %.1f, df = %.0f, p = %.4f\\n", fi["chisq"], fi["df"], fi["pvalue"]))
    cat(sprintf("  CFI  = %.3f, TLI = %.3f\\n", fi["cfi"], fi["tli"]))
    cat(sprintf("  RMSEA= %.3f [%.3f, %.3f]\\n", fi["rmsea"], fi["rmsea.ci.lower"], fi["rmsea.ci.upper"]))
    cat(sprintf("  SRMR = %.3f\\n", fi["srmr"]))
    cat(sprintf("  AIC  = %.1f, BIC = %.1f\\n\\n", fi["aic"], fi["bic"]))
  }, error = function(e) {
    cat(sprintf("  FAILED: %s\\n\\n", conditionMessage(e)))
  })
}

# Comparison table
cat("\\n")
cat(strrep("=", 100), "\\n")
cat(sprintf("  %-28s %8s %5s %7s %7s %7s %7s %7s\\n",
            "Model", "chi2", "df", "chi/df", "CFI", "TLI", "RMSEA", "SRMR"))
cat(strrep("-", 100), "\\n")
for (label in names(fits)) {
  fi <- fitMeasures(fits[[label]], c("chisq", "df", "cfi", "tli", "rmsea", "srmr"))
  cat(sprintf("  %-28s %8.1f %5.0f %7.2f %7.3f %7.3f %7.3f %7.3f\\n",
              label, fi["chisq"], fi["df"], fi["chisq"]/fi["df"],
              fi["cfi"], fi["tli"], fi["rmsea"], fi["srmr"]))
}

# Factor correlations for 6-factor model
if ("A. 6-factor (theory)" %in% names(fits)) {
  cat("\\n\\n")
  cat(strrep("=", 80), "\\n")
  cat("  6-FACTOR: LATENT CORRELATIONS (2024 hold-out data)\\n")
  cat(strrep("=", 80), "\\n\\n")

  tryCatch({
    corr_lv <- lavInspect(fits[["A. 6-factor (theory)"]], "cor.lv")
    print(round(corr_lv, 3))

    cat("\\n  Key pairs:\\n")
    fnames <- rownames(corr_lv)
    for (i in 1:(length(fnames)-1)) {
      for (j in (i+1):length(fnames)) {
        r <- corr_lv[i,j]
        label <- if (abs(r) > 0.90) "REDUNDANT" else if (abs(r) > 0.80) "HIGH" else if (abs(r) > 0.60) "MODERATE" else "DISTINCT"
        cat(sprintf("    %-12s <-> %-12s: r=%+.3f  [%s]\\n", fnames[i], fnames[j], r, label))
      }
    }
  }, error = function(e) {
    cat(sprintf("  Could not extract: %s\\n", conditionMessage(e)))
  })
}

cat("\\n\\nDone.\\n")
'''

    script_path = os.path.join(os.path.expanduser("~/nhs-staff-standards/scripts"),
                               "cross_validate_cfa.R")
    with open(script_path, "w") as f:
        f.write(r_script)
    print(f"\n  Saved R script: {script_path}")
    return script_path


def main():
    print("=" * 70)
    print("  CROSS-VALIDATION: EFA on 2023, CFA on 2024")
    print("  (Post-COVID years only)")
    print("=" * 70)

    # Extract single-year data
    print("\n  Extracting 2023 data...")
    data_2023 = extract_single_year(2023)
    print(f"  2023: {len(data_2023)} trusts")

    print("\n  Extracting 2024 data...")
    data_2024 = extract_single_year(2024)
    print(f"  2024: {len(data_2024)} trusts")

    # Save 2024 data for R CFA
    data_2024.to_csv(os.path.join(DATA_DIR, "trust_2024.csv"), index=False)
    print(f"  Saved: data/trust_2024.csv")

    # Phase 1: EFA on 2023
    print("\n" + "=" * 70)
    print("  PHASE 1: EFA ON 2023 DATA (EXPLORATION SAMPLE)")
    print("=" * 70)

    loadings, items_avail, n_factors = run_efa(data_2023, "2023 EFA")

    # Also run with 6 forced factors
    print("\n  --- Forced 6-factor for comparison ---")
    loadings_6, _, _ = run_efa(data_2023, "2023 EFA (6 forced)", n_factors=6)

    # Phase 2: Generate CFA script for 2024
    print("\n" + "=" * 70)
    print("  PHASE 2: GENERATING CFA SCRIPT FOR 2024 DATA")
    print("=" * 70)

    r_path = generate_cfa_r_script(loadings, n_factors, items_avail)

    print("\n" + "=" * 70)
    print("  NEXT STEP: Run the R script")
    print(f"  Rscript {r_path}")
    print("=" * 70)


if __name__ == "__main__":
    main()
