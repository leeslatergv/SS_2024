# NHS Staff Standards - CFA v2 (lavaan with relaxed convergence)
# The 6-factor model has near-singular covariance matrix because FlexWorking,
# LineMgmt, and SupportEnv correlate at r=0.97-0.99. This IS the finding.

library(lavaan)

df <- read.csv("~/nhs-staff-standards/data/expanded_trust_means.csv")

# Create composites
mgr_items <- c("mgr_encourages", "mgr_clear_feedback", "mgr_asks_opinion",
               "mgr_interest_wellbeing", "mgr_values_work", "mgr_understanding_problems",
               "mgr_listens", "mgr_cares_concerns", "mgr_effective_action")
df$mgr_composite <- rowMeans(df[, mgr_items], na.rm = TRUE)
people_items <- c("people_kind_to_each_other", "people_polite_respectful", "people_show_appreciation")
df$people_culture <- rowMeans(df[, people_items], na.rm = TRUE)

items <- c(
  "violence_patients", "harassment_patients", "harassment_managers", "harassment_colleagues",
  "fair_career_progression", "discrim_patients", "discrim_colleagues", "org_respects_differences",
  "sexual_behaviour_patients", "sexual_behaviour_staff",
  "satisfaction_flexible_working", "good_wl_balance",
  "mgr_composite", "had_appraisal", "appraisal_improved_job", "access_learning_dev",
  "adequate_materials_equipment", "enough_staff", "people_culture", "org_positive_wellbeing"
)

dat <- df[complete.cases(df[, items]), items]
cat(sprintf("N = %d, p = %d, N:p = %.1f:1\n\n", nrow(dat), ncol(dat), nrow(dat)/ncol(dat)))

# ── Model A: 1-factor ──────────────────────────────────────────────
mA <- paste0("General =~ ", paste(items, collapse = " + "))

# ── Model B: 3-factor (maximally parsimonious) ────────────────────
mB <- '
  GoodManagement =~ mgr_composite + satisfaction_flexible_working + good_wl_balance +
                     people_culture + adequate_materials_equipment + enough_staff +
                     org_positive_wellbeing + had_appraisal + appraisal_improved_job + access_learning_dev
  PatientHarm =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues +
                 sexual_behaviour_patients + sexual_behaviour_staff
  Fairness =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
'

# ── Model C: 4-factor (EFA-driven: splits appraisal out) ──────────
mC <- '
  ManagerTeam =~ mgr_composite + satisfaction_flexible_working + good_wl_balance +
                 people_culture + adequate_materials_equipment + enough_staff +
                 org_positive_wellbeing + access_learning_dev
  PatientHarm =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues +
                 sexual_behaviour_patients + sexual_behaviour_staff
  Fairness =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
  Appraisal =~ had_appraisal + appraisal_improved_job + access_learning_dev
'

# ── Model D: 6-factor correlated (theory, relaxed convergence) ────
mD <- '
  Violence =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues
  Racism =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
  SexSafety =~ sexual_behaviour_patients + sexual_behaviour_staff
  FlexWorking =~ satisfaction_flexible_working + good_wl_balance
  LineMgmt =~ mgr_composite + had_appraisal + appraisal_improved_job + access_learning_dev
  SupportEnv =~ adequate_materials_equipment + enough_staff + people_culture + org_positive_wellbeing
'

# ── Fit all ──────────────────────────────────────────────────────────
cat(strrep("=", 90), "\n")
cat("  CFA MODEL COMPARISON (lavaan, ML estimator)\n")
cat(strrep("=", 90), "\n\n")

models <- list(
  "A. 1-factor" = mA,
  "B. 3-factor (merged)" = mB,
  "C. 4-factor (EFA-split)" = mC,
  "D. 6-factor (theory)" = mD
)

fits <- list()
for (label in names(models)) {
  cat(sprintf("--- %s ---\n", label))
  tryCatch({
    # Use check.gradient = FALSE for the 6-factor model
    fit <- cfa(models[[label]], data = dat, estimator = "ML",
               check.gradient = FALSE, optim.force.converged = TRUE)
    fits[[label]] <- fit

    fi <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea",
                              "rmsea.ci.lower", "rmsea.ci.upper", "srmr", "aic", "bic"))
    cat(sprintf("  chi2 = %.1f, df = %.0f, p = %.4f\n", fi["chisq"], fi["df"], fi["pvalue"]))
    cat(sprintf("  CFI  = %.3f, TLI = %.3f\n", fi["cfi"], fi["tli"]))
    cat(sprintf("  RMSEA= %.3f [%.3f, %.3f]\n", fi["rmsea"], fi["rmsea.ci.lower"], fi["rmsea.ci.upper"]))
    cat(sprintf("  SRMR = %.3f\n", fi["srmr"]))
    cat(sprintf("  AIC  = %.1f, BIC = %.1f\n\n", fi["aic"], fi["bic"]))
  }, error = function(e) {
    cat(sprintf("  FAILED: %s\n\n", conditionMessage(e)))
  })
}

# Comparison table
cat("\n")
cat(strrep("=", 110), "\n")
cat(sprintf("  %-28s %8s %5s %7s %7s %7s %7s %7s\n",
            "Model", "chi2", "df", "chi/df", "CFI", "TLI", "RMSEA", "SRMR"))
cat(strrep("-", 110), "\n")
for (label in names(fits)) {
  fi <- fitMeasures(fits[[label]], c("chisq", "df", "cfi", "tli", "rmsea", "srmr"))
  cat(sprintf("  %-28s %8.1f %5.0f %7.2f %7.3f %7.3f %7.3f %7.3f\n",
              label, fi["chisq"], fi["df"], fi["chisq"]/fi["df"],
              fi["cfi"], fi["tli"], fi["rmsea"], fi["srmr"]))
}

# ── 6-factor results (if converged) ────────────────────────────────
if ("D. 6-factor (theory)" %in% names(fits)) {
  fit6 <- fits[["D. 6-factor (theory)"]]

  cat("\n\n")
  cat(strrep("=", 80), "\n")
  cat("  6-FACTOR CFA: LATENT FACTOR CORRELATIONS\n")
  cat(strrep("=", 80), "\n\n")

  tryCatch({
    corr_lv <- lavInspect(fit6, "cor.lv")
    print(round(corr_lv, 3))

    cat("\n  Interpretation:\n")
    fnames <- rownames(corr_lv)
    for (i in 1:(length(fnames)-1)) {
      for (j in (i+1):length(fnames)) {
        r <- corr_lv[i,j]
        label <- if (abs(r) > 0.90) "REDUNDANT" else if (abs(r) > 0.80) "HIGH" else if (abs(r) > 0.60) "MODERATE" else "DISTINCT"
        cat(sprintf("    %-12s <-> %-12s: r=%+.3f  [%s]\n", fnames[i], fnames[j], r, label))
      }
    }
  }, error = function(e) {
    cat(sprintf("  Could not extract: %s\n", conditionMessage(e)))
  })

  cat("\n\n")
  cat(strrep("=", 80), "\n")
  cat("  6-FACTOR CFA: STANDARDISED LOADINGS\n")
  cat(strrep("=", 80), "\n\n")

  tryCatch({
    std_sol <- standardizedSolution(fit6)
    loadings <- std_sol[std_sol$op == "=~", ]
    cat(sprintf("  %-12s %-35s %8s %10s\n", "Factor", "Item", "Std.Beta", "p-value"))
    cat(strrep("-", 72), "\n")
    for (i in 1:nrow(loadings)) {
      cat(sprintf("  %-12s %-35s %8.3f %10.4f\n",
                  loadings$lhs[i], loadings$rhs[i],
                  loadings$est.std[i], loadings$pvalue[i]))
    }
  }, error = function(e) {
    cat(sprintf("  Could not extract: %s\n", conditionMessage(e)))
  })
}

# ── Also show 4-factor loadings ────────────────────────────────────
if ("C. 4-factor (EFA-split)" %in% names(fits)) {
  fit4 <- fits[["C. 4-factor (EFA-split)"]]

  cat("\n\n")
  cat(strrep("=", 80), "\n")
  cat("  4-FACTOR CFA: FACTOR CORRELATIONS\n")
  cat(strrep("=", 80), "\n\n")

  tryCatch({
    corr_lv4 <- lavInspect(fit4, "cor.lv")
    print(round(corr_lv4, 3))
  }, error = function(e) {
    cat(sprintf("  Could not extract: %s\n", conditionMessage(e)))
  })
}

# ── Chi-square difference tests (nested models) ───────────────────
cat("\n\n")
cat(strrep("=", 80), "\n")
cat("  NESTED MODEL COMPARISONS\n")
cat(strrep("=", 80), "\n")

if (all(c("A. 1-factor", "B. 3-factor (merged)") %in% names(fits))) {
  cat("\n  3-factor vs 1-factor:\n")
  tryCatch(print(anova(fits[["A. 1-factor"]], fits[["B. 3-factor (merged)"]])), error = function(e) cat("  error\n"))
}
if (all(c("B. 3-factor (merged)", "C. 4-factor (EFA-split)") %in% names(fits))) {
  cat("\n  4-factor vs 3-factor:\n")
  tryCatch(print(anova(fits[["B. 3-factor (merged)"]], fits[["C. 4-factor (EFA-split)"]])), error = function(e) cat("  error\n"))
}
