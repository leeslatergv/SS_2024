# Cross-validation CFA: structure found in 2023, tested on 2024
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
cat(sprintf("\nCROSS-VALIDATION CFA: 2024 DATA (N = %d)\n", nrow(dat)))
cat(sprintf("Structure discovered in 2023, tested on independent 2024 data\n\n"))

cat(strrep("=", 90), "\n")
cat("  MODEL COMPARISON ON HELD-OUT 2024 DATA\n")
cat(strrep("=", 90), "\n\n")

models <- list(
  "A. 6-factor (theory)" = mA,
  "B. 3-factor (merged)" = mB,
  "C. 1-factor" = mC
)

fits <- list()
for (label in names(models)) {
  cat(sprintf("--- %s ---\n", label))
  tryCatch({
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
cat(strrep("=", 100), "\n")
cat(sprintf("  %-28s %8s %5s %7s %7s %7s %7s %7s\n",
            "Model", "chi2", "df", "chi/df", "CFI", "TLI", "RMSEA", "SRMR"))
cat(strrep("-", 100), "\n")
for (label in names(fits)) {
  fi <- fitMeasures(fits[[label]], c("chisq", "df", "cfi", "tli", "rmsea", "srmr"))
  cat(sprintf("  %-28s %8.1f %5.0f %7.2f %7.3f %7.3f %7.3f %7.3f\n",
              label, fi["chisq"], fi["df"], fi["chisq"]/fi["df"],
              fi["cfi"], fi["tli"], fi["rmsea"], fi["srmr"]))
}

# Factor correlations for 6-factor model
if ("A. 6-factor (theory)" %in% names(fits)) {
  cat("\n\n")
  cat(strrep("=", 80), "\n")
  cat("  6-FACTOR: LATENT CORRELATIONS (2024 hold-out data)\n")
  cat(strrep("=", 80), "\n\n")

  tryCatch({
    corr_lv <- lavInspect(fits[["A. 6-factor (theory)"]], "cor.lv")
    print(round(corr_lv, 3))

    cat("\n  Key pairs:\n")
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
}

cat("\n\nDone.\n")
