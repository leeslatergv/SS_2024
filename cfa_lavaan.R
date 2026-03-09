# NHS Staff Standards - CFA using lavaan (gold standard)
# Compares: 1-factor, 4-factor (EFA-driven), 6-factor, 5-factor (merged), bifactor

library(lavaan)

# Load data
df <- read.csv("~/nhs-staff-standards/data/expanded_trust_means.csv")

# Create composites for near-collinear items
mgr_items <- c("mgr_encourages", "mgr_clear_feedback", "mgr_asks_opinion",
               "mgr_interest_wellbeing", "mgr_values_work", "mgr_understanding_problems",
               "mgr_listens", "mgr_cares_concerns", "mgr_effective_action")
df$mgr_composite <- rowMeans(df[, mgr_items], na.rm = TRUE)

people_items <- c("people_kind_to_each_other", "people_polite_respectful", "people_show_appreciation")
df$people_culture <- rowMeans(df[, people_items], na.rm = TRUE)

# Select items and drop NAs
items <- c(
  "violence_patients", "harassment_patients", "harassment_managers", "harassment_colleagues",
  "fair_career_progression", "discrim_patients", "discrim_colleagues", "org_respects_differences",
  "sexual_behaviour_patients", "sexual_behaviour_staff",
  "satisfaction_flexible_working", "good_wl_balance",
  "mgr_composite", "had_appraisal", "appraisal_improved_job", "access_learning_dev",
  "adequate_materials_equipment", "enough_staff", "people_culture", "org_positive_wellbeing"
)

dat <- df[complete.cases(df[, items]), items]
cat(sprintf("N = %d trusts, %d items\n\n", nrow(dat), ncol(dat)))

# ── Model 1: Single factor ──────────────────────────────────────────
m1_spec <- paste0("General =~ ", paste(items, collapse = " + "))

# ── Model 2: 4-factor (EFA-driven) ─────────────────────────────────
m2_spec <- '
  ManagerTeam =~ mgr_composite + satisfaction_flexible_working + good_wl_balance + people_culture
  OrgCulture =~ fair_career_progression + org_respects_differences + adequate_materials_equipment + enough_staff + org_positive_wellbeing
  PatientHarm =~ violence_patients + harassment_patients + sexual_behaviour_patients + sexual_behaviour_staff
  AppraisalQual =~ had_appraisal + appraisal_improved_job + access_learning_dev
'

# ── Model 3: 6-factor correlated (theory) ──────────────────────────
m3_spec <- '
  Violence =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues
  Racism =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
  SexSafety =~ sexual_behaviour_patients + sexual_behaviour_staff
  FlexWorking =~ satisfaction_flexible_working + good_wl_balance
  LineMgmt =~ mgr_composite + had_appraisal + appraisal_improved_job + access_learning_dev
  SupportEnv =~ adequate_materials_equipment + enough_staff + people_culture + org_positive_wellbeing
'

# ── Model 4: 3-factor (maximally merged) ──────────────────────────
m4_spec <- '
  GoodManagement =~ mgr_composite + satisfaction_flexible_working + good_wl_balance + people_culture + adequate_materials_equipment + enough_staff + org_positive_wellbeing + had_appraisal + appraisal_improved_job + access_learning_dev
  PatientHarm =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues + sexual_behaviour_patients + sexual_behaviour_staff
  Fairness =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
'

# ── Model 5: Bifactor (general + 6 specific) ──────────────────────
m5_spec <- '
  General =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues + fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences + sexual_behaviour_patients + sexual_behaviour_staff + satisfaction_flexible_working + good_wl_balance + mgr_composite + had_appraisal + appraisal_improved_job + access_learning_dev + adequate_materials_equipment + enough_staff + people_culture + org_positive_wellbeing
  Violence =~ violence_patients + harassment_patients + harassment_managers + harassment_colleagues
  Racism =~ fair_career_progression + discrim_patients + discrim_colleagues + org_respects_differences
  SexSafety =~ sexual_behaviour_patients + sexual_behaviour_staff
  FlexWorking =~ satisfaction_flexible_working + good_wl_balance
  LineMgmt =~ mgr_composite + had_appraisal + appraisal_improved_job + access_learning_dev
  SupportEnv =~ adequate_materials_equipment + enough_staff + people_culture + org_positive_wellbeing
'
# Bifactor constraints: all factors orthogonal
m5_constraints <- '
  General ~~ 0*Violence + 0*Racism + 0*SexSafety + 0*FlexWorking + 0*LineMgmt + 0*SupportEnv
  Violence ~~ 0*Racism + 0*SexSafety + 0*FlexWorking + 0*LineMgmt + 0*SupportEnv
  Racism ~~ 0*SexSafety + 0*FlexWorking + 0*LineMgmt + 0*SupportEnv
  SexSafety ~~ 0*FlexWorking + 0*LineMgmt + 0*SupportEnv
  FlexWorking ~~ 0*LineMgmt + 0*SupportEnv
  LineMgmt ~~ 0*SupportEnv
'
m5_full <- paste0(m5_spec, "\n", m5_constraints)

# Fit all models
cat("Fitting models...\n")
results <- list()

for (label in c("1-factor", "4-factor", "6-factor", "3-factor", "Bifactor")) {
  spec <- switch(label,
    "1-factor" = m1_spec,
    "4-factor" = m2_spec,
    "6-factor" = m3_spec,
    "3-factor" = m4_spec,
    "Bifactor" = m5_full
  )

  cat(sprintf("\n--- %s ---\n", label))
  tryCatch({
    fit <- cfa(spec, data = dat, estimator = "ML")
    results[[label]] <- fit

    fi <- fitMeasures(fit, c("chisq", "df", "cfi", "tli", "rmsea", "srmr", "aic", "bic"))
    cat(sprintf("  chi2 = %.1f, df = %.0f, chi2/df = %.2f\n", fi["chisq"], fi["df"], fi["chisq"]/fi["df"]))
    cat(sprintf("  CFI  = %.3f  %s\n", fi["cfi"], ifelse(fi["cfi"] > 0.95, "GOOD", ifelse(fi["cfi"] > 0.90, "ACCEPTABLE", "POOR"))))
    cat(sprintf("  TLI  = %.3f  %s\n", fi["tli"], ifelse(fi["tli"] > 0.95, "GOOD", ifelse(fi["tli"] > 0.90, "ACCEPTABLE", "POOR"))))
    cat(sprintf("  RMSEA= %.3f  %s\n", fi["rmsea"], ifelse(fi["rmsea"] < 0.06, "GOOD", ifelse(fi["rmsea"] < 0.08, "ACCEPTABLE", "POOR"))))
    cat(sprintf("  SRMR = %.3f  %s\n", fi["srmr"], ifelse(fi["srmr"] < 0.08, "GOOD", "POOR")))
    cat(sprintf("  AIC  = %.1f\n", fi["aic"]))
    cat(sprintf("  BIC  = %.1f\n", fi["bic"]))
  }, warning = function(w) {
    cat(sprintf("  WARNING: %s\n", conditionMessage(w)))
    tryCatch({
      fit <- cfa(spec, data = dat, estimator = "ML")
      results[[label]] <<- fit
      fi <- fitMeasures(fit, c("chisq", "df", "cfi", "tli", "rmsea", "srmr", "aic", "bic"))
      cat(sprintf("  chi2 = %.1f, df = %.0f, chi2/df = %.2f\n", fi["chisq"], fi["df"], fi["chisq"]/fi["df"]))
      cat(sprintf("  CFI  = %.3f\n  TLI  = %.3f\n  RMSEA= %.3f\n  SRMR = %.3f\n", fi["cfi"], fi["tli"], fi["rmsea"], fi["srmr"]))
    }, error = function(e) {
      cat(sprintf("  FAILED: %s\n", conditionMessage(e)))
    })
  }, error = function(e) {
    cat(sprintf("  FAILED: %s\n", conditionMessage(e)))
  })
}

# ── Model comparison table ─────────────────────────────────────────
cat("\n\n")
cat(strrep("=", 100))
cat("\n  MODEL COMPARISON\n")
cat(strrep("=", 100))
cat("\n\n")
cat(sprintf("  %-25s %8s %5s %7s %7s %7s %7s %12s %12s\n",
            "Model", "chi2", "df", "chi2/df", "CFI", "TLI", "RMSEA", "AIC", "BIC"))
cat(strrep("-", 100))
cat("\n")

for (label in names(results)) {
  fit <- results[[label]]
  fi <- fitMeasures(fit, c("chisq", "df", "cfi", "tli", "rmsea", "aic", "bic"))
  cat(sprintf("  %-25s %8.1f %5.0f %7.2f %7.3f %7.3f %7.3f %12.1f %12.1f\n",
              label, fi["chisq"], fi["df"], fi["chisq"]/fi["df"],
              fi["cfi"], fi["tli"], fi["rmsea"], fi["aic"], fi["bic"]))
}

# ── Factor correlations for 6-factor model ─────────────────────────
if ("6-factor" %in% names(results)) {
  cat("\n\n")
  cat(strrep("=", 80))
  cat("\n  6-FACTOR CFA: FACTOR CORRELATIONS\n")
  cat(strrep("=", 80))
  cat("\n\n")

  fit6 <- results[["6-factor"]]
  vcov <- lavInspect(fit6, "cor.lv")
  print(round(vcov, 3))

  cat("\n  Key redundancies:\n")
  fnames <- rownames(vcov)
  for (i in 1:(length(fnames)-1)) {
    for (j in (i+1):length(fnames)) {
      r <- vcov[i,j]
      label <- if (abs(r) > 0.90) "REDUNDANT" else if (abs(r) > 0.80) "HIGH" else if (abs(r) > 0.60) "MODERATE" else "DISTINCT"
      cat(sprintf("    %-15s <-> %-15s: r=%+.3f  [%s]\n", fnames[i], fnames[j], r, label))
    }
  }
}

# ── Standardised loadings for 6-factor model ───────────────────────
if ("6-factor" %in% names(results)) {
  cat("\n\n")
  cat(strrep("=", 80))
  cat("\n  6-FACTOR CFA: STANDARDISED LOADINGS\n")
  cat(strrep("=", 80))
  cat("\n\n")

  std_est <- standardizedSolution(results[["6-factor"]])
  loadings <- std_est[std_est$op == "=~", c("lhs", "rhs", "est.std", "pvalue")]
  cat(sprintf("  %-15s %-35s %8s %10s\n", "Factor", "Item", "Std.Load", "p-value"))
  cat(strrep("-", 72))
  cat("\n")
  for (i in 1:nrow(loadings)) {
    cat(sprintf("  %-15s %-35s %8.3f %10.4f\n",
                loadings$lhs[i], loadings$rhs[i], loadings$est.std[i], loadings$pvalue[i]))
  }
}

# ── Chi-square difference tests ────────────────────────────────────
cat("\n\n")
cat(strrep("=", 80))
cat("\n  CHI-SQUARE DIFFERENCE TESTS\n")
cat(strrep("=", 80))
cat("\n")

if (all(c("1-factor", "6-factor") %in% names(results))) {
  cat("\n  6-factor vs 1-factor:\n")
  print(anova(results[["1-factor"]], results[["6-factor"]]))
}
if (all(c("3-factor", "6-factor") %in% names(results))) {
  cat("\n  6-factor vs 3-factor:\n")
  print(anova(results[["3-factor"]], results[["6-factor"]]))
}
if (all(c("4-factor", "6-factor") %in% names(results))) {
  cat("\n  6-factor vs 4-factor:\n")
  print(anova(results[["4-factor"]], results[["6-factor"]]))
}
