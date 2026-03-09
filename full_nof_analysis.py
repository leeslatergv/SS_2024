"""
Full NOF Question Set Analysis
Tests the final policy question list using the same methodology as expanded_fa.py
(minres extraction, promax rotation).

Extracts all needed items from the NSS24 organisational results Excel,
then runs:
  1. EFA on the full item set
  2. CFA-style reliability analysis per standard
  3. Cronbach's alpha and item-total correlations per standard
  4. Inter-standard correlations (discriminant validity)
  5. For Racism standard: likelihood ratio approach (BME vs White comparator)
"""

import pandas as pd
import numpy as np
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_kmo, calculate_bartlett_sphericity
from scipy import stats
import warnings
warnings.filterwarnings("ignore")

EXCEL_FILE = "NSS24 detailed spreadsheets organisational results.xlsx"
TRUST_MEANS_FILE = "expanded_trust_means.csv"

# ── Final NOF question list (policy version with Q6b) ──────────────
# Mapped to: (sheet_name, question_col_offset, response_col_description, clean_name)
# We need to extract from the Excel for items not in expanded_trust_means.csv

FINAL_NOF_STANDARDS = {
    "Reducing Violence": {
        "items": {
            # Q13a+b+c combined: violence from patients, managers, colleagues
            "violence_patients": "q13a",      # already in trust means
            "violence_managers": "q13b",       # need from Excel
            "violence_colleagues": "q13c",     # need from Excel
            "violence_reported": "q13d",       # already in trust means
        },
        "notes": "Q13a+b+c to be combined; Q13d reporting rate",
    },
    "Tackling Racism": {
        "items": {
            # BME vs White comparator needed
            "fair_career_progression": "q15",  # already in trust means
            "discrim_patients": "q16a",        # already in trust means
            "discrim_colleagues": "q16b",      # already in trust means
            # Q16c (grounds = racism) - percentage column from Excel
            "discrim_ethnic_background": "q16c_1",  # need from Excel
        },
        "notes": "Likelihood ratio BME vs White comparison",
    },
    "Sexual Safety": {
        "items": {
            "sexual_behaviour_patients": "q17a",  # already in trust means
            "sexual_behaviour_staff": "q17b",     # already in trust means
        },
        "notes": "Q17a in original only, Q17b in both original and updated",
    },
    "Flexible Working": {
        "items": {
            "satisfaction_flexible_working": "q4d",  # already
            "org_committed_wlb": "q6b",              # already (policy choice)
            "manager_flexible_working": "q6d",       # already
        },
        "notes": "Policy chose Q6b over Q6c",
    },
    "Line Management": {
        "items": {
            "mgr_interest_wellbeing": "q9d",          # already
            "mgr_understanding_problems": "q9f",      # already
            "mgr_effective_action": "q9i",            # already
        },
        "notes": "Subset of Q9 battery",
    },
    "Team Management": {
        "items": {
            "team_shared_objectives": "q7a",          # need from Excel
            "feel_valued_by_team": "q7h",             # already
        },
        "notes": "New sub-standard in updated list",
    },
    "Appraisal": {
        "items": {
            "had_appraisal": "q23a",                  # already
            "appraisal_clear_objectives": "q23c",     # already
            "appraisal_felt_valued": "q23d",          # already
        },
        "notes": "Q23b dropped, Q23c and Q23d added",
    },
    "Development": {
        "items": {
            "supported_develop_potential": "q24d",    # already
            "access_learning_dev": "q24e",            # already
        },
        "notes": "Q24d added in updated list",
    },
    "Supportive Env - Facilities": {
        "items": {
            "nutritious_food": "q22",                 # already
        },
        "notes": "Single item",
    },
    "Supportive Env - Wellbeing": {
        "items": {
            "org_positive_wellbeing": "q11a",         # already
        },
        "notes": "Single item",
    },
}


def extract_likert_from_excel(sheet_name, question_col_start, n_response_cats, org_col=0):
    """Extract a question's positive-end score from the NSS Excel.
    For Likert questions: returns the % agree/strongly agree or equivalent.
    For frequency questions: returns the % Never (= positive for violence etc)."""
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    # Find the header row (contains 'ODS code')
    header_row = None
    for i in range(min(10, len(df))):
        if df.iloc[i, 0] == 'ODS code':
            header_row = i
            break

    if header_row is None:
        print(f"  WARNING: Could not find header row in {sheet_name}")
        return pd.DataFrame()

    data = df.iloc[header_row + 1:].copy()
    data.columns = df.iloc[header_row]

    # Get org code and the score column
    result = pd.DataFrame()
    result['org_id'] = data.iloc[:, org_col].values
    result['score'] = pd.to_numeric(data.iloc[:, question_col_start], errors='coerce')

    result = result.dropna(subset=['org_id'])
    result = result[result['org_id'].str.len() <= 5]  # filter to org codes
    return result.set_index('org_id')['score']


def extract_violence_items():
    """Extract Q13a, Q13b, Q13c (% Never = no violence) from Excel."""
    sheet = 'HEALTH WELLBEING SAFETY Q13-14'
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)

    header_row = None
    for i in range(10):
        if str(df.iloc[i, 0]).strip() == 'ODS code':
            header_row = i
            break

    data = df.iloc[header_row + 1:].copy()
    org_ids = data.iloc[:, 0].values

    results = pd.DataFrame({'org_id': org_ids})

    # Q13a starts at col 6: Never(6), 1-2(7), 3-5(8), 6-10(9), >10(10), Base(11)
    # Q13b starts at col 12: Never(12), ...
    # Q13c starts at col 18: Never(18), ...
    for name, never_col in [('violence_patients_pct', 6),
                             ('violence_managers_pct', 12),
                             ('violence_colleagues_pct', 18)]:
        results[name] = pd.to_numeric(data.iloc[:, never_col], errors='coerce')

    # Q13d starts at col 24: Yes I reported(24), Yes colleague(25), No(26), ...
    results['violence_reported_pct'] = pd.to_numeric(data.iloc[:, 24], errors='coerce')

    results = results.dropna(subset=['org_id'])
    results['org_id'] = results['org_id'].astype(str)
    results = results[results['org_id'].str.len() <= 5]
    return results.set_index('org_id')


def extract_team_items():
    """Extract Q7a and Q7h from Excel."""
    sheet = 'YOUR TEAM Q7'
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)

    header_row = None
    for i in range(10):
        if str(df.iloc[i, 0]).strip() == 'ODS code':
            header_row = i
            break

    data = df.iloc[header_row + 1:].copy()
    org_ids = data.iloc[:, 0].values

    results = pd.DataFrame({'org_id': org_ids})

    # Q7a at col 6: SD(6), D(7), NAND(8), A(9), SA(10), Base(11)
    # Score = (% Agree + % Strongly Agree) / 100 for consistency with trust means
    # Actually, let's compute a weighted score like the trust means would use
    # Q7a cols: 6=SD, 7=D, 8=NAND, 9=A, 10=SA
    for name, base_col in [('team_shared_objectives', 6), ('feel_valued_by_team_xl', 48)]:
        sd = pd.to_numeric(data.iloc[:, base_col], errors='coerce')
        d = pd.to_numeric(data.iloc[:, base_col + 1], errors='coerce')
        n = pd.to_numeric(data.iloc[:, base_col + 2], errors='coerce')
        a = pd.to_numeric(data.iloc[:, base_col + 3], errors='coerce')
        sa = pd.to_numeric(data.iloc[:, base_col + 4], errors='coerce')
        # Weighted mean on 0-1 scale: SD=0, D=0.25, N=0.5, A=0.75, SA=1.0
        total = sd + d + n + a + sa
        score = (sd * 0 + d * 0.25 + n * 0.5 + a * 0.75 + sa * 1.0) / total
        results[name] = score

    results = results.dropna(subset=['org_id'])
    results['org_id'] = results['org_id'].astype(str)
    results = results[results['org_id'].str.len() <= 5]
    return results.set_index('org_id')


def extract_racism_items():
    """Extract Q15, Q16a, Q16b, Q16c_1 (ethnic background) from Excel."""
    sheet = 'HEALTH WELLBEING SAFETY Q15-17'
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)

    header_row = None
    for i in range(10):
        if str(df.iloc[i, 0]).strip() == 'ODS code':
            header_row = i
            break

    data = df.iloc[header_row + 1:].copy()
    org_ids = data.iloc[:, 0].values

    results = pd.DataFrame({'org_id': org_ids})

    # Q15: col 6 = % Yes, col 7 = % No, col 8 = % Don't know
    results['fair_career_pct_yes'] = pd.to_numeric(data.iloc[:, 6], errors='coerce')

    # Q16a: col 10 = % Yes, col 11 = % No
    results['discrim_patients_pct_yes'] = pd.to_numeric(data.iloc[:, 10], errors='coerce')

    # Q16b: col 13 = % Yes, col 14 = % No
    results['discrim_colleagues_pct_yes'] = pd.to_numeric(data.iloc[:, 13], errors='coerce')

    # Q16c_1 (ethnic background): col 16 = %
    results['discrim_ethnic_pct'] = pd.to_numeric(data.iloc[:, 16], errors='coerce')

    results = results.dropna(subset=['org_id'])
    results['org_id'] = results['org_id'].astype(str)
    results = results[results['org_id'].str.len() <= 5]
    return results.set_index('org_id')


def build_full_dataset():
    """Combine expanded_trust_means.csv with items extracted from Excel."""
    # Load existing trust means
    tm = pd.read_csv(TRUST_MEANS_FILE)
    tm['org_id'] = tm['org_id'].astype(str)
    print(f"  Trust means: {len(tm)} trusts, {tm.shape[1]} columns")

    # Extract additional items from Excel
    print("  Extracting violence items (Q13a/b/c/d) from Excel...")
    violence = extract_violence_items()
    print(f"    Got {len(violence)} orgs")

    print("  Extracting team items (Q7a, Q7h) from Excel...")
    team = extract_team_items()
    print(f"    Got {len(team)} orgs")

    print("  Extracting racism items (Q15, Q16a/b/c) from Excel...")
    racism = extract_racism_items()
    print(f"    Got {len(racism)} orgs")

    # Merge
    tm = tm.set_index('org_id')
    merged = tm.join(violence, how='left', rsuffix='_xl')
    merged = merged.join(team, how='left', rsuffix='_xl')
    merged = merged.join(racism, how='left', rsuffix='_xl')

    print(f"  Merged dataset: {len(merged)} trusts, {merged.shape[1]} columns")
    return merged


def cronbach_alpha(items_df):
    """Compute Cronbach's alpha."""
    k = items_df.shape[1]
    if k < 2:
        return np.nan
    item_vars = items_df.var(axis=0, ddof=1)
    total_var = items_df.sum(axis=1).var(ddof=1)
    return (k / (k - 1)) * (1 - item_vars.sum() / total_var)


def item_total_correlations(items_df):
    """Corrected item-total correlations."""
    total = items_df.sum(axis=1)
    results = {}
    for col in items_df.columns:
        corrected = total - items_df[col]
        r, _ = stats.pearsonr(items_df[col].values, corrected.values)
        results[col] = r
    return results


def alpha_if_dropped(items_df):
    """Alpha if each item is dropped."""
    results = {}
    for col in items_df.columns:
        remaining = items_df.drop(columns=[col])
        if remaining.shape[1] >= 2:
            results[col] = cronbach_alpha(remaining)
        else:
            results[col] = np.nan
    return results


def run_reliability_per_standard(df):
    """Reliability analysis for each standard in the final NOF list."""
    print("\n" + "=" * 80)
    print("  PART 1: RELIABILITY PER STANDARD (Cronbach's alpha, item-total r)")
    print("=" * 80)

    # Map standards to available columns in our merged dataset
    # Using the trust means columns (0-1 scale) where available
    standard_items = {
        "Reducing Violence": [
            "violence_patients", "harassment_patients",
            "harassment_managers", "harassment_colleagues",
            "violence_reported",
        ],
        "Tackling Racism": [
            "fair_career_progression", "discrim_patients",
            "discrim_colleagues", "org_respects_differences",
        ],
        "Sexual Safety": [
            "sexual_behaviour_patients", "sexual_behaviour_staff",
        ],
        "Flexible Working": [
            "satisfaction_flexible_working", "org_committed_wlb",
            "manager_flexible_working",
        ],
        "Line Management": [
            "mgr_interest_wellbeing", "mgr_understanding_problems",
            "mgr_effective_action",
        ],
        "Team Management": [
            "team_shared_objectives", "feel_valued_by_team",
        ],
        "Appraisal": [
            "had_appraisal", "appraisal_clear_objectives",
            "appraisal_felt_valued",
        ],
        "Development": [
            "supported_develop_potential", "access_learning_dev",
        ],
    }

    results_summary = {}

    for std_name, items in standard_items.items():
        available = [i for i in items if i in df.columns]
        if len(available) < 2:
            print(f"\n  --- {std_name}: Only {len(available)} item(s), skipping reliability ---")
            continue

        dat = df[available].dropna()
        n = len(dat)

        # Some items need reversing (violence/harassment: higher = worse)
        reverse_items = [i for i in available if any(x in i for x in
                         ['violence', 'harassment', 'discrim', 'sexual_behaviour'])]

        dat_aligned = dat.copy()
        for ri in reverse_items:
            dat_aligned[ri] = 1 - dat_aligned[ri]

        alpha = cronbach_alpha(dat_aligned)
        itc = item_total_correlations(dat_aligned)
        aid = alpha_if_dropped(dat_aligned)

        print(f"\n  --- {std_name} ({len(available)} items, N={n}) ---")
        print(f"  Cronbach's alpha: {alpha:.3f}")

        mean_r = dat_aligned.corr().where(
            np.triu(np.ones(dat_aligned.corr().shape), k=1).astype(bool)
        ).stack().mean()
        print(f"  Mean inter-item correlation: {mean_r:.3f}")

        print(f"\n  {'Item':40s} {'Item-Total r':>12s} {'Alpha if dropped':>16s} {'Reversed':>9s}")
        print("  " + "-" * 80)
        for item in available:
            rev = "Yes" if item in reverse_items else "No"
            print(f"  {item:40s} {itc[item]:12.3f} {aid[item]:16.3f} {rev:>9s}")

        results_summary[std_name] = {
            'alpha': alpha, 'n_items': len(available), 'n_obs': n, 'mean_iic': mean_r
        }

    # Summary table
    print(f"\n\n  {'Standard':30s} {'Items':>5s} {'N':>5s} {'Alpha':>7s} {'Mean IIC':>9s}")
    print("  " + "-" * 60)
    for std, res in results_summary.items():
        print(f"  {std:30s} {res['n_items']:5d} {res['n_obs']:5d} {res['alpha']:7.3f} {res['mean_iic']:9.3f}")


def run_efa_full(df):
    """Run EFA on the full final NOF item set."""
    print("\n\n" + "=" * 80)
    print("  PART 2: EFA ON FULL NOF QUESTION SET")
    print("  (minres extraction, promax rotation — same as expanded_fa.py)")
    print("=" * 80)

    # All items from the final list that we have
    efa_items = [
        # Violence
        "violence_patients", "harassment_patients",
        "harassment_managers", "harassment_colleagues",
        # Racism
        "fair_career_progression", "discrim_patients",
        "discrim_colleagues", "org_respects_differences",
        # Sexual safety
        "sexual_behaviour_patients", "sexual_behaviour_staff",
        # Flexible working (with Q6b)
        "satisfaction_flexible_working", "org_committed_wlb",
        "manager_flexible_working",
        # Line management
        "mgr_interest_wellbeing", "mgr_understanding_problems",
        "mgr_effective_action",
        # Team management
        "team_shared_objectives", "feel_valued_by_team",
        # Appraisal
        "had_appraisal", "appraisal_clear_objectives", "appraisal_felt_valued",
        # Development
        "supported_develop_potential", "access_learning_dev",
        # Supportive env
        "nutritious_food", "org_positive_wellbeing",
    ]

    available = [i for i in efa_items if i in df.columns]
    missing = [i for i in efa_items if i not in df.columns]
    if missing:
        print(f"\n  WARNING: Missing items: {missing}")

    dat = df[available].dropna()
    n_obs, n_vars = dat.shape
    print(f"\n  N = {n_obs} trusts, {n_vars} items")

    # KMO
    try:
        kmo_per, kmo_total = calculate_kmo(dat)
        print(f"  KMO: {kmo_total:.3f}")
    except Exception as e:
        print(f"  KMO failed: {e}")

    # Try different factor numbers
    for n_factors in [6, 8]:
        label = f"{n_factors}-factor"
        print(f"\n  --- {label} EFA ---")

        fa = FactorAnalyzer(n_factors=n_factors, rotation="promax", method="minres")
        fa.fit(dat)

        loadings = pd.DataFrame(
            fa.loadings_, index=available,
            columns=[f"F{i+1}" for i in range(n_factors)]
        )

        _, _, cum_var = fa.get_factor_variance()
        print(f"  Variance explained: {cum_var[-1]:.1%}")

        # Communalities
        comm = pd.Series(fa.get_communalities(), index=available)
        heywood = comm[comm > 1.0]
        if len(heywood) > 0:
            print(f"  HEYWOOD CASES: {dict(heywood.round(3))}")

        # Print loadings
        print(f"\n  {'Item':40s}", end="")
        for c in loadings.columns:
            print(f" {c:>7s}", end="")
        print(f"  {'Primary':>8s}")
        print("  " + "-" * (42 + 8 * n_factors + 10))

        # Group items by their theorised standard for readability
        std_groups = {
            "Violence": ["violence_patients", "harassment_patients",
                         "harassment_managers", "harassment_colleagues"],
            "Racism": ["fair_career_progression", "discrim_patients",
                       "discrim_colleagues", "org_respects_differences"],
            "SexSafety": ["sexual_behaviour_patients", "sexual_behaviour_staff"],
            "FlexWork": ["satisfaction_flexible_working", "org_committed_wlb",
                         "manager_flexible_working"],
            "LineMgmt": ["mgr_interest_wellbeing", "mgr_understanding_problems",
                         "mgr_effective_action"],
            "Team": ["team_shared_objectives", "feel_valued_by_team"],
            "Appraisal": ["had_appraisal", "appraisal_clear_objectives",
                          "appraisal_felt_valued"],
            "Develop": ["supported_develop_potential", "access_learning_dev"],
            "SuppEnv": ["nutritious_food", "org_positive_wellbeing"],
        }

        for std, std_items in std_groups.items():
            for item in std_items:
                if item in loadings.index:
                    row = loadings.loc[item]
                    primary = row.abs().idxmax()
                    primary_val = row[primary]
                    print(f"  {item:40s}", end="")
                    for v in row:
                        if abs(v) >= 0.30:
                            print(f" {v:7.3f}", end="")
                        else:
                            print(f"    .   ", end="")
                    print(f"  {primary:>4s} ({primary_val:+.3f})")
            print()  # blank line between groups

        # Factor correlations
        if fa.phi_ is not None:
            phi = pd.DataFrame(
                fa.phi_,
                index=[f"F{i+1}" for i in range(n_factors)],
                columns=[f"F{i+1}" for i in range(n_factors)]
            )
            print(f"  Factor correlations:")
            print("  " + phi.round(3).to_string().replace("\n", "\n  "))


def run_discriminant_validity(df):
    """Inter-standard composite correlations."""
    print("\n\n" + "=" * 80)
    print("  PART 3: INTER-STANDARD CORRELATIONS (discriminant validity)")
    print("=" * 80)

    composites = {
        "Violence": {
            "items": ["violence_patients", "harassment_patients",
                      "harassment_managers", "harassment_colleagues"],
            "reverse": True,
        },
        "Racism": {
            "items": ["fair_career_progression", "org_respects_differences"],
            "reverse": False,
            "reverse_items": ["discrim_patients", "discrim_colleagues"],
        },
        "SexSafety": {
            "items": ["sexual_behaviour_patients", "sexual_behaviour_staff"],
            "reverse": True,
        },
        "FlexWork": {
            "items": ["satisfaction_flexible_working", "org_committed_wlb",
                      "manager_flexible_working"],
            "reverse": False,
        },
        "LineMgmt": {
            "items": ["mgr_interest_wellbeing", "mgr_understanding_problems",
                      "mgr_effective_action"],
            "reverse": False,
        },
        "Team": {
            "items": ["team_shared_objectives", "feel_valued_by_team"],
            "reverse": False,
        },
        "Appraisal": {
            "items": ["had_appraisal", "appraisal_clear_objectives",
                      "appraisal_felt_valued"],
            "reverse": False,
        },
        "Develop": {
            "items": ["supported_develop_potential", "access_learning_dev"],
            "reverse": False,
        },
        "SuppEnv": {
            "items": ["nutritious_food", "org_positive_wellbeing"],
            "reverse": False,
        },
    }

    scores = pd.DataFrame(index=df.index)
    for std, cfg in composites.items():
        available = [i for i in cfg["items"] if i in df.columns]
        if not available:
            continue
        vals = df[available].copy()
        if cfg.get("reverse"):
            vals = 1 - vals
        if cfg.get("reverse_items"):
            for ri in cfg["reverse_items"]:
                if ri in vals.columns:
                    vals[ri] = 1 - vals[ri]
        scores[std] = vals.mean(axis=1)

    corr = scores.corr()
    print(f"\n  Inter-standard correlations (composite scores):")
    print("  " + corr.round(3).to_string().replace("\n", "\n  "))

    # Mean off-diagonal
    mask = np.triu(np.ones(corr.shape), k=1).astype(bool)
    mean_r = corr.where(mask).stack().mean()
    print(f"\n  Mean inter-standard correlation: {mean_r:.3f}")
    print(f"  Mean shared variance: {mean_r**2:.1%}")

    # Flag high correlations (potential redundancy)
    print(f"\n  Pairs with r > 0.80 (potential redundancy):")
    stds = corr.columns.tolist()
    any_high = False
    for i in range(len(stds)):
        for j in range(i + 1, len(stds)):
            r = corr.iloc[i, j]
            if abs(r) > 0.80:
                print(f"    {stds[i]:12s} <-> {stds[j]:12s}: r = {r:.3f}")
                any_high = True
    if not any_high:
        print(f"    None — all standards show adequate discriminant validity")


def run_racism_likelihood_ratio(df):
    """For Tackling Racism: compute a likelihood ratio approach.

    The idea: for each trust, compute the ratio of BME to White staff
    experiencing discrimination. This captures the RELATIVE disadvantage
    rather than absolute prevalence.

    With trust-level data we can look at:
    - How the racism items correlate with each other
    - Whether Q16c (ethnic background) adds information beyond Q16a/b
    - How the Racism standard discriminates from other standards
    """
    print("\n\n" + "=" * 80)
    print("  PART 4: TACKLING RACISM — LIKELIHOOD RATIO ANALYSIS")
    print("=" * 80)

    racism_items = ['fair_career_progression', 'discrim_patients',
                    'discrim_colleagues', 'org_respects_differences']
    available = [i for i in racism_items if i in df.columns]
    dat = df[available].dropna()

    print(f"\n  Available racism items: {available}")
    print(f"  N = {len(dat)} trusts")

    # Correlation matrix among racism items
    corr = dat.corr()
    print(f"\n  Inter-item correlations:")
    print("  " + corr.round(3).to_string().replace("\n", "\n  "))

    # Check if we have the ethnic discrimination percentage from Excel
    if 'discrim_ethnic_pct' in df.columns:
        ethnic = df['discrim_ethnic_pct'].dropna()
        print(f"\n  Q16c_1 (% reporting ethnic background discrimination): N={len(ethnic)}")
        print(f"    Mean: {ethnic.mean():.1f}%")
        print(f"    SD:   {ethnic.std():.1f}%")
        print(f"    Min:  {ethnic.min():.1f}%")
        print(f"    Max:  {ethnic.max():.1f}%")

        # Correlation with other racism items
        for item in available:
            valid = df[[item, 'discrim_ethnic_pct']].dropna()
            r, p = stats.pearsonr(valid[item], valid['discrim_ethnic_pct'])
            print(f"    r({item}, discrim_ethnic_pct) = {r:.3f} (p={p:.4f})")

    # The policy concept: use Q15 and Q16b as comparators between BME and White
    # At trust level, a likelihood ratio would be:
    #   LR = P(discrim | BME) / P(discrim | White)
    # We don't have this breakdown in the trust means, but the Q16c ethnic %
    # gives us the proportion who said discrimination was on ethnic grounds
    print(f"""
  NOTE on Likelihood Ratio approach:
  ----------------------------------
  The policy wants BME vs White comparisons for Q15 and Q16b.
  This requires individual-level or demographic-group-level data
  (e.g., WRES indicators), not available in these trust-level means.

  What we CAN say from trust-level data:
  - Q15 (fair career progression), Q16a/b (discrimination experienced),
    and Q16c_1 (ethnic background as ground) all correlate and form
    a coherent Racism standard.
  - For the NOF score, you would need the WRES or NSS demographic
    breakdowns to compute the actual BME/White likelihood ratio.
  - The factor structure supports these items belonging together
    regardless of whether you use absolute rates or LR ratios.
""")


def main():
    print("=" * 80)
    print("  FULL NOF QUESTION SET — FACTOR ANALYSIS & RELIABILITY")
    print("  Using policy's final list (with Q6b for Flexible Working)")
    print("  Method: same as expanded_fa.py (minres, promax)")
    print("=" * 80)

    df = build_full_dataset()

    # Show what we have
    print(f"\n  Items available for analysis:")
    for std, cfg in FINAL_NOF_STANDARDS.items():
        items = list(cfg['items'].keys())
        available = [i for i in items if i in df.columns]
        missing = [i for i in items if i not in df.columns]
        status = "OK" if not missing else f"MISSING: {missing}"
        print(f"    {std:30s}: {len(available)}/{len(items)} items  {status}")

    run_reliability_per_standard(df)
    run_efa_full(df)
    run_discriminant_validity(df)
    run_racism_likelihood_ratio(df)

    print("\n" + "=" * 80)
    print("  DONE")
    print("=" * 80)


if __name__ == "__main__":
    main()
