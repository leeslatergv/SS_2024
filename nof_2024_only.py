"""
Full NOF Question Set Analysis — 2024 NSS DATA ONLY
Extracts ALL items directly from the NSS24 organisational results Excel.
No multi-year trust means — pure 2024 snapshot.

Scoring: weighted mean on 0-1 scale for all Likert items.
  - 5-point agree:  SD=0, D=0.25, N=0.5, A=0.75, SA=1.0
  - 5-point satisfy: VD=0, D=0.25, N=0.5, S=0.75, VS=1.0
  - 5-point freq:    Never=0, Rarely=0.25, Sometimes=0.5, Often=0.75, Always=1.0
  - Violence freq:   Never=0, 1-2=0.25, 3-5=0.5, 6-10=0.75, >10=1.0
  - Yes/No:          Yes=1, No=0 (or weighted % Yes / 100)
  - 3-point appraisal: Yes definitely=1, Yes to some extent=0.5, No=0

Method: same as expanded_fa.py (minres extraction, promax rotation)
"""

import pandas as pd
import numpy as np
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_kmo
from scipy import stats
import warnings
warnings.filterwarnings("ignore")

EXCEL_FILE = "NSS24 detailed spreadsheets organisational results.xlsx"


def read_sheet(sheet_name):
    """Read a sheet, find header row, return clean DataFrame."""
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)
    header_row = None
    for i in range(10):
        if str(df.iloc[i, 0]).strip() == 'ODS code':
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"No header row in {sheet_name}")
    data = df.iloc[header_row + 1:].reset_index(drop=True)
    # Filter to real org codes (3-5 chars, alphanumeric)
    org_ids = data.iloc[:, 0].astype(str)
    mask = org_ids.str.match(r'^[A-Z0-9]{2,5}$', na=False)
    data = data[mask].reset_index(drop=True)
    return data


def score_5pt_likert(data, col_start):
    """Score 5-point Likert (SD/D/N/A/SA or VD/D/N/S/VS) as 0-1 weighted mean."""
    c1 = pd.to_numeric(data.iloc[:, col_start], errors='coerce')      # SD/VD = 0
    c2 = pd.to_numeric(data.iloc[:, col_start + 1], errors='coerce')  # D = 0.25
    c3 = pd.to_numeric(data.iloc[:, col_start + 2], errors='coerce')  # N = 0.5
    c4 = pd.to_numeric(data.iloc[:, col_start + 3], errors='coerce')  # A/S = 0.75
    c5 = pd.to_numeric(data.iloc[:, col_start + 4], errors='coerce')  # SA/VS = 1.0
    total = c1 + c2 + c3 + c4 + c5
    score = (c1 * 0 + c2 * 0.25 + c3 * 0.5 + c4 * 0.75 + c5 * 1.0) / total
    return score


def score_5pt_frequency(data, col_start):
    """Score 5-point frequency (Never/Rarely/Sometimes/Often/Always) as 0-1."""
    c1 = pd.to_numeric(data.iloc[:, col_start], errors='coerce')      # Never = 0
    c2 = pd.to_numeric(data.iloc[:, col_start + 1], errors='coerce')  # Rarely = 0.25
    c3 = pd.to_numeric(data.iloc[:, col_start + 2], errors='coerce')  # Sometimes = 0.5
    c4 = pd.to_numeric(data.iloc[:, col_start + 3], errors='coerce')  # Often = 0.75
    c5 = pd.to_numeric(data.iloc[:, col_start + 4], errors='coerce')  # Always = 1.0
    total = c1 + c2 + c3 + c4 + c5
    score = (c1 * 0 + c2 * 0.25 + c3 * 0.5 + c4 * 0.75 + c5 * 1.0) / total
    return score


def score_violence_freq(data, col_start):
    """Score violence frequency (Never/1-2/3-5/6-10/>10) as 0-1.
    Higher = MORE violence."""
    c1 = pd.to_numeric(data.iloc[:, col_start], errors='coerce')      # Never = 0
    c2 = pd.to_numeric(data.iloc[:, col_start + 1], errors='coerce')  # 1-2 = 0.25
    c3 = pd.to_numeric(data.iloc[:, col_start + 2], errors='coerce')  # 3-5 = 0.5
    c4 = pd.to_numeric(data.iloc[:, col_start + 3], errors='coerce')  # 6-10 = 0.75
    c5 = pd.to_numeric(data.iloc[:, col_start + 4], errors='coerce')  # >10 = 1.0
    total = c1 + c2 + c3 + c4 + c5
    score = (c1 * 0 + c2 * 0.25 + c3 * 0.5 + c4 * 0.75 + c5 * 1.0) / total
    return score


def score_yes_no(data, col_yes):
    """Score Yes/No as proportion Yes (0-1). Higher = more Yes."""
    return pd.to_numeric(data.iloc[:, col_yes], errors='coerce') / 100.0


def score_3pt_appraisal(data, col_start):
    """Score 3-point appraisal (Yes definitely / Yes to some extent / No) as 0-1."""
    c1 = pd.to_numeric(data.iloc[:, col_start], errors='coerce')      # Yes def = 1.0
    c2 = pd.to_numeric(data.iloc[:, col_start + 1], errors='coerce')  # Yes some = 0.5
    c3 = pd.to_numeric(data.iloc[:, col_start + 2], errors='coerce')  # No = 0
    total = c1 + c2 + c3
    score = (c1 * 1.0 + c2 * 0.5 + c3 * 0) / total
    return score


def score_reported(data, col_start):
    """Score reporting (Yes I reported / Yes colleague / Both / No / Don't know / NA).
    Score = proportion who reported (self or colleague or both)."""
    yes_self = pd.to_numeric(data.iloc[:, col_start], errors='coerce')
    yes_coll = pd.to_numeric(data.iloc[:, col_start + 1], errors='coerce')
    no = pd.to_numeric(data.iloc[:, col_start + 2], errors='coerce')
    dk = pd.to_numeric(data.iloc[:, col_start + 3], errors='coerce')
    both = pd.to_numeric(data.iloc[:, col_start + 4], errors='coerce')
    reported = yes_self + yes_coll + both
    total = reported + no + dk
    return reported / total


def extract_all_nof_items():
    """Extract every NOF question from the 2024 Excel, scored on 0-1 scale."""

    items = {}

    # ── Violence: Q13a/b/c (freq), Q13d (reported) ─────────────────
    d = read_sheet('HEALTH WELLBEING SAFETY Q13-14')
    items['org_id'] = d.iloc[:, 0].astype(str)
    items['org_name'] = d.iloc[:, 1].astype(str)
    items['violence_patients'] = score_violence_freq(d, 6)    # Q13a col 6
    items['violence_managers'] = score_violence_freq(d, 12)   # Q13b col 12
    items['violence_colleagues'] = score_violence_freq(d, 18) # Q13c col 18
    items['violence_reported'] = score_reported(d, 24)        # Q13d col 24
    # Q14a/b/c harassment
    items['harassment_patients'] = score_violence_freq(d, 31)   # Q14a col 31
    items['harassment_managers'] = score_violence_freq(d, 37)   # Q14b col 37
    items['harassment_colleagues'] = score_violence_freq(d, 43) # Q14c col 43

    df = pd.DataFrame(items)

    # ── Racism: Q15, Q16a, Q16b, Q16c_1 ─────────────────────────
    d = read_sheet('HEALTH WELLBEING SAFETY Q15-17')
    racism = pd.DataFrame()
    racism['org_id'] = d.iloc[:, 0].astype(str)
    racism['fair_career_progression'] = score_yes_no(d, 6)      # Q15 col 6 (% Yes)
    racism['discrim_patients'] = score_yes_no(d, 10)             # Q16a col 10 (% Yes)
    racism['discrim_colleagues'] = score_yes_no(d, 13)           # Q16b col 13 (% Yes)
    # Q16c_1 (% ethnic background) - of those who experienced discrimination
    racism['discrim_ethnic_pct'] = pd.to_numeric(d.iloc[:, 16], errors='coerce')

    # Sexual safety from same sheet
    racism['sexual_behaviour_patients'] = score_violence_freq(d, 30)  # Q17a col 30
    racism['sexual_behaviour_staff'] = score_violence_freq(d, 36)     # Q17b col 36

    df = df.merge(racism, on='org_id', how='outer')

    # ── Org respects differences: Q21 ────────────────────────────
    d = read_sheet('HEALTH WELLBEING SAFETY Q18-22')
    q21 = pd.DataFrame()
    q21['org_id'] = d.iloc[:, 0].astype(str)
    q21['org_respects_differences'] = score_5pt_likert(d, 49)  # Q21 col 49
    q21['nutritious_food'] = score_5pt_frequency(d, 55)         # Q22 col 55
    df = df.merge(q21, on='org_id', how='outer')

    # ── Flexible Working: Q4d, Q6b, Q6c, Q6d ────────────────────
    d = read_sheet('YOUR JOB Q4-6')
    flex = pd.DataFrame()
    flex['org_id'] = d.iloc[:, 0].astype(str)
    # Q4d: satisfaction with flexible working (5pt satisfaction, col 24)
    flex['satisfaction_flexible_working'] = score_5pt_likert(d, 24)
    # Q6b: org committed WLB (5pt agree, col 55)
    flex['org_committed_wlb'] = score_5pt_likert(d, 55)
    # Q6c: good WL balance (5pt agree, col 61)
    flex['good_wl_balance'] = score_5pt_likert(d, 61)
    # Q6d: manager flexible working (5pt agree, col 67)
    flex['manager_flexible_working'] = score_5pt_likert(d, 67)
    df = df.merge(flex, on='org_id', how='outer')

    # ── Line Management: Q9d, Q9f, Q9i ──────────────────────────
    d = read_sheet('YOUR MANAGERS Q9')
    mgr = pd.DataFrame()
    mgr['org_id'] = d.iloc[:, 0].astype(str)
    mgr['mgr_interest_wellbeing'] = score_5pt_likert(d, 24)        # Q9d col 24
    mgr['mgr_understanding_problems'] = score_5pt_likert(d, 36)    # Q9f col 36
    mgr['mgr_effective_action'] = score_5pt_likert(d, 54)          # Q9i col 54
    df = df.merge(mgr, on='org_id', how='outer')

    # ── Team Management: Q7a, Q7h ────────────────────────────────
    d = read_sheet('YOUR TEAM Q7')
    team = pd.DataFrame()
    team['org_id'] = d.iloc[:, 0].astype(str)
    team['team_shared_objectives'] = score_5pt_likert(d, 6)    # Q7a col 6
    team['feel_valued_by_team'] = score_5pt_likert(d, 48)      # Q7h col 48
    df = df.merge(team, on='org_id', how='outer')

    # ── Appraisal: Q23a, Q23c, Q23d ─────────────────────────────
    d = read_sheet('PERSONAL DEVELOPMENT Q23-24')
    appr = pd.DataFrame()
    appr['org_id'] = d.iloc[:, 0].astype(str)
    appr['had_appraisal'] = score_yes_no(d, 6)                        # Q23a col 6 (% Yes)
    appr['appraisal_clear_objectives'] = score_3pt_appraisal(d, 14)   # Q23c col 14
    appr['appraisal_felt_valued'] = score_3pt_appraisal(d, 18)        # Q23d col 18

    # Development: Q24d, Q24e
    appr['supported_develop_potential'] = score_5pt_likert(d, 40)  # Q24d col 40
    appr['access_learning_dev'] = score_5pt_likert(d, 46)         # Q24e col 46
    df = df.merge(appr, on='org_id', how='outer')

    # ── Wellbeing: Q11a ──────────────────────────────────────────
    d = read_sheet('HEALTH WELLBEING SAFETY Q10-11')
    wb = pd.DataFrame()
    wb['org_id'] = d.iloc[:, 0].astype(str)
    wb['org_positive_wellbeing'] = score_5pt_likert(d, 19)    # Q11a col 19
    df = df.merge(wb, on='org_id', how='outer')

    print(f"  Extracted {len(df)} organisations, {df.shape[1]} columns")
    print(f"  Columns: {list(df.columns)}")

    # Filter to trusts only (exclude CSUs etc) - keep orgs with reasonable data
    n_items = df.select_dtypes(include=[np.number]).count(axis=1)
    df = df[n_items >= 10].reset_index(drop=True)
    print(f"  After filtering (>=10 items): {len(df)} organisations")

    return df


def cronbach_alpha(items_df):
    k = items_df.shape[1]
    if k < 2:
        return np.nan
    item_vars = items_df.var(axis=0, ddof=1)
    total_var = items_df.sum(axis=1).var(ddof=1)
    return (k / (k - 1)) * (1 - item_vars.sum() / total_var)


def item_total_correlations(items_df):
    total = items_df.sum(axis=1)
    results = {}
    for col in items_df.columns:
        corrected = total - items_df[col]
        r, _ = stats.pearsonr(items_df[col].values, corrected.values)
        results[col] = r
    return results


def alpha_if_dropped(items_df):
    results = {}
    for col in items_df.columns:
        remaining = items_df.drop(columns=[col])
        results[col] = cronbach_alpha(remaining) if remaining.shape[1] >= 2 else np.nan
    return results


def run_q6b_vs_q6c(df):
    """Quick head-to-head: Q6b vs Q6c in FlexWorking, same as compare_q6b_q6c.py."""
    print("\n" + "=" * 80)
    print("  Q6b vs Q6c COMPARISON (2024 data only)")
    print("=" * 80)

    for label, flex_item in [("Q6c (good_wl_balance)", "good_wl_balance"),
                              ("Q6b (org_committed_wlb)", "org_committed_wlb")]:
        items = ["satisfaction_flexible_working", flex_item, "manager_flexible_working"]
        available = [i for i in items if i in df.columns]
        dat = df[available].dropna()
        n = len(dat)

        alpha = cronbach_alpha(dat)
        itc = item_total_correlations(dat)
        mean_iic = dat.corr().where(
            np.triu(np.ones(dat.corr().shape), k=1).astype(bool)
        ).stack().mean()

        print(f"\n  --- FlexWorking with {label} (N={n}) ---")
        print(f"  Cronbach's alpha: {alpha:.3f}")
        print(f"  Mean inter-item r: {mean_iic:.3f}")
        for item, r in itc.items():
            print(f"    {item:40s}  item-total r = {r:.3f}")

    # EFA with both to see loadings
    all_flex = ["satisfaction_flexible_working", "good_wl_balance",
                "org_committed_wlb", "manager_flexible_working"]
    available = [i for i in all_flex if i in df.columns]
    dat = df[available].dropna()

    print(f"\n  Inter-correlations (2024, N={len(dat)}):")
    print("  " + dat.corr().round(3).to_string().replace("\n", "\n  "))


def run_reliability(df):
    """Reliability analysis per standard."""
    print("\n\n" + "=" * 80)
    print("  PART 1: RELIABILITY PER STANDARD (2024 data)")
    print("=" * 80)

    standard_items = {
        "Reducing Violence": {
            "items": ["violence_patients", "violence_managers", "violence_colleagues",
                      "violence_reported"],
            "reverse": ["violence_patients", "violence_managers", "violence_colleagues"],
        },
        "Tackling Racism": {
            "items": ["fair_career_progression", "discrim_patients",
                      "discrim_colleagues", "org_respects_differences"],
            "reverse": ["discrim_patients", "discrim_colleagues"],
        },
        "Sexual Safety": {
            "items": ["sexual_behaviour_patients", "sexual_behaviour_staff"],
            "reverse": ["sexual_behaviour_patients", "sexual_behaviour_staff"],
        },
        "Flexible Working": {
            "items": ["satisfaction_flexible_working", "org_committed_wlb",
                      "manager_flexible_working"],
            "reverse": [],
        },
        "Line Management": {
            "items": ["mgr_interest_wellbeing", "mgr_understanding_problems",
                      "mgr_effective_action"],
            "reverse": [],
        },
        "Team Management": {
            "items": ["team_shared_objectives", "feel_valued_by_team"],
            "reverse": [],
        },
        "Appraisal": {
            "items": ["had_appraisal", "appraisal_clear_objectives",
                      "appraisal_felt_valued"],
            "reverse": [],
        },
        "Development": {
            "items": ["supported_develop_potential", "access_learning_dev"],
            "reverse": [],
        },
    }

    summary = {}

    for std_name, cfg in standard_items.items():
        items = cfg["items"]
        rev = cfg["reverse"]
        available = [i for i in items if i in df.columns]
        if len(available) < 2:
            print(f"\n  --- {std_name}: Only {len(available)} item(s), skipping ---")
            continue

        dat = df[available].dropna()
        n = len(dat)

        dat_aligned = dat.copy()
        for ri in rev:
            if ri in dat_aligned.columns:
                dat_aligned[ri] = 1 - dat_aligned[ri]

        alpha = cronbach_alpha(dat_aligned)
        itc = item_total_correlations(dat_aligned)
        aid = alpha_if_dropped(dat_aligned)
        mean_iic = dat_aligned.corr().where(
            np.triu(np.ones(dat_aligned.corr().shape), k=1).astype(bool)
        ).stack().mean()

        print(f"\n  --- {std_name} ({len(available)} items, N={n}) ---")
        print(f"  Cronbach's alpha: {alpha:.3f}")
        print(f"  Mean inter-item r: {mean_iic:.3f}")
        print(f"\n  {'Item':40s} {'Item-Total r':>12s} {'Alpha if drop':>14s} {'Reversed':>9s}")
        print("  " + "-" * 78)
        for item in available:
            r_flag = "Yes" if item in rev else "No"
            print(f"  {item:40s} {itc[item]:12.3f} {aid[item]:14.3f} {r_flag:>9s}")

        summary[std_name] = {'alpha': alpha, 'n': len(available), 'N': n, 'mean_iic': mean_iic}

    print(f"\n\n  {'Standard':25s} {'Items':>5s} {'N':>5s} {'Alpha':>7s} {'Mean IIC':>9s}")
    print("  " + "-" * 55)
    for std, s in summary.items():
        print(f"  {std:25s} {s['n']:5d} {s['N']:5d} {s['alpha']:7.3f} {s['mean_iic']:9.3f}")

    return summary


def run_efa(df):
    """EFA on the full NOF item set."""
    print("\n\n" + "=" * 80)
    print("  PART 2: EFA ON FULL NOF QUESTION SET (2024 data)")
    print("  (minres extraction, promax rotation)")
    print("=" * 80)

    efa_items = [
        # Violence (Q13a+b+c combined = use all 3)
        "violence_patients", "violence_managers", "violence_colleagues",
        # Harassment (Q14a/b/c)
        "harassment_patients", "harassment_managers", "harassment_colleagues",
        # Racism
        "fair_career_progression", "discrim_patients", "discrim_colleagues",
        "org_respects_differences",
        # Sexual safety
        "sexual_behaviour_patients", "sexual_behaviour_staff",
        # Flexible working (with Q6b)
        "satisfaction_flexible_working", "org_committed_wlb", "manager_flexible_working",
        # Line management
        "mgr_interest_wellbeing", "mgr_understanding_problems", "mgr_effective_action",
        # Team
        "team_shared_objectives", "feel_valued_by_team",
        # Appraisal
        "had_appraisal", "appraisal_clear_objectives", "appraisal_felt_valued",
        # Development
        "supported_develop_potential", "access_learning_dev",
        # Supportive environment
        "nutritious_food", "org_positive_wellbeing",
    ]

    available = [i for i in efa_items if i in df.columns]
    missing = [i for i in efa_items if i not in df.columns]
    if missing:
        print(f"\n  WARNING: Missing items: {missing}")

    dat = df[available].dropna()
    n_obs, n_vars = dat.shape
    print(f"\n  N = {n_obs} organisations, {n_vars} items")

    try:
        kmo_per, kmo_total = calculate_kmo(dat)
        print(f"  KMO: {kmo_total:.3f}")
        bad_kmo = [(c, v) for c, v in zip(dat.columns, kmo_per) if v < 0.5]
        if bad_kmo:
            print(f"  Items with KMO < 0.5:")
            for c, v in bad_kmo:
                print(f"    {c}: {v:.3f}")
    except Exception as e:
        print(f"  KMO failed: {e}")

    # Theoretical groupings for display
    std_groups = {
        "Violence": ["violence_patients", "violence_managers", "violence_colleagues"],
        "Harassment": ["harassment_patients", "harassment_managers", "harassment_colleagues"],
        "Racism": ["fair_career_progression", "discrim_patients", "discrim_colleagues",
                    "org_respects_differences"],
        "SexSafety": ["sexual_behaviour_patients", "sexual_behaviour_staff"],
        "FlexWork": ["satisfaction_flexible_working", "org_committed_wlb",
                     "manager_flexible_working"],
        "LineMgmt": ["mgr_interest_wellbeing", "mgr_understanding_problems",
                     "mgr_effective_action"],
        "Team": ["team_shared_objectives", "feel_valued_by_team"],
        "Appraisal": ["had_appraisal", "appraisal_clear_objectives", "appraisal_felt_valued"],
        "Develop": ["supported_develop_potential", "access_learning_dev"],
        "SuppEnv": ["nutritious_food", "org_positive_wellbeing"],
    }

    for n_factors in [6, 8, 10]:
        print(f"\n  --- {n_factors}-factor EFA ---")

        fa = FactorAnalyzer(n_factors=n_factors, rotation="promax", method="minres")
        fa.fit(dat)

        loadings = pd.DataFrame(
            fa.loadings_, index=available,
            columns=[f"F{i+1}" for i in range(n_factors)]
        )

        _, _, cum_var = fa.get_factor_variance()
        print(f"  Variance explained: {cum_var[-1]:.1%}")

        comm = pd.Series(fa.get_communalities(), index=available)
        heywood = comm[comm > 1.0]
        if len(heywood) > 0:
            print(f"  HEYWOOD CASES: {list(heywood.index)}")

        # Print loadings grouped by standard
        print(f"\n  {'Item':40s}", end="")
        for c in loadings.columns:
            print(f" {c:>6s}", end="")
        print(f"  {'Primary':>8s}")
        print("  " + "-" * (42 + 7 * n_factors + 10))

        for std, std_items in std_groups.items():
            for item in std_items:
                if item in loadings.index:
                    row = loadings.loc[item]
                    primary = row.abs().idxmax()
                    primary_val = row[primary]
                    print(f"  {item:40s}", end="")
                    for v in row:
                        if abs(v) >= 0.30:
                            print(f" {v:6.3f}", end="")
                        else:
                            print(f"   .  ", end="")
                    print(f"  {primary:>4s} ({primary_val:+.3f})")
            print()

        # Factor correlations
        if fa.phi_ is not None:
            phi = pd.DataFrame(
                fa.phi_,
                index=[f"F{i+1}" for i in range(n_factors)],
                columns=[f"F{i+1}" for i in range(n_factors)]
            )
            print(f"  Factor correlations:")
            print("  " + phi.round(3).to_string().replace("\n", "\n  "))


def run_discriminant(df):
    """Inter-standard composite correlations."""
    print("\n\n" + "=" * 80)
    print("  PART 3: INTER-STANDARD CORRELATIONS (2024 data)")
    print("=" * 80)

    composites = {
        "Violence": (["violence_patients", "violence_managers", "violence_colleagues",
                       "harassment_patients", "harassment_managers", "harassment_colleagues"], True),
        "Racism": (["fair_career_progression", "org_respects_differences",
                     "discrim_patients", "discrim_colleagues"], False),
        "SexSafety": (["sexual_behaviour_patients", "sexual_behaviour_staff"], True),
        "FlexWork": (["satisfaction_flexible_working", "org_committed_wlb",
                       "manager_flexible_working"], False),
        "LineMgmt": (["mgr_interest_wellbeing", "mgr_understanding_problems",
                       "mgr_effective_action"], False),
        "Team": (["team_shared_objectives", "feel_valued_by_team"], False),
        "Appraisal": (["had_appraisal", "appraisal_clear_objectives",
                        "appraisal_felt_valued"], False),
        "Develop": (["supported_develop_potential", "access_learning_dev"], False),
        "SuppEnv": (["nutritious_food", "org_positive_wellbeing"], False),
    }

    scores = pd.DataFrame(index=df.index)
    for std, (items, needs_reverse) in composites.items():
        available = [i for i in items if i in df.columns]
        if not available:
            continue
        vals = df[available].copy()
        if needs_reverse:
            # For violence/harassment/sex: reverse so higher = better
            reverse_cols = [c for c in available if any(x in c for x in
                           ['violence', 'harassment', 'sexual_behaviour', 'discrim'])]
            for rc in reverse_cols:
                vals[rc] = 1 - vals[rc]
        else:
            # For racism: reverse discrim items
            for rc in [c for c in available if 'discrim' in c]:
                vals[rc] = 1 - vals[rc]
        scores[std] = vals.mean(axis=1)

    valid = scores.dropna()
    corr = valid.corr()
    print(f"\n  Inter-standard correlations (N={len(valid)}):")
    print("  " + corr.round(3).to_string().replace("\n", "\n  "))

    mask = np.triu(np.ones(corr.shape), k=1).astype(bool)
    mean_r = corr.where(mask).stack().mean()
    print(f"\n  Mean inter-standard r: {mean_r:.3f}")

    print(f"\n  Pairs with r > 0.80:")
    stds = corr.columns.tolist()
    any_high = False
    for i in range(len(stds)):
        for j in range(i + 1, len(stds)):
            r = corr.iloc[i, j]
            if abs(r) > 0.80:
                print(f"    {stds[i]:12s} <-> {stds[j]:12s}: r = {r:.3f}")
                any_high = True
    if not any_high:
        print(f"    None")


def main():
    print("=" * 80)
    print("  FULL NOF ANALYSIS — 2024 NSS DATA ONLY")
    print("  All items extracted from NSS24 organisational results Excel")
    print("  Method: minres extraction, promax rotation (same as expanded_fa.py)")
    print("=" * 80)

    df = extract_all_nof_items()

    # Summary stats
    print(f"\n  Descriptive statistics:")
    num_cols = [c for c in df.columns if c not in ['org_id', 'org_name']]
    desc = df[num_cols].describe().T[['count', 'mean', 'std', 'min', 'max']]
    for idx, row in desc.iterrows():
        print(f"    {idx:40s}  N={row['count']:5.0f}  mean={row['mean']:.3f}  "
              f"sd={row['std']:.3f}  [{row['min']:.3f}, {row['max']:.3f}]")

    run_q6b_vs_q6c(df)
    run_reliability(df)
    run_efa(df)
    run_discriminant(df)

    print("\n\n" + "=" * 80)
    print("  DONE — all results from 2024 NSS data only")
    print("=" * 80)


if __name__ == "__main__":
    main()
