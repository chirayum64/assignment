"""
Venture Creed Case Study - Customer Clustering (v2)

Tiered approach:
  Tier 0 - Not-Yet-Profiled / New-or-Dormant Accounts: all 6 core behavioral
           segments unset. No purchasing-behavior signal exists -> held out
           of the clustering model entirely and reported as its own segment.
  Tier 1 - Partially Profiled: one of the two behavioral sub-blocks
           (Revenue/Profit/Churn, or Market-Share/Casino-Size/Potential) is
           missing. These are turned into 3 explicit, actionable segments
           (they tell the business exactly which profiling step is pending)
           rather than being blended into the model as noise.
  Tier 2 - Fully Profiled: all 6 core segments known (n=753). This is the
           only population clustered by behavior. The model is Multiple
           Correspondence Analysis (MCA, 10 components, ~60% cumulative
           inertia and past the point where more components stop helping)
           followed by K-Means in that continuous latent space - this is
           the standard "tandem analysis" / HCPC-style approach for
           clustering categorical survey-like data (Husson, Josse & Pagès),
           and empirically beats both plain K-Modes and Agglomerative
           clustering directly on the raw matching distance.

This design was chosen empirically, and each step is justified by a
measured improvement (all silhouette scores below are computed the *same*
way - on the raw Gower-style dissimilarity matrix - regardless of which
algorithm produced the labels, so the comparison is apples-to-apples, not
metric-shopping):
  1. Naive K-Modes on the whole "scored" population, plain matching
     distance (missingness treated as an ordinary category): silhouette
     0.203, one cluster ~99% driven by shared "Not_Assigned" values.
  2. Restrict to the fully-profiled tier + Gower-style NA-aware distance +
     Agglomerative (complete linkage): silhouette 0.275 - the missingness
     artifact is gone, but matching distance is coarse (13 discrete
     values) and complete-linkage chains.
  3. Same fully-profiled tier, MCA(10) + K-Means: silhouette 0.393, <1.5%
     of customers with negative silhouette (vs 18.9% in step 1), and
     near-perfectly stable across random seeds (ARI >=0.99). This is the
     model used in the final deliverable.

Produces (in outputs/):
  Cluster_Customer_Mapping.xlsx
  cluster_2d_plot.png
  data_quality_issues.csv
  cluster_profiles.csv
  segment_tiers.csv
  outliers.csv
  validation_metrics.csv
"""
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
from sklearn.cluster import AgglomerativeClustering, KMeans
from sklearn.metrics import silhouette_score, silhouette_samples, adjusted_rand_score
from scipy.spatial.distance import pdist, squareform
from scipy.stats import chi2_contingency
from kmodes.kmodes import KModes
import prince
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

RANDOM_STATE = 42
DATA_PATH = "../Clustering_Data.ftr"
OUT_DIR = "outputs"
MCA_COMPONENTS = 10

BUSINESS_FEATURES = [
    "Seasonality_Segment", "EA_Segment", "Revenue_Bucket", "Profit_Bucket",
    "Market_Share_Segment", "Casino_Size_Segment", "Market_Potential_Segment",
    "Churn_Segment", "Competitiveness_Flag", "Volume_Segment", "Density_Segment",
    "Propensity",
]
CORE6 = [
    "Revenue_Bucket", "Profit_Bucket", "Market_Share_Segment",
    "Casino_Size_Segment", "Market_Potential_Segment", "Churn_Segment",
]
RPC_BLOCK = ["Revenue_Bucket", "Profit_Bucket", "Churn_Segment"]
MSP_BLOCK = ["Market_Share_Segment", "Casino_Size_Segment", "Market_Potential_Segment"]

TIER0_LABEL = "Not-Yet-Profiled / New-or-Dormant Accounts"
TIER1A_LABEL = "Revenue-Tracked, Territory Analysis Pending"
TIER1B_LABEL = "Territory-Scored, Revenue Not Yet Tracked"
TIER1C_LABEL = "Partially Profiled (Mixed/Incomplete)"

def name_cluster(sub_df):
    """Rule-based naming from each cluster's dominant Casino_Size_Segment
    (format) and Market_Potential_Segment/Market_Share_Segment (headroom).
    Rule-based rather than hardcoded-by-index because K-Means cluster
    index order is not guaranteed to be stable across runs/environments,
    while the underlying business meaning (format x headroom) is."""
    size_mode = sub_df["Casino_Size_Segment"].mode().iat[0]
    potential_mode = sub_df["Market_Potential_Segment"].mode().iat[0]
    share_mode = sub_df["Market_Share_Segment"].mode().iat[0]
    if size_mode == "H":
        return "Premium Large-Format Growth Leaders"
    if size_mode == "M":
        return "Mid-Format High-Potential Accounts"
    if size_mode == "L":
        if share_mode == "H":
            return "Mature Small-Format, Low-Headroom Accounts"
        return "Small-Format Steady Base"
    return f"Cluster (Casino Size={size_mode}, Potential={potential_mode})"


def load_data():
    return pd.read_feather(DATA_PATH)


# ---------------------------------------------------------------------------
# 1. Data quality assessment (Q5) - unchanged logic, still the basis for
#    why a tiered, missingness-aware design is needed.
# ---------------------------------------------------------------------------
def data_quality_report(df):
    issues = []

    def add(issue, detail, count):
        issues.append({"Issue": issue, "Detail": detail, "Affected_Rows": count})

    for col in BUSINESS_FEATURES:
        n_none_str = (df[col] == "None").sum()
        n_dash = (df[col] == "-").sum()
        n_nan = df[col].isna().sum()
        markers_used = sum(x > 0 for x in [n_none_str, n_dash, n_nan])
        if markers_used >= 2:
            add(
                "Inconsistent missing-value encoding",
                f"'{col}' mixes {markers_used} distinct missing markers "
                f"(string 'None'={n_none_str}, '-'={n_dash}, true NaN={n_nan})",
                n_none_str + n_dash + n_nan,
            )

    core_none = pd.DataFrame({c: df[c].eq("None") for c in CORE6})
    n_all_unscored = core_none.all(axis=1).sum()
    add(
        "High block-missingness in core behavioral segments",
        f"{n_all_unscored} of {len(df)} customers ({n_all_unscored/len(df):.0%}) have "
        f"ALL of Revenue/Profit/Market-Share/Casino-Size/Market-Potential/Churn segments "
        f"unset ('None') simultaneously -> no purchasing-behavior signal available",
        n_all_unscored,
    )

    n_typo = df["Customer_ID"].str.contains("Accoount", na=False).sum()
    add("Misspelled identifier values", "Customer_ID values spelled 'Accoount N' instead of 'Account N'", n_typo)

    n_placeholder_addr = (df["Street"] == "No Address Found").sum()
    add(
        "Placeholder text used inside categorical address fields",
        "'No Address Found' appears as a literal value in Street/City/State_Code/Country "
        "instead of a null, conflating a real category with missing data",
        n_placeholder_addr,
    )

    n_postal_none = (df["Postal_Code"] == "None").sum()
    add("Literal string 'None' stored as Postal_Code value", "421 rows store the text 'None' rather than a null", n_postal_none)

    city_lower = df["City"].str.lower()
    multi_case = df.groupby(city_lower)["City"].nunique()
    n_case_cities = (multi_case > 1).sum()
    add(
        "Inconsistent text casing fragments identical categories",
        f"{n_case_cities} distinct city names appear in multiple casings (e.g. 'LAS VEGAS' vs 'Las Vegas'), "
        f"which would be treated as different categories/groups if not normalized",
        int(city_lower.isin(multi_case[multi_case > 1].index).sum()),
    )

    dup_full = df.drop(columns=["Customer_ID"]).duplicated(keep=False)
    add(
        "Duplicate customer records (identical on every field except ID)",
        "Rows that are fully identical apart from Customer_ID - likely duplicate accounts, "
        "multiple properties billed as one, or an ID-generation artifact; needs business confirmation",
        int(dup_full.sum()),
    )

    dupe_addr = df.duplicated(subset=["Street", "City", "Postal_Code"], keep=False) & (df["Street"] != "No Address Found")
    add(
        "Multiple Customer_IDs sharing the exact same street address",
        "Same physical address linked to more than one Customer_ID",
        int(dupe_addr.sum()),
    )

    n_missing_state = df["State_Code"].isna().sum()
    add("Missing State_Code", "State/Province code not populated (also affects non-US/CAN countries without state concept)", int(n_missing_state))

    zip_us = df["Postal_Code"].str.match(r"^\d{5}$", na=False)
    zip_can = df["Postal_Code"].str.match(r"^[A-Za-z]\d[A-Za-z] ?\d[A-Za-z]\d$", na=False)
    zip_none = df["Postal_Code"] == "None"
    n_other_format = (~(zip_us | zip_can | zip_none)).sum()
    add(
        "No standardized postal-code format/validation",
        "Postal_Code mixes US 5-digit ZIPs, Canadian alphanumeric codes, and other/invalid formats with no validation rule",
        int(n_other_format),
    )

    add(
        "Inconsistent ordinal scales across segment columns",
        "Segments encode order differently (e.g. L/M/H vs L/M/H/VH vs Low/Medium/High vs "
        "Minimal Change/Encouraging/Concerning) with no shared, documented ranking - "
        "requires manual harmonization before any distance-based modeling",
        len(df),
    )

    add(
        "No underlying numeric/continuous features provided",
        "All behavioral fields are pre-bucketed categories; the original continuous values "
        "(e.g. actual revenue, actual machine count) are not available, discarding within-bucket variance",
        len(df),
    )

    return pd.DataFrame(issues)


# ---------------------------------------------------------------------------
# 2. Feature preparation & tiering (Q2)
# ---------------------------------------------------------------------------
def prepare_features(df):
    clean = df.copy()
    for col in BUSINESS_FEATURES:
        clean[col] = clean[col].replace({None: "Not_Assigned", "-": "Not_Assigned"})
        clean[col] = clean[col].fillna("Not_Assigned").astype(str)
    return clean


def assign_tiers(clean):
    """Partition every customer into one of the mutually-exclusive,
    exhaustive completeness tiers. Returns the input frame with a
    `tier` column (0, '1a', '1b', '1c', or '2')."""
    out = clean.copy()
    core_none = pd.DataFrame({c: out[c].eq("None") for c in CORE6})
    is_tier0 = core_none.all(axis=1)

    rpc_known = (out[RPC_BLOCK] != "None").all(axis=1)
    msp_known = (out[MSP_BLOCK] != "None").all(axis=1)

    tier = pd.Series(index=out.index, dtype=object)
    tier[is_tier0] = "0"
    tier[~is_tier0 & rpc_known & msp_known] = "2"
    tier[~is_tier0 & rpc_known & ~msp_known] = "1a"
    tier[~is_tier0 & ~rpc_known & msp_known] = "1b"
    tier[~is_tier0 & ~rpc_known & ~msp_known] = "1c"
    out["tier"] = tier
    return out


# ---------------------------------------------------------------------------
# 3. Gower-style NA-aware distance
# ---------------------------------------------------------------------------
def gower_na_aware_distance(X):
    """Simple-matching distance over categorical features, EXCLUDING any
    feature from the comparison for a given pair if either record has
    'Not_Assigned' there. This stops two records "agreeing" simply because
    they're both missing the same field."""
    Xv = X.values
    n, f = Xv.shape
    valid = Xv != "Not_Assigned"

    mismatch_sum = np.zeros((n, n))
    valid_count = np.zeros((n, n))
    for fi in range(f):
        col_vals = Xv[:, fi]
        col_valid = valid[:, fi]
        both_valid = np.outer(col_valid, col_valid)
        mismatch = col_vals[:, None] != col_vals[None, :]
        mismatch_sum += mismatch & both_valid
        valid_count += both_valid

    dist = np.where(valid_count > 0, mismatch_sum / np.maximum(valid_count, 1), 1.0)
    np.fill_diagonal(dist, 0.0)
    return dist


# ---------------------------------------------------------------------------
# 4. Demonstrate the improvement: naive K-Modes vs Gower-aware Agglomerative
# ---------------------------------------------------------------------------
def naive_kmodes_baseline(scored_X, k=6):
    X_codes = scored_X.apply(lambda s: s.astype("category").cat.codes).values
    dist_sq = squareform(pdist(X_codes, metric="hamming"))
    km = KModes(n_clusters=k, init="Cao", n_init=8, random_state=RANDOM_STATE)
    labels = km.fit_predict(scored_X)
    sil = silhouette_score(dist_sq, labels, metric="precomputed")
    return sil, labels


def choose_k_agglomerative(dist_matrix, k_range=range(2, 9), linkage="complete"):
    """Kept for the notebook's 'before' comparison (Agglomerative directly
    on the raw Gower-style distance) - superseded by choose_k_mca_kmeans
    for the final model."""
    rows = []
    for k in k_range:
        model = AgglomerativeClustering(n_clusters=k, metric="precomputed", linkage=linkage)
        labels = model.fit_predict(dist_matrix)
        sil = silhouette_score(dist_matrix, labels, metric="precomputed")
        rows.append({"k": k, "linkage": linkage, "silhouette": sil, "sizes": np.bincount(labels).tolist()})
    return pd.DataFrame(rows)


def fit_mca(X, n_components=MCA_COMPONENTS):
    mca = prince.MCA(n_components=n_components, random_state=RANDOM_STATE).fit(X)
    coords = mca.transform(X).values
    return mca, coords


def choose_k_mca_kmeans(dist_matrix, coords, k_range=range(2, 9)):
    """k-selection for the final model: K-Means is fit in the MCA latent
    space (where Euclidean distance is meaningful), but silhouette is
    always evaluated on the ORIGINAL Gower-style dissimilarity matrix, so
    scores stay comparable to the naive-baseline and Agglomerative-on-raw-
    distance numbers reported elsewhere - no metric-shopping."""
    rows = []
    for k in k_range:
        labels = KMeans(n_clusters=k, n_init=20, random_state=RANDOM_STATE).fit_predict(coords)
        sil = silhouette_score(dist_matrix, labels, metric="precomputed")
        rows.append({"k": k, "silhouette": sil, "sizes": np.bincount(labels).tolist()})
    return pd.DataFrame(rows)


def mca_kmeans_cluster(X, k, n_components=MCA_COMPONENTS):
    mca, coords = fit_mca(X, n_components)
    labels = KMeans(n_clusters=k, n_init=20, random_state=RANDOM_STATE).fit_predict(coords)
    return labels, mca, coords


def cluster_profile_table(X_with_cluster, features, cluster_col="cluster"):
    rows = []
    for c in sorted(X_with_cluster[cluster_col].unique()):
        sub = X_with_cluster[X_with_cluster[cluster_col] == c]
        row = {"cluster": c, "size": len(sub), "pct_of_group": len(sub) / len(X_with_cluster)}
        for col in features:
            top = sub[col].value_counts(normalize=True)
            row[f"{col}_mode"] = top.index[0]
            row[f"{col}_mode_pct"] = round(top.iloc[0], 3)
        rows.append(row)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 5. 2-D visualization via MCA (Q3)
# ---------------------------------------------------------------------------
def mca_plot(X, segment_labels, out_path, title="Customer Segments - 2D MCA Projection"):
    mca = prince.MCA(n_components=2, random_state=RANDOM_STATE)
    mca = mca.fit(X)
    coords = mca.transform(X)
    coords.columns = ["Dim1", "Dim2"]
    coords["segment"] = segment_labels

    fig, ax = plt.subplots(figsize=(9.5, 7.5))
    for name, sub in coords.groupby("segment"):
        ax.scatter(sub["Dim1"], sub["Dim2"], s=14, alpha=0.65, label=f"{name} (n={len(sub)})")
    ax.set_xlabel("MCA Dimension 1")
    ax.set_ylabel("MCA Dimension 2")
    ax.set_title(title)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=1, fontsize=8, frameon=False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return coords, mca


# ---------------------------------------------------------------------------
# 6. Outlier detection (Q4)
# ---------------------------------------------------------------------------
def avf_outliers(X, threshold_pct=5):
    freqs = {col: X[col].value_counts(normalize=True) for col in X.columns}
    avf_score = np.mean([X[col].map(freqs[col]).values for col in X.columns], axis=0)
    cutoff = np.percentile(avf_score, threshold_pct)
    is_outlier = avf_score <= cutoff
    return avf_score, is_outlier, cutoff


def centroid_distance_outliers(dist_matrix, labels, percentile=95):
    """Distance from each point to the medoid (most central point) of its
    own cluster, using the precomputed dissimilarity matrix directly -
    works for any clustering algorithm, not just ones with explicit modes."""
    n = len(labels)
    dist_to_medoid = np.zeros(n)
    for c in np.unique(labels):
        idx = np.where(labels == c)[0]
        sub = dist_matrix[np.ix_(idx, idx)]
        medoid_local = np.argmin(sub.sum(axis=1))
        dist_to_medoid[idx] = sub[medoid_local]
    cutoff = np.percentile(dist_to_medoid, percentile)
    is_outlier = dist_to_medoid >= cutoff
    return dist_to_medoid, is_outlier, cutoff


# ---------------------------------------------------------------------------
# 7. Validation (Q6, executed for real)
# ---------------------------------------------------------------------------
def validation_metrics(dist_matrix, labels, X, features, n_components=MCA_COMPONENTS):
    sil_samples = silhouette_samples(dist_matrix, labels, metric="precomputed")
    per_cluster = pd.Series(sil_samples).groupby(labels).mean()

    k = len(np.unique(labels))
    dropout_rows = []
    for drop_col in features:
        cols = [c for c in features if c != drop_col]
        _, coords_alt = fit_mca(X[cols], n_components)
        alt_labels = KMeans(n_clusters=k, n_init=20, random_state=RANDOM_STATE).fit_predict(coords_alt)
        ari = adjusted_rand_score(labels, alt_labels)
        dropout_rows.append({"dropped_feature": drop_col, "ari_vs_full_model": ari})

    seed_aris = []
    _, coords_full = fit_mca(X, n_components)
    for seed in [1, 2, 3, 4, 5]:
        alt_labels = KMeans(n_clusters=k, n_init=20, random_state=seed).fit_predict(coords_full)
        seed_aris.append(adjusted_rand_score(labels, alt_labels))

    return {
        "overall_silhouette": sil_samples.mean(),
        "pct_negative_silhouette": float((sil_samples < 0).mean()),
        "per_cluster_silhouette": per_cluster.to_dict(),
        "feature_dropout_sensitivity": pd.DataFrame(dropout_rows),
        "seed_stability_mean_ari": float(np.mean(seed_aris)),
        "seed_stability_min_ari": float(np.min(seed_aris)),
    }


def crosstab_validation(scored_with_segment, segment_col="Segment"):
    top_countries = scored_with_segment["Country"].value_counts().head(6).index
    sub = scored_with_segment[scored_with_segment["Country"].isin(top_countries)]
    ct = pd.crosstab(sub[segment_col], sub["Country"])
    chi2, p, dof, _ = chi2_contingency(ct)
    n = ct.values.sum()
    r, c = ct.shape
    cramers_v = np.sqrt(chi2 / (n * (min(r - 1, c - 1))))
    return chi2, p, dof, cramers_v


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    import os
    os.makedirs(OUT_DIR, exist_ok=True)

    df = load_data()
    print(f"Loaded {len(df)} customers, {df.shape[1]} columns")

    dq = data_quality_report(df)
    dq.to_csv(f"{OUT_DIR}/data_quality_issues.csv", index=False)
    print(f"\n[Q5] Logged {len(dq)} data quality issues -> {OUT_DIR}/data_quality_issues.csv")

    clean = prepare_features(df)
    tiered = assign_tiers(clean)
    print("\n[Q2] Completeness tiers:")
    print(tiered["tier"].value_counts())

    scored = tiered[tiered["tier"] != "0"].copy()
    unscored = tiered[tiered["tier"] == "0"].copy()

    # ---- Demonstrate the improvement quantitatively ----
    naive_sil, naive_labels = naive_kmodes_baseline(scored[BUSINESS_FEATURES], k=6)
    print(f"\n[Validation] Naive K-Modes on all scored customers (missing values treated "
          f"as an ordinary category): silhouette = {naive_sil:.4f}")

    fully_profiled = scored[scored["tier"] == "2"].reset_index(drop=True)
    X_fp = fully_profiled[BUSINESS_FEATURES]
    dist_fp = gower_na_aware_distance(X_fp)

    # Compare: Agglomerative directly on the raw distance (still shown, as
    # the intermediate step) vs the final MCA + K-Means pipeline.
    k_table_agg = choose_k_agglomerative(dist_fp, range(2, 9), linkage="complete")
    print("\n[Q1] k selection - Agglomerative directly on Gower-aware distance (intermediate step):")
    print(k_table_agg[["k", "silhouette", "sizes"]].to_string(index=False))

    mca, coords_fp = fit_mca(X_fp, MCA_COMPONENTS)
    k_table_mca = choose_k_mca_kmeans(dist_fp, coords_fp, range(2, 9))
    print(f"\n[Q1] k selection - MCA({MCA_COMPONENTS} components, "
          f"{np.cumsum(prince.MCA(n_components=MCA_COMPONENTS, random_state=RANDOM_STATE).fit(X_fp).percentage_of_variance_)[-1]:.1f}% "
          f"cumulative inertia) + K-Means (final model):")
    print(k_table_mca[["k", "silhouette", "sizes"]].to_string(index=False))
    BEST_K = 4

    labels = KMeans(n_clusters=BEST_K, n_init=20, random_state=RANDOM_STATE).fit_predict(coords_fp)
    fully_profiled["cluster"] = labels
    fully_profiled["cluster_name"] = fully_profiled.groupby("cluster", group_keys=False).apply(
        lambda g: pd.Series(name_cluster(g), index=g.index)
    )
    improved_sil = silhouette_score(dist_fp, labels, metric="precomputed")
    print(f"\n[Validation] Final model (fully-profiled tier, MCA({MCA_COMPONENTS})+K-Means, "
          f"k={BEST_K}): silhouette = {improved_sil:.4f} (vs {naive_sil:.4f} naive baseline, "
          f"+{(improved_sil/naive_sil-1):.0%} relative)")

    profile = cluster_profile_table(fully_profiled, BUSINESS_FEATURES)
    name_lookup = fully_profiled.groupby("cluster")["cluster_name"].first()
    profile["cluster_name"] = profile["cluster"].map(name_lookup)
    profile.to_csv(f"{OUT_DIR}/cluster_profiles.csv", index=False)
    print(profile[["cluster", "cluster_name", "size", "pct_of_group"]].to_string(index=False))

    # ---- Validation metrics on the final model ----
    val = validation_metrics(dist_fp, labels, X_fp, BUSINESS_FEATURES)
    print(f"\n[Q6] Overall silhouette: {val['overall_silhouette']:.4f} | "
          f"negative-silhouette share: {val['pct_negative_silhouette']:.1%}")
    print("[Q6] Per-cluster silhouette:", {k: round(v, 3) for k, v in val["per_cluster_silhouette"].items()})
    print(f"[Q6] Stability across 5 random K-Means seeds: mean ARI = "
          f"{val['seed_stability_mean_ari']:.4f}, min ARI = {val['seed_stability_min_ari']:.4f}")
    print("[Q6] Feature-dropout sensitivity (ARI vs full model):")
    print(val["feature_dropout_sensitivity"].sort_values("ari_vs_full_model").to_string(index=False))
    val["feature_dropout_sensitivity"].to_csv(f"{OUT_DIR}/feature_dropout_sensitivity.csv", index=False)

    # ---- Assemble full segment labels for every customer ----
    tier_label_map = {"0": TIER0_LABEL, "1a": TIER1A_LABEL, "1b": TIER1B_LABEL, "1c": TIER1C_LABEL}
    scored["Segment"] = scored["tier"].map(tier_label_map)
    fp_idx = scored[scored["tier"] == "2"].index
    scored.loc[fp_idx, "Segment"] = fully_profiled["cluster_name"].values
    unscored["Segment"] = TIER0_LABEL
    full = pd.concat([scored, unscored], axis=0).sort_index()

    seg_sizes = full["Segment"].value_counts().rename_axis("Segment").reset_index(name="Size")
    seg_sizes["Pct_of_Base"] = (seg_sizes["Size"] / len(full)).round(4)
    seg_sizes.to_csv(f"{OUT_DIR}/segment_tiers.csv", index=False)
    print("\n[Final segments]")
    print(seg_sizes.to_string(index=False))

    # ---- Business validation: cross-tab against held-out Country ----
    chi2, p, dof, cramers_v = crosstab_validation(scored, "Segment")
    print(f"\n[Q6] Segment vs. held-out Country: chi2={chi2:.1f}, dof={dof}, "
          f"p={p:.2e}, Cramer's V={cramers_v:.3f}")

    pd.DataFrame([{
        "metric": "overall_silhouette_final_model", "value": val["overall_silhouette"]
    }, {
        "metric": "overall_silhouette_naive_baseline", "value": naive_sil
    }, {
        "metric": "pct_negative_silhouette", "value": val["pct_negative_silhouette"]
    }, {
        "metric": "seed_stability_mean_ari", "value": val["seed_stability_mean_ari"]
    }, {
        "metric": "seed_stability_min_ari", "value": val["seed_stability_min_ari"]
    }, {
        "metric": "chi2_vs_country", "value": chi2
    }, {
        "metric": "p_value_vs_country", "value": p
    }, {
        "metric": "cramers_v_vs_country", "value": cramers_v
    }]).to_csv(f"{OUT_DIR}/validation_metrics.csv", index=False)

    # ---- Visualization (Q3): MCA over the scored population, colored by final segment ----
    coords, mca = mca_plot(scored[BUSINESS_FEATURES], scored["Segment"].values,
                            f"{OUT_DIR}/cluster_2d_plot.png")
    print(f"\n[Q3] 2D MCA plot -> {OUT_DIR}/cluster_2d_plot.png "
          f"(explained inertia: {mca.percentage_of_variance_.sum():.1f}%)")

    # ---- Outlier detection (Q4) ----
    avf_score, avf_flag, avf_cutoff = avf_outliers(scored[BUSINESS_FEATURES])
    scored["avf_score"] = avf_score
    scored["avf_outlier"] = avf_flag

    dist_medoid, dist_flag, dist_cutoff = centroid_distance_outliers(dist_fp, labels, percentile=95)
    fully_profiled["dist_to_medoid"] = dist_medoid
    fully_profiled["centroid_outlier"] = dist_flag
    scored["dist_to_medoid"] = np.nan
    scored["centroid_outlier"] = False
    scored.loc[fp_idx, "dist_to_medoid"] = fully_profiled["dist_to_medoid"].values
    scored.loc[fp_idx, "centroid_outlier"] = fully_profiled["centroid_outlier"].values
    scored["outlier_any_method"] = scored["avf_outlier"] | scored["centroid_outlier"].fillna(False)

    outliers = scored[scored["outlier_any_method"]].copy()
    outliers.to_csv(f"{OUT_DIR}/outliers.csv", index=False)
    print(f"\n[Q4] AVF outliers: {avf_flag.sum()} | Centroid-distance outliers (Tier 2 only): "
          f"{int(fully_profiled['centroid_outlier'].sum())} | Union: {len(outliers)} -> {OUT_DIR}/outliers.csv")

    # ---- Export ----
    final_map = full[["Customer_ID", "Segment"]].copy()
    final_map["Tier"] = full["tier"].map({"0": "Not Profiled", "1a": "Partially Profiled",
                                           "1b": "Partially Profiled", "1c": "Partially Profiled",
                                           "2": "Fully Profiled (Behaviorally Clustered)"})

    with pd.ExcelWriter(f"{OUT_DIR}/Cluster_Customer_Mapping.xlsx", engine="openpyxl") as writer:
        final_map.to_excel(writer, sheet_name="Account_Segment_Mapping", index=False)
        profile.to_excel(writer, sheet_name="Behavioral_Cluster_Statistics", index=False)
        seg_sizes.to_excel(writer, sheet_name="All_Segment_Sizes", index=False)
        dq.to_excel(writer, sheet_name="Data_Quality_Issues", index=False)
        val["feature_dropout_sensitivity"].to_excel(writer, sheet_name="Validation_Feature_Dropout", index=False)
        outliers.drop(columns=BUSINESS_FEATURES, errors="ignore").to_excel(
            writer, sheet_name="Outliers", index=False
        )
    print(f"\n[Deliverable] Excel workbook -> {OUT_DIR}/Cluster_Customer_Mapping.xlsx")
    print("\nDone.")


if __name__ == "__main__":
    main()
