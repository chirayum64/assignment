"""
Venture Creed Case Study - Customer Clustering
End-to-end analysis: data quality checks, feature selection, K-Modes clustering,
2-D MCA visualization, and categorical outlier detection.

Produces:
  analysis/outputs/Cluster_Customer_Mapping.xlsx
  analysis/outputs/cluster_2d_plot.png
  analysis/outputs/data_quality_issues.csv
  analysis/outputs/cluster_profiles.csv
  analysis/outputs/outliers.csv
"""
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
from scipy.spatial.distance import pdist, squareform
from sklearn.metrics import silhouette_score
from kmodes.kmodes import KModes
import prince
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

RANDOM_STATE = 42
DATA_PATH = "../Clustering_Data.ftr"
OUT_DIR = "outputs"

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

CLUSTER_NAMES = {
    0: "Steady Small-Format Core",
    1: "Under-Profiled Moderate Accounts",
    2: "Emerging High-Potential Adopters",
    3: "Premium Growth Leaders",
    4: "Mature Low-Headroom Small Accounts",
    5: "High-Density Competitive, Thin-Profile Accounts",
}
UNSCORED_LABEL = "Not-Yet-Profiled / New-or-Dormant Accounts"


def load_data():
    df = pd.read_feather(DATA_PATH)
    return df


# ---------------------------------------------------------------------------
# 1. Data quality assessment (Q5)
# ---------------------------------------------------------------------------
def data_quality_report(df):
    issues = []

    def add(issue, detail, count):
        issues.append({"Issue": issue, "Detail": detail, "Affected_Rows": count})

    # Inconsistent null encodings
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

    n_inconsistent_scales = 5
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
# 2. Feature preparation (Q2)
# ---------------------------------------------------------------------------
def prepare_features(df):
    clean = df.copy()
    for col in BUSINESS_FEATURES:
        clean[col] = clean[col].replace({None: "Not_Assigned", "-": "Not_Assigned"})
        clean[col] = clean[col].fillna("Not_Assigned").astype(str)
    return clean


def split_scored_unscored(clean):
    core_none = pd.DataFrame({c: clean[c].eq("None") for c in CORE6})
    is_unscored = core_none.all(axis=1)
    return clean.loc[~is_unscored].copy(), clean.loc[is_unscored].copy()


# ---------------------------------------------------------------------------
# 3. K-Modes clustering + k selection (Q1)
# ---------------------------------------------------------------------------
def choose_k(X, k_range=range(2, 9)):
    X_codes = X.apply(lambda s: s.astype("category").cat.codes).values
    dist_sq = squareform(pdist(X_codes, metric="hamming"))
    rows = []
    for k in k_range:
        km = KModes(n_clusters=k, init="Cao", n_init=8, random_state=RANDOM_STATE)
        labels = km.fit_predict(X)
        sil = silhouette_score(dist_sq, labels, metric="precomputed")
        rows.append({"k": k, "cost": km.cost_, "silhouette": sil})
    return pd.DataFrame(rows)


def fit_kmodes(X, k):
    km = KModes(n_clusters=k, init="Cao", n_init=15, random_state=RANDOM_STATE)
    labels = km.fit_predict(X)
    return km, labels


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
# 4. 2-D visualization via MCA (Q3)
# ---------------------------------------------------------------------------
def mca_plot(X, cluster_labels, group_names, out_path):
    mca = prince.MCA(n_components=2, random_state=RANDOM_STATE)
    mca = mca.fit(X)
    coords = mca.transform(X)
    coords.columns = ["Dim1", "Dim2"]
    coords["cluster"] = cluster_labels
    coords["name"] = coords["cluster"].map(group_names)

    fig, ax = plt.subplots(figsize=(9, 7))
    for name, sub in coords.groupby("name"):
        ax.scatter(sub["Dim1"], sub["Dim2"], s=14, alpha=0.6, label=f"{name} (n={len(sub)})")
    ax.set_xlabel("MCA Dimension 1")
    ax.set_ylabel("MCA Dimension 2")
    ax.set_title("Customer Clusters - 2D MCA Projection")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=1, fontsize=8, frameon=False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return coords, mca


# ---------------------------------------------------------------------------
# 5. Outlier detection (Q4)
# ---------------------------------------------------------------------------
def avf_outliers(X, threshold_pct=5):
    freqs = {col: X[col].value_counts(normalize=True) for col in X.columns}
    avf_score = np.mean([X[col].map(freqs[col]).values for col in X.columns], axis=0)
    cutoff = np.percentile(avf_score, threshold_pct)
    is_outlier = avf_score <= cutoff
    return avf_score, is_outlier, cutoff


def cluster_distance_outliers(X, labels, km_model, features, percentile=95):
    # Hamming distance to the assigned cluster's mode-vector is bounded and
    # takes only len(features)+1 discrete values, so a Tukey IQR fence rarely
    # exceeds the observed max (it never triggers here). A percentile cutoff
    # is the appropriate rule for a bounded, discrete dissimilarity score.
    modes = km_model.cluster_centroids_
    X_arr = X[features].values
    dist_to_own_centroid = np.array([
        np.mean(X_arr[i] != modes[labels[i]]) for i in range(len(X_arr))
    ])
    cutoff = np.percentile(dist_to_own_centroid, percentile)
    is_outlier = dist_to_own_centroid >= cutoff
    return dist_to_own_centroid, is_outlier, cutoff


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
    scored, unscored = split_scored_unscored(clean)
    print(f"\n[Feature prep] Scored (behaviorally profiled): {len(scored)} | "
          f"Not-yet-profiled: {len(unscored)}")

    X = scored[BUSINESS_FEATURES]
    k_table = choose_k(X)
    print("\n[Q1] k selection (cost + silhouette on Hamming distance):")
    print(k_table.to_string(index=False))
    best_k = 6

    km, labels = fit_kmodes(X, best_k)
    scored = scored.copy()
    scored["cluster"] = labels
    scored["cluster_name"] = scored["cluster"].map(CLUSTER_NAMES)

    profile = cluster_profile_table(scored, BUSINESS_FEATURES)
    profile["cluster_name"] = profile["cluster"].map(CLUSTER_NAMES)
    profile.to_csv(f"{OUT_DIR}/cluster_profiles.csv", index=False)
    print(f"\n[Q1] Cluster profile table -> {OUT_DIR}/cluster_profiles.csv")
    print(profile[["cluster", "cluster_name", "size", "pct_of_group"]].to_string(index=False))

    unscored = unscored.copy()
    unscored["cluster"] = -1
    unscored["cluster_name"] = UNSCORED_LABEL

    full = pd.concat([scored, unscored], axis=0).sort_index()
    group_names = {**CLUSTER_NAMES, -1: UNSCORED_LABEL}

    coords, mca = mca_plot(X, scored["cluster"].values, CLUSTER_NAMES, f"{OUT_DIR}/cluster_2d_plot.png")
    print(f"\n[Q3] 2D MCA plot -> {OUT_DIR}/cluster_2d_plot.png "
          f"(explained inertia: {mca.percentage_of_variance_.sum():.1f}% in 2 dims)")

    avf_score, avf_flag, avf_cutoff = avf_outliers(scored[BUSINESS_FEATURES])
    dist_score, dist_flag, dist_cutoff = cluster_distance_outliers(scored, labels, km, BUSINESS_FEATURES)
    scored["avf_score"] = avf_score
    scored["avf_outlier"] = avf_flag
    scored["dist_to_centroid"] = dist_score
    scored["centroid_outlier"] = dist_flag
    scored["outlier_any_method"] = avf_flag | dist_flag

    outliers = scored[scored["outlier_any_method"]].copy()
    outliers.to_csv(f"{OUT_DIR}/outliers.csv", index=False)
    print(f"\n[Q4] AVF outliers: {avf_flag.sum()} (score <= {avf_cutoff:.3f}) | "
          f"Centroid-distance outliers: {dist_flag.sum()} (dist > {dist_cutoff:.3f}) | "
          f"Union: {len(outliers)} -> {OUT_DIR}/outliers.csv")

    final_map = full[["Customer_ID", "cluster", "cluster_name"]].rename(
        columns={"cluster": "Cluster_ID", "cluster_name": "Cluster_Name"}
    )

    with pd.ExcelWriter(f"{OUT_DIR}/Cluster_Customer_Mapping.xlsx", engine="openpyxl") as writer:
        final_map.to_excel(writer, sheet_name="Account_Cluster_Mapping", index=False)
        profile.to_excel(writer, sheet_name="Cluster_Statistics", index=False)
        pd.DataFrame({
            "Group": list(group_names.values()),
            "Cluster_ID": list(group_names.keys()),
            "Size": [full[full["cluster_name"] == n].shape[0] for n in group_names.values()],
        }).drop_duplicates().to_excel(writer, sheet_name="Group_Sizes", index=False)
        dq.to_excel(writer, sheet_name="Data_Quality_Issues", index=False)
        outliers.drop(columns=BUSINESS_FEATURES, errors="ignore").to_excel(
            writer, sheet_name="Outliers", index=False
        )
    print(f"\n[Deliverable] Excel workbook -> {OUT_DIR}/Cluster_Customer_Mapping.xlsx")

    print("\nDone.")


if __name__ == "__main__":
    main()
