# Venture Creed Case Study — Customer Clustering

Solution to the ProcDNA "Venture Creed" data science case study: cluster Venture
Creed's 3,030 customers by purchasing behavior so marketing strategy can be
tailored per cluster.

## Deliverables

| File | Answers |
|---|---|
| `Venture_Creed_Case_Study_Answers.pptx` | Written answers to Q1, Q2, Q4, Q5, Q6, plus the 2-D cluster plot (Q3) |
| `Cluster_Customer_Mapping.xlsx` | Account → cluster mapping, per-cluster statistics, data quality issue log, and flagged outliers |
| `analysis/clustering_analysis.ipynb` | Full, executed analysis: EDA, data quality checks, feature selection, K-Modes clustering with k-selection, MCA visualization, and outlier detection code (Q1–Q4 code) |
| `analysis/outputs/` | Generated artifacts (plot image, CSVs) referenced by the notebook and slides |

## Reproducing the analysis

```bash
pip install -r requirements.txt
cd analysis
python run_analysis.py                 # regenerates analysis/outputs/*
jupyter nbconvert --to notebook --execute --inplace clustering_analysis.ipynb
python build_pptx.py                    # regenerates the answer deck
```

## Approach summary

- All 12 relevant customer features are categorical, so clustering uses
  **K-Modes** (matching-dissimilarity based), not K-Means, and **MCA**
  (Multiple Correspondence Analysis) rather than PCA for the 2-D plot.
- ~59% of customers have no purchasing-behavior segments computed at all
  (all 6 core segments unset simultaneously) — these are split out as a
  separate "Not-Yet-Profiled / New-or-Dormant Accounts" business segment
  rather than being forced through the clustering model, which otherwise
  produces a dominant cluster driven purely by missing data rather than
  genuine behavior.
- The remaining ~1,241 behaviorally-profiled customers are split into
  **6 clusters** (k chosen via K-Modes cost + silhouette score), each named
  and profiled for marketing use.
- Full reasoning for every decision (feature selection, k choice, outlier
  method, data quality findings, validation checklist) is in the notebook
  and the answer deck.
