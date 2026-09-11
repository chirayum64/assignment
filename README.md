# Venture Creed Clustering

ProcDNA case study submission.

Files:
- `Venture_Creed_Case_Study_Answers.pptx` - answers to Q1, Q2, Q4, Q5, Q6, plus the cluster plot for Q3
- `Cluster_Customer_Mapping.xlsx` - account to segment mapping, cluster stats, data quality log, outliers
- `analysis/clustering_analysis.ipynb` - full analysis and code

To rerun:

```
pip install -r requirements.txt
cd analysis
python run_analysis.py
jupyter nbconvert --to notebook --execute --inplace clustering_analysis.ipynb
python build_pptx.py
```

Short version of the approach: all the usable features are categorical, so
clustering is done via MCA + K-Means instead of plain K-Means on raw codes.
About 59% of customers have none of the 6 core behavior segments filled in
at all, so those are pulled out as their own segment instead of being
clustered on missing data. The remaining ~753 fully-profiled customers are
split into 4 clusters. Details and reasoning are in the notebook and the
slides.
