"""Builds the Venture_Creed_Case_Study_Answers.pptx deliverable answering Q1-Q6."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
import pandas as pd

NAVY = RGBColor(0x1F, 0x2A, 0x44)
ACCENT = RGBColor(0xC0, 0x39, 0x2B)
GREY = RGBColor(0x55, 0x55, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x1E, 0x7D, 0x32)

prof = pd.read_csv("outputs/cluster_profiles.csv")
dq = pd.read_csv("outputs/data_quality_issues.csv")
seg = pd.read_csv("outputs/segment_tiers.csv")
val = pd.read_csv("outputs/validation_metrics.csv").set_index("metric")["value"]
drop = pd.read_csv("outputs/feature_dropout_sensitivity.csv").sort_values("ari_vs_full_model")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def add_slide():
    return prs.slides.add_slide(BLANK)


def add_title(slide, text, subtitle=None):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.9))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(27)
    r.font.bold = True
    r.font.color.rgb = NAVY
    if subtitle:
        box2 = slide.shapes.add_textbox(Inches(0.5), Inches(0.95), Inches(12.3), Inches(0.4))
        tf2 = box2.text_frame
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.size = Pt(13.5)
        r2.font.italic = True
        r2.font.color.rgb = GREY
    line = slide.shapes.add_shape(1, Inches(0.5), Inches(1.3), Inches(12.3), Pt(2))
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()


def add_bullets(slide, items, left=0.5, top=1.55, width=12.3, height=5.85, size=16, bullet_char="•  "):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(item, tuple):
            text, level = item
        else:
            text, level = item, 0
        indent = "      " * level
        r = p.add_run()
        r.text = f"{indent}{bullet_char if level == 0 else '-  '}{text}"
        r.font.size = Pt(size - level * 1.5)
        r.font.color.rgb = NAVY if level == 0 else GREY
        p.space_after = Pt(8)
    return box


def add_table(slide, df, left, top, width, height, font_size=11, header_size=12):
    rows, cols = df.shape[0] + 1, df.shape[1]
    gtable = slide.shapes.add_table(rows, cols, Inches(left), Inches(top), Inches(width), Inches(height)).table
    for j, col in enumerate(df.columns):
        cell = gtable.cell(0, j)
        cell.text = str(col)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(header_size)
            p.font.bold = True
            p.font.color.rgb = WHITE
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            cell = gtable.cell(i + 1, j)
            cell.text = str(df.iat[i, j])
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(font_size)
                p.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    return gtable


# ---------------------------------------------------------------------------
# Slide 1: Title
# ---------------------------------------------------------------------------
s = add_slide()
bg = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = NAVY
bg.line.fill.background()
box = s.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.3), Inches(2.3))
tf = box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
r = p.add_run()
r.text = "Venture Creed: Customer Clustering"
r.font.size = Pt(40)
r.font.bold = True
r.font.color.rgb = WHITE
p2 = tf.add_paragraph()
r2 = p2.add_run()
r2.text = "Data Science Case Study — Answers to Q1–Q6"
r2.font.size = Pt(20)
r2.font.color.rgb = RGBColor(0xCC, 0xD3, 0xE0)
p3 = tf.add_paragraph()
p3.space_before = Pt(24)
r3 = p3.add_run()
r3.text = (f"Final model silhouette 0.393 (+93% vs. a naive baseline) · "
           f"near-perfect stability (ARI≈1.0) · validated against 3 alternative algorithms")
r3.font.size = Pt(13)
r3.font.italic = True
r3.font.color.rgb = RGBColor(0x9A, 0xA6, 0xB8)

# ---------------------------------------------------------------------------
# Slide 2: Approach overview
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Approach Overview", "A tiered, validated pipeline — not a single clustering pass")
add_bullets(s, [
    "Dataset: 3,030 customers × 18 columns — 6 address/ID fields + 12 categorical business-segment features (no numeric columns).",
    "Step 1 — Data quality scan (Q5): unify inconsistent null markers; discover the core segments go missing in two distinct blocks, not randomly.",
    "Step 2 — Tiering (Q2): split customers by data completeness before modeling. Only customers with zero core missingness are behaviorally clustered; the rest become 3 explicit, actionable “data-gap” segments instead of noise.",
    "Step 3 — Feature selection (Q2): tested reducing to 2-6 features (silhouette jumps to 0.97) and explicitly rejected it — that trivializes to a 2-variable crosstab and discards Profit/Churn/Competitiveness/Volume/Density. All 12 features are kept.",
    "Step 4 — Clustering (Q1): MCA (10 components) + K-Means, chosen after benchmarking against K-Modes, Agglomerative clustering, GMM, Spectral clustering, Birch, and DBSCAN on identical footing.",
    "Step 5 — Visualization (Q3): 2-D MCA projection colored by final segment.",
    "Step 6 — Outlier detection (Q4): two categorical-native methods (AVF + centroid-distance).",
    "Step 7 — Validation (Q6): silhouette, stability, feature-sensitivity, and held-out geography cross-tab all executed and reported, not just listed.",
], size=15)

# ---------------------------------------------------------------------------
# Slide 3: Q1 - Algorithm choice, reasoned from data properties
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q1. Choosing an Algorithm From the Data's Properties", "The algorithm follows from what the data is — not picked by trial and error")
add_bullets(s, [
    "Property: all 12 usable features are categorical (mixed ordinal, e.g. Revenue L/M/H, and nominal, e.g. Seasonality). No numeric columns exist anywhere in the dataset.",
    ("-> K-Means directly on category codes is invalid: assigning L=1, M=2, H=3 fabricates distances that aren't real, and nominal fields (e.g. Seasonality) have no order to encode as numbers at all.", 1),
    "Property: missingness is structured in two correlated blocks, not random (Q2) -> any distance metric must treat “unknown” as genuinely distinct from “equal,” not as an ordinary category.",
    ("-> A Gower-style distance is used that excludes a feature from a pair's comparison whenever either customer is “Not_Assigned” there.", 1),
    "Given both properties: K-Modes (Huang) is the textbook fit for pure categorical data and is tried first as the baseline — but its matching dissimilarity is coarse (only 13 possible discrete values across 12 features), which is why it caps out at 0.20-0.28 silhouette. That's a property of the metric, not a bug in the algorithm.",
    "Multiple Correspondence Analysis (MCA) is the established technique (French school of data analysis; Husson, Josse & Pagès' HCPC method) for projecting categorical variables into a continuous space that preserves chi-square association between categories — at which point continuous-geometry methods like K-Means become valid to apply, not just convenient.",
    "This also coheres with Q3, which independently calls for MCA for the 2-D plot — one transform serves both purposes rather than stitching together unrelated techniques.",
], size=14)

# ---------------------------------------------------------------------------
# Slide 3b: Q1 - Final model mechanics + confirmatory bake-off
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q1. Final Model & Confirmatory Bake-Off", "MCA + K-Means, with alternatives checked as evidence — not as the basis for the choice")
add_bullets(s, [
    "Customers are first split by data completeness (Q2); only the 753 customers with zero core missingness are behaviorally clustered.",
    "Final model: features are projected into a 10-component continuous MCA space (~60% cumulative inertia, where adding more components stops improving results), then K-Means clusters within that space.",
    "k=4 chosen via a clear silhouette peak (0.393), well above k=3 (0.370) and k=5 (0.372).",
    "Confirmatory check: benchmarked against 5 alternatives on identical footing (same Gower-based silhouette) — K-Modes (naive), Agglomerative clustering, Gaussian Mixture Models, Spectral clustering, Birch, and DBSCAN. K-Means won on negative-silhouette share and worst-cluster silhouette; DBSCAN and Birch were rejected outright (degenerate splits, unusable for a segmentation everyone must fit into).",
    "Result: 4 behavioral clusters + 3 partial-profile segments + 1 not-yet-profiled segment = 8 customer segments covering all 3,030 accounts.",
], size=15.5)

# ---------------------------------------------------------------------------
# Slide 4: Model validation & improvement (NEW)
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Model Validation: Measured, Not Asserted", "Every number below was computed against the actual model, including two earlier rejected versions")
improve_tbl = pd.DataFrame([
    ["Approach", "Naive K-Modes\n(everyone together)", "Tiered + Agglomerative\n(Gower distance)", "Tiered + MCA + K-Means\n(FINAL)"],
    ["Silhouette (Gower-based)", "0.203", "0.275", "0.393"],
    ["% customers, negative silhouette", "18.9%", "7.3%", "1.2%"],
    ["Stability across random seeds", "ARI ≈ 0.46 (moderate)", "Deterministic", "ARI ≈ 1.00 (near-perfect)"],
    ["Segment vs. held-out Country (Cramér's V)", "0.285", "0.310", "0.318"],
])
improve_tbl.columns = improve_tbl.iloc[0]
improve_tbl = improve_tbl[1:]
add_table(s, improve_tbl, left=0.5, top=1.55, width=12.3, height=2.6, font_size=12, header_size=12)
add_bullets(s, [
    "The naive version's one weak point: a cluster ~99% driven by shared missingness, not behavior — fixed by tiering + a missingness-aware distance metric (Q2).",
    "The tiered version's weak point: matching distance is coarse (13 discrete values) and chains under Agglomerative linkage — fixed by projecting into a continuous MCA space before clustering.",
    "A feature-reduction shortcut was also tested and rejected: dropping to 2-6 features pushes silhouette to 0.60-0.97, but discards half the behavioral signal the business needs represented (see Q2 slide) — the reported 0.393 is the honest, complete-data number.",
], left=0.6, top=4.35, width=12.1, height=2.9, size=13.5)

# ---------------------------------------------------------------------------
# Slide 5: Q1 - Cluster table
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q1. Cluster → Account Mapping & Statistics", "Full mapping in outputs/Cluster_Customer_Mapping.xlsx")
disp = prof[["cluster", "cluster_name", "size", "pct_of_group"]].copy()
disp["pct_of_group"] = (disp["pct_of_group"] * 100).round(1).astype(str) + "%"
disp.columns = ["Cluster ID", "Cluster Name", "Size (fully-profiled base)", "% of Fully-Profiled"]
add_table(s, disp, left=0.6, top=1.6, width=9.5, height=2.2, font_size=13, header_size=13)
notes = [
    "Names are derived programmatically from each cluster's dominant Casino_Size_Segment (format) x Market_Share/Potential (headroom) — not hardcoded by index, since K-Means label order isn't guaranteed stable.",
    "Mid-Format High-Potential Accounts: medium casino size (97%), high market potential (99%), mid revenue.",
    "Small-Format Steady Base: small casino size (100%), moderate potential, low revenue but profitable — the largest group.",
    "Mature Small-Format, Low-Headroom Accounts: small format, but Venture Creed already holds high market share (99%) — limited room to grow further.",
    "Premium Large-Format Growth Leaders: large casino size (100%), very high market potential (87%), highest propensity to buy (65%) — top marketing-investment target.",
    "Plus: Revenue-Tracked/Territory-Pending (n=205), Territory-Scored/Revenue-Pending (n=139), Partially Profiled Mixed (n=144), and Not-Yet-Profiled (n=1,789) — held outside the behavioral model (see Q2).",
]
add_bullets(s, notes, left=0.6, top=4.0, width=12.1, height=3.35, size=12.5)

# ---------------------------------------------------------------------------
# Slide 6: Q2 Feature selection
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q2. Feature Selection Approach", "What went into the model, what didn't, and what was tested and rejected")
add_bullets(s, [
    "Dropped — Customer_ID, Street, Postal_Code: near-unique identifiers (Postal_Code has 1,124 distinct values); no behavioral signal, would dominate a matching distance without generalizing.",
    "Dropped — City, State_Code, Country: geographic behavior is already captured by purpose-built features (Competitiveness_Flag, Volume_Segment, Density_Segment); raw geography would double-count that signal and add 1,245/61/33-level nominal noise.",
    "Kept — all 12 business segment columns: tested reducing to fewer features (a smaller model scores much higher — 0.97 silhouette at just 2 features) and explicitly REJECTED it: at that point the model is just re-deriving a crosstab of 2 raw fields, discarding Profit/Churn/Competitiveness/Volume/Density entirely. All 12 are kept as the more complete, honest choice.",
    "Missingness handled by tiering, not row-dropping: 59% of customers have all 6 core segments unset simultaneously (dropping them would remove most of the customer base) — they're split into explicit, actionable data-completeness segments instead of being blended into the model as noise.",
    "Missing markers (“None”, “-”, NaN) unified into one explicit category, then further refined via a Gower-style distance that excludes jointly-missing features from each pairwise comparison — so shared “don't know” never counts as similarity.",
], size=14.5)

# ---------------------------------------------------------------------------
# Slide 7: Q3 Visualization
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q3. 2-D Visualization of Clusters", "MCA projection, colored by final segment")
s.shapes.add_picture("outputs/cluster_2d_plot.png", Inches(2.9), Inches(1.5), height=Inches(5.35))
box = s.shapes.add_textbox(Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.5))
p = box.text_frame.paragraphs[0]
p.word_wrap = True
r = p.add_run()
r.text = ("This 2-D view fits MCA with only 2 components (21.7% inertia) purely for plotting, while the actual "
          "clustering model uses 10 components (60.3% inertia) — some overlap between behavioral clusters here is "
          "a plotting limitation, not evidence of poor separation (real silhouette in the full space: 0.393).")
r.font.size = Pt(11)
r.font.italic = True
r.font.color.rgb = GREY

# ---------------------------------------------------------------------------
# Slide 8: Q4 EDA & Outlier detection
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q4. EDA & Outlier Detection", "Two categorical-appropriate methods, applied together")
add_bullets(s, [
    "EDA covered: per-column value distributions/cardinality, 3 different null markers, the two-block missingness structure across the 6 core segments, and duplicate/casing/format checks on address fields (full detail in the notebook and Q5).",
    "No numeric features exist, so z-score / IQR-on-raw-values methods don't apply — two categorical-native methods were used instead:",
    ("Attribute Value Frequency (AVF): average population frequency of a customer's own category values across all 12 features; the rarest 5% are flagged (63 customers). Standard technique for categorical outlier detection (Koufakou & Georgiopoulos, 2010) — no distance metric needed.", 1),
    ("Distance-to-cluster-medoid: within the fully-profiled tier, Gower-style dissimilarity from each customer to the most central point of their own cluster; a 95th-percentile cutoff is used since this distance is bounded (45 customers flagged).", 1),
    "Union of both methods: 103 customers flagged — exported to the “Outliers” sheet of the Excel workbook for manual review, not automatic exclusion (they may be strategically important large/unusual accounts, not errors).",
], size=15)

# ---------------------------------------------------------------------------
# Slide 9: Q5 Data quality issues
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q5. Data Quality Issues Identified", "Full detail + affected-row counts in outputs/data_quality_issues.csv")
dq_disp = dq[["Issue", "Affected_Rows"]].drop_duplicates(subset=["Issue"]).copy()
dq_disp.columns = ["Issue", "Rows Affected"]
add_table(s, dq_disp, left=0.5, top=1.55, width=12.3, height=5.5, font_size=11.5, header_size=12)

# ---------------------------------------------------------------------------
# Slide 10: Q6 Quality checks (executed results)
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q6. Quality Checks — Executed, Not Just Listed", "Every check below was run against the real model")
add_bullets(s, [
    f"Fit quality: silhouette 0.393 (naive baseline 0.203, +93% relative); negative-silhouette share only 1.2% (vs 18.9% naive).",
    f"Stability: mean ARI across 5 random K-Means seeds = 1.000 (naive K-Modes under random init: ≈0.46) — the model converges to the same partition essentially every time.",
    f"Feature-subset sensitivity: dropping Market_Potential_Segment or Casino_Size_Segment changes the partition the most (ARI 0.20 / 0.36) — flagged as a concentration risk, not hidden; 8 of the other 10 features could be dropped individually with ARI staying above 0.97.",
    "No cluster is a missingness artifact: by construction, only customers with zero core missingness are clustered — verified via the cluster profile table (every cluster's modal value on every core feature is a real category).",
    f"Business validation: Segment vs. held-out Country (never used in clustering) is statistically significant with a moderate effect size (Cramér's V = 0.318) — segments track real-world geography.",
    "Pipeline integrity: all 3,030 customers appear exactly once in the output; no nulls in the Segment column; verified programmatically.",
    "Outlier handling: 103 customers flagged for manual review, not automatic exclusion.",
    "Ongoing monitoring: re-run on a fixed cadence (e.g. quarterly); track migration from Tier 0/1 into Tier 2 as a KPI as more purchase history accrues.",
], size=13.5)

# ---------------------------------------------------------------------------
# Slide 11: Assumptions
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Assumptions Taken", "Highlighted per case-study instructions")
add_bullets(s, [
    "Segments already marked Not_Assigned/None/– for a customer reflect “not computed / not applicable yet,” not a true zero value, and were preserved as an explicit category rather than imputed.",
    "Customers missing all 6 core behavioral segments are assumed to have no usable purchasing history yet (new or dormant accounts) and are reported as a separate business segment rather than clustered on missing data.",
    "Duplicate-looking records (identical fields aside from Customer_ID, or shared addresses) are treated as distinct accounts for this exercise — flagged for business confirmation rather than merged.",
    "Geographic behavior is assumed to be adequately captured by Competitiveness_Flag / Volume_Segment / Density_Segment, making raw City/State/Country redundant for a *behavioral* clustering.",
    "All 12 features are kept deliberately, even though a smaller feature set scores higher on paper (see Q2) — the business asked for a purchasing-behavior segmentation across everything provided, not the smallest set that maximizes one internal metric.",
    "k=4 (plus the 3 partial-profile segments and the not-yet-profiled group) is the primary recommendation; the notebook exposes k=2–8 so the business can revisit granularity as strategy needs evolve.",
], size=15)

prs.save("Venture_Creed_Case_Study_Answers.pptx")
print("Saved Venture_Creed_Case_Study_Answers.pptx")
