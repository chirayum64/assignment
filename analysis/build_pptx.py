"""Builds the Venture_Creed_Case_Study_Answers.pptx deliverable answering Q1-Q6."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
import pandas as pd

NAVY = RGBColor(0x1F, 0x2A, 0x44)
ACCENT = RGBColor(0xC0, 0x39, 0x2B)
GREY = RGBColor(0x55, 0x55, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

prof = pd.read_csv("outputs/cluster_profiles.csv")
dq = pd.read_csv("outputs/data_quality_issues.csv")

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
    r.font.size = Pt(28)
    r.font.bold = True
    r.font.color.rgb = NAVY
    if subtitle:
        box2 = slide.shapes.add_textbox(Inches(0.5), Inches(0.95), Inches(12.3), Inches(0.4))
        tf2 = box2.text_frame
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.size = Pt(14)
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
        p.level = 0
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
box = s.shapes.add_textbox(Inches(1), Inches(2.6), Inches(11.3), Inches(2))
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
p3.space_before = Pt(30)
r3 = p3.add_run()
r3.text = "Approach, model, visualization, data quality, and validation plan"
r3.font.size = Pt(14)
r3.font.italic = True
r3.font.color.rgb = RGBColor(0x9A, 0xA6, 0xB8)

# ---------------------------------------------------------------------------
# Slide 2: Approach overview
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Approach Overview", "How the six questions map to one coherent pipeline")
add_bullets(s, [
    "Dataset: 3,030 customers × 18 columns — 6 address/ID fields + 12 categorical business-segment features (no numeric columns).",
    "Step 1 — Data quality scan (Q5): unify inconsistent null markers, flag typos/duplicates/format issues.",
    "Step 2 — Feature selection (Q2): drop identifiers/high-cardinality address fields; separate customers with zero behavioral signal from the rest.",
    "Step 3 — Clustering (Q1): K-Modes on the 12 behavioral segments for the profiled customers; k chosen via cost + silhouette.",
    "Step 4 — Visualization (Q3): 2-D Multiple Correspondence Analysis (MCA), colored by cluster.",
    "Step 5 — Outlier detection (Q4): Attribute Value Frequency + distance-to-cluster-centroid, both suited to categorical data.",
    "Step 6 — Validation plan (Q6): stability, sensitivity, and business-facing checks before the segmentation is used operationally.",
    "All code lives in analysis/clustering_analysis.ipynb (executed, outputs embedded); all data outputs in analysis/outputs/.",
], size=17)

# ---------------------------------------------------------------------------
# Slide 3: Q1 - Algorithm
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q1. Clustering Logic & Algorithm", "K-Modes on categorical purchasing-behavior segments")
add_bullets(s, [
    "All retained features are categorical (ordinal or nominal) — K-Means / Euclidean distance do not apply.",
    "K-Modes (Huang, 1998) is used instead: cluster centers are the modal category per feature; dissimilarity between two customers = fraction of the 12 features where their values differ (simple matching dissimilarity).",
    "Customers are first split into “behaviorally scored” (n=1,241) vs. “not-yet-profiled” (n=1,789, all 6 core segments unset) — clustering only the scored group avoids one giant cluster driven purely by missing data (see Q2).",
    "Choosing k: tracked K-Modes cost (elbow) and silhouette score (on the full pairwise Hamming-distance matrix) for k = 2–8.",
    "Cost drops sharply and silhouette roughly doubles (0.14 → 0.20) going from k=5 to k=6; k=7 gives only a marginal further gain at the cost of one small, less interpretable cluster → k = 6 selected.",
    "Result: 6 behavioral clusters + 1 “not-yet-profiled” group = 7 customer segments covering all 3,030 accounts.",
], size=16)

# ---------------------------------------------------------------------------
# Slide 4: Q1 - Cluster table
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q1. Cluster → Account Mapping & Statistics", "Full mapping in outputs/Cluster_Customer_Mapping.xlsx")
disp = prof[["cluster", "cluster_name", "size", "pct_of_group"]].copy()
disp["pct_of_group"] = (disp["pct_of_group"] * 100).round(1).astype(str) + "%"
disp.columns = ["Cluster ID", "Cluster Name", "Size (scored base)", "% of Scored"]
add_table(s, disp, left=0.6, top=1.6, width=8.0, height=2.6, font_size=13, header_size=13)
notes = [
    "0 — Steady Small-Format Core: small casino size, moderate share/potential, minimal churn, high-volume/density geography.",
    "1 — Under-Profiled Moderate Accounts: share/size/potential mostly unscored; medium volume.",
    "2 — Emerging High-Potential Adopters: early adopters, high profit & potential, low volume — emerging markets.",
    "3 — Premium Growth Leaders: high revenue/profit/potential/propensity, competitive & dense geography — top marketing-investment target.",
    "4 — Mature Low-Headroom Small Accounts: small, already-high share, low remaining potential.",
    "5 — High-Density Competitive, Thin-Profile: low revenue in dense, competitive markets; share/size/potential mostly unscored.",
    "Plus: Not-Yet-Profiled / New-or-Dormant Accounts (n=1,789) — held out of the model, no behavioral data available yet.",
]
add_bullets(s, notes, left=0.6, top=4.35, width=12.1, height=3.0, size=13)

# ---------------------------------------------------------------------------
# Slide 5: Q2 Feature selection
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q2. Feature Selection Approach", "What went into the model, and why")
add_bullets(s, [
    "Dropped — Customer_ID, Street, Postal_Code: identifiers / near-unique free text (Postal_Code alone has 1,124 distinct values); no direct behavioral signal, would dominate a matching-based distance metric without generalizing.",
    "Dropped — City, State_Code, Country: the brief asks for behavior-based clusters, and geographic behavior is already captured by purpose-built features (Competitiveness_Flag, Volume_Segment, Density_Segment). Keeping raw geography too would double-count that signal and add 1,245 / 61 / 33-level nominal noise.",
    "Kept — all 12 business segment columns: the only fields that directly describe purchasing/behavioral characteristics.",
    "Missingness treated as a modeling decision, not a row-drop: 59% of customers have all 6 core behavioral segments unset simultaneously — dropping incomplete rows would discard most of the customer base.",
    "All missing markers (“None”, “-”, NaN) unified into one explicit category, Not_Assigned, per column — missing-ness can itself be informative, and K-Modes treats it as just another category.",
    "Customers with zero behavioral signal are separated out before clustering rather than forced into the model — confirmed empirically: clustering the full base produces a dominant cluster that is really just “all fields Not_Assigned”, not a genuine behavioral group (low silhouette regardless of k).",
], size=15.5)

# ---------------------------------------------------------------------------
# Slide 6: Q3 Visualization
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q3. 2-D Visualization of Clusters", "Multiple Correspondence Analysis (MCA) — the categorical analogue of PCA")
s.shapes.add_picture("outputs/cluster_2d_plot.png", Inches(2.6), Inches(1.55), height=Inches(5.6))
box = s.shapes.add_textbox(Inches(0.5), Inches(7.15), Inches(12.3), Inches(0.3))
p = box.text_frame.paragraphs[0]
r = p.add_run()
r.text = "MCA reduces the 12 categorical features to 2 dimensions (21.7% of inertia explained) for plotting; PCA is not applicable since there are no numeric features."
r.font.size = Pt(11)
r.font.italic = True
r.font.color.rgb = GREY

# ---------------------------------------------------------------------------
# Slide 7: Q4 EDA & Outlier detection
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q4. EDA & Outlier Detection", "Two categorical-appropriate methods, applied together")
add_bullets(s, [
    "EDA covered: per-column value distributions and cardinality, missingness by column (3 different null markers found), block-missingness across the 6 core segments, and duplicate/casing/format checks on address fields (full detail in the notebook and Q5).",
    "No numeric features exist, so z-score / IQR-on-raw-values methods don't apply — two categorical-native methods were used instead:",
    ("Attribute Value Frequency (AVF): for each customer, average the population frequency of the categories they hold across all 12 features; customers whose values are individually rare are flagged (bottom 5th percentile → 63 customers). Standard technique for categorical outlier detection (Koufakou & Georgiopoulos, 2010) — needs no distance metric.", 1),
    ("Distance-to-own-cluster-centroid: Hamming distance from each customer to their assigned cluster's modal profile. This distance is bounded and only takes 13 discrete values, so a Tukey IQR fence never triggers — a 95th-percentile cutoff is used instead (86 customers flagged).", 1),
    "Union of both methods: 104 customers flagged — exported to outputs/outliers.csv / the “Outliers” sheet of the Excel workbook for manual review (code in the notebook, Section 6).",
], size=15)

# ---------------------------------------------------------------------------
# Slide 8: Q5 Data quality issues
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q5. Data Quality Issues Identified", "Full detail + affected-row counts in outputs/data_quality_issues.csv")
dq_disp = dq[["Issue", "Affected_Rows"]].drop_duplicates(subset=["Issue"]).copy()
dq_disp.columns = ["Issue", "Rows Affected"]
add_table(s, dq_disp, left=0.5, top=1.55, width=12.3, height=5.5, font_size=11.5, header_size=12)

# ---------------------------------------------------------------------------
# Slide 9: Q6 Quality checks
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Q6. Quality Checks to Validate the Solution", "Exhaustive checklist across pipeline, model, and business fit")
add_bullets(s, [
    "Data/pipeline integrity: every customer appears exactly once in the output; no nulls in Cluster_ID/Cluster_Name; cluster IDs within expected range.",
    "No degenerate clusters: enforce a minimum cluster-size floor (e.g. ≥3% of the scored base).",
    "Fit quality: silhouette score above a minimum bar for the chosen k; compare against k±1 to confirm k wasn't cherry-picked off a local blip.",
    "Stability: re-run K-Modes across multiple seeds/n_init and measure agreement via Adjusted Rand Index (target ≥0.8) before trusting the assignment.",
    "Feature-subset sensitivity: re-run leaving one feature out at a time; structure should degrade gracefully, not collapse — confirms no single feature drives all separation.",
    "Guard against the missingness artifact: verify no cluster's dominant modal value across most features is Not_Assigned (the failure mode found and corrected during modeling).",
    "Business validation: cross-tab clusters against held-out fields (e.g. City/Country) for real-world sense-check; review cluster profiles and names with Venture Creed sales/marketing for actionability.",
    "Outlier handling: confirm flagged outliers don't disproportionately anchor a centroid (re-fit excluding them and compare); route to manual review rather than silent drop.",
    "Ongoing monitoring: re-run on a fixed cadence (e.g. quarterly) as more customers become “scored”; track migration between clusters/out of “not-yet-profiled” to monitor drift.",
], size=14.5)

# ---------------------------------------------------------------------------
# Slide 10: Assumptions
# ---------------------------------------------------------------------------
s = add_slide()
add_title(s, "Assumptions Taken", "Highlighted per case-study instructions")
add_bullets(s, [
    "Segments already marked Not_Assigned/None/– for a customer reflect “not computed / not applicable yet”, not a true zero value, and were preserved as an explicit category rather than imputed.",
    "Customers missing all 6 core behavioral segments are assumed to have no usable purchasing history yet (new or dormant accounts) and are reported as a separate business segment rather than clustered on missing data.",
    "Duplicate-looking records (identical fields aside from Customer_ID, or shared addresses) are treated as distinct accounts for this exercise — flagged for business confirmation rather than merged, since merging without domain input risks losing genuinely separate accounts (e.g. multiple properties).",
    "Geographic behavior is assumed to be adequately captured by Competitiveness_Flag / Volume_Segment / Density_Segment, making raw City/State/Country redundant for a *behavioral* clustering (as opposed to a purely geographic one).",
    "k=6 (plus the separate not-yet-profiled group) is treated as the primary recommendation; the notebook exposes k=2–8 so the business can revisit granularity as strategy needs evolve.",
], size=16)

prs.save("Venture_Creed_Case_Study_Answers.pptx")
print("Saved Venture_Creed_Case_Study_Answers.pptx")
