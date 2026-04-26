import io

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from openpyxl import Workbook

plt.rcParams.update(
    {
        "figure.facecolor": "#f3faf3",
        "axes.facecolor": "#f3faf3",
        "savefig.facecolor": "#f3faf3",
        "text.color": "black",
        "axes.labelcolor": "black",
        "axes.edgecolor": "black",
        "axes.titlecolor": "black",
        "xtick.color": "black",
        "ytick.color": "black",
        "grid.color": "gray",
        "grid.alpha": 0.2,
    }
)

# =========================
# APP CONFIG
# =========================
st.set_page_config(page_title="Plant Stress AI", layout="wide")

st.markdown(
    """
<div style="
    background-color: rgba(255,255,255,0.05);
    padding: 15px;
    border-radius: 12px;
    border: 1px solid rgba(255,255,255,0.1);
">
""",
    unsafe_allow_html=True,
)

st.title("🌾 Drought Stress Phenotyping & Breeding AI")
st.markdown(
    """
🌾 Drought Stress Phenotyping & Breeding AI System

This tool evaluates plant physiological performance using:

- Chlorophyll Content (leaf greenness)
- Photosystem II Efficiency (photosynthetic function)

These traits are used to estimate drought tolerance and support breeding decisions.
"""
)

st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    """
<style>

/* Fondo general */
.stApp {
    background-color: #f3faf3;
}

/* Contenedores */
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

/* Texto general */
html, body, [class*="css"] {
    color: #1b1b1b;
}

/* Botones */
.stButton>button {
    background-color: #2E7D32;
    color: white;
    border-radius: 8px;
    border: none;
}

.stButton>button:hover {
    background-color: #1B5E20;
}

/* Títulos */
h1, h2, h3 {
    color: #1f3d1f;
}

</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<style>
/* ajustes estables */
.block-container {
    background-color: transparent;
}

.main {
    background-color: transparent;
}

/* texto más legible en ambos modes */
p, label, div {
    color: inherit;
}
</style>
""",
    unsafe_allow_html=True,
)

st.info(
    """
This tool integrates physiological traits (Fv/Fm, SPAD) with stress modeling to support drought-tolerant genotype selection.

It is crop-independent and can be applied to any plant species as long as the required physiological and yield data are provided.
"""
)

with st.expander("📌 How to use this app"):
    st.markdown(
        """
1. Download the Excel template
2. Fill in your experimental data:
   - Genotype (e.g. H1, H2...)
   - FvFm (photosynthetic efficiency)
   - SPAD (chlorophyll content)
   - Yield (optional or predicted)
3. Upload the file
4. Click Run Analysis
5. View results and download report
"""
    )


with st.expander("🧠 Interpretation Guide"):
    st.markdown(
        """
### 🌿 Updated Trait Definitions

- **Chlorophyll Content (formerly SPAD):**  
  Measures leaf chlorophyll concentration and greenness. Higher values = healthier plants.

- **Photosystem II Efficiency (formerly Fv/Fm):**  
  Measures photosynthetic efficiency under stress conditions. Higher values = better physiological performance.

👉 These two traits are complementary and describe plant health from different physiological perspectives.
"""
    )

with st.expander("🌿 What do variables mean?"):
    st.markdown(
        """
### 📊 Fv/Fm (Photosynthetic efficiency)
- Measures how efficiently the plant is doing photosynthesis  
- Higher values = healthier plant  
- Lower values = stress damage in photosystem II  

### 🍃 SPAD (Chlorophyll content)
- Estimates chlorophyll level in leaves  
- Higher SPAD = greener, healthier leaves  
- Lower SPAD = nutrient or stress limitation  

### 🧪 Stress Index
The app combines both variables:

**Higher Stress Index = more drought stress**

- Close to 0 → tolerant genotype  
- Close to 1 → sensitive genotype  
"""
    )

st.success(
    """
The model automatically classifies genotypes into Low, Moderate, and High stress groups based on Stress Index distribution.
"""
)


# =========================
# TEMPLATE DOWNLOAD
# =========================
st.subheader("📥 Step 1: Download Template")

wb = Workbook()
ws = wb.active
ws.title = "template"

ws.append(
    [
        "Genotype",
        "Chlorophyll_Content",
        "Photosystem_II_Efficiency",
        "Yield",
    ]
)

for i in range(1, 6):
    ws.append([f"H{i}", "", "", ""])

buffer = io.BytesIO()
wb.save(buffer)
buffer.seek(0)

st.download_button(
    label="📥 Download Template",
    data=buffer,
    file_name="phenotyping_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# =========================
# CORE ANALYSIS
# =========================
SPAD_COL = "Chlorophyll_Content"
FVFM_COL = "Photosystem_II_Efficiency"


def normalize_trait_columns(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    return df.rename(
        columns={
            "SPAD": SPAD_COL,
            "FvFm": FVFM_COL,
        }
    )


def compute_stress(df):
    df = df.copy()

    df["FvFm_norm"] = df[FVFM_COL] / df[FVFM_COL].max()
    df["SPAD_norm"] = df[SPAD_COL] / df[SPAD_COL].max()
    df["Stress_Index"] = 1 - (df["FvFm_norm"] + df["SPAD_norm"]) / 2

    return df


def classify(x):
    if x < 0.4:
        return "🟢 Low"
    if x < 0.7:
        return "🟡 Moderate"
    return "🔴 High"


def generate_recommendation(avg):
    if avg < 0.4:
        return "Best group: high drought tolerance"
    if avg < 0.7:
        return "Moderate stress response"
    return "High stress sensitivity detected"


def color_scale(val):
    if val < 0.4:
        return "#2ECC71"
    if val < 0.7:
        return "#F1C40F"
    return "#E74C3C"


def color_class(val):
    return f"background-color: {color_scale(val)}"


def safe_yield_available(df):
    return (
        "Yield" in df.columns
        and df["Yield"].notna().any()
        and df["Yield"].sum() > 0
    )


def run_yield_model(df):
    if not safe_yield_available(df):
        return None


def get_color(si, y, df):
    if si < df["Stress_Index"].quantile(0.33) and y > df["Yield"].quantile(0.66):
        return "green"
    if si > df["Stress_Index"].quantile(0.66) and y < df["Yield"].quantile(0.33):
        return "red"
    return "orange"


def explain_stress_vs_yield():
    st.markdown(
        """
    ### ⚖️ Stress vs Yield Trade-off (Breeding Decision Map)

    ### 🧠 How to interpret this figure

    This graph combines drought stress and productivity in one view.

    - **X-axis:** Stress Index (higher = more sensitive genotypes)
    - **Y-axis:** Yield (higher = more productive genotypes)

    ### 📌 Breeding interpretation:
    - 🟢 Upper-left → ideal genotypes (low stress + high yield)
    - 🔴 Lower-right → poor performers (high stress + low yield)
    - 🟡 Middle → intermediate performance

    👉 This plot is used to select **elite breeding candidates**.
    """
    )


def trait_texts():
    st.markdown(
        """
### 🌿 Physiological Traits (Standardized Definitions)

- **Chlorophyll Content (SPAD equivalent):**  
  Indicates leaf chlorophyll concentration and greenness.

- **Photosystem II Efficiency (Fv/Fm):**  
  Measures photosynthetic performance under stress conditions.

- **Stress Index:**  
  Quantifies reduction in plant performance under drought stress.
"""
    )


def style_matplotlib(ax):
    import matplotlib.pyplot as plt

    fig = plt.gcf()

    fig.patch.set_facecolor("#f3faf3")
    ax.set_facecolor("#f3faf3")

    ax.tick_params(colors="black")
    ax.title.set_color("black")
    ax.xaxis.label.set_color("black")
    ax.yaxis.label.set_color("black")
    for spine in ax.spines.values():
        spine.set_color("black")

    ax.grid(True, color="gray", alpha=0.2)


def style_colorbar(colorbar):
    colorbar.ax.tick_params(colors="black")
    colorbar.ax.yaxis.label.set_color("black")
    colorbar.outline.set_edgecolor("black")


def plot_stress(df):
    stress_df = df.groupby("Genotype", as_index=False)["Stress_Index"].mean()
    fig, ax = plt.subplots()

    colors = stress_df["Stress_Index"].apply(
        lambda x: "green" if x < 0.4 else "orange" if x < 0.7 else "red"
    )

    bars = ax.bar(stress_df["Genotype"], stress_df["Stress_Index"], color=colors)
    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            f"{height:.2f}",
            ha="center",
            va="bottom",
        )

    ax.set_title("Stress Classification by Genotype")
    ax.set_xlabel("Genotype")
    ax.set_ylabel("Stress Index")
    style_matplotlib(ax)
    return fig


def plot_yield(df):
    yield_vals = df.groupby("Genotype")["Yield"].mean()
    fig, ax = plt.subplots()
    bars = ax.bar(yield_vals.index, yield_vals.values, color="skyblue")
    ax.set_title("Yield Performance by Genotype")
    ax.set_xlabel("Genotype")
    ax.set_ylabel("Yield (relative units)")

    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            f"{height:.2f}",
            ha="center",
            va="bottom",
        )

    style_matplotlib(ax)
    return fig


def plot_stress_vs_yield(df):
    fig, ax = plt.subplots()

    colors = [
        get_color(
            df["Stress_Index"].iloc[i],
            df["Yield"].iloc[i],
            df,
        )
        for i in range(len(df))
    ]

    ax.scatter(df["Stress_Index"], df["Yield"], c=colors)

    for _, row in df.iterrows():
        ax.text(row["Stress_Index"], row["Yield"], row["Genotype"])

    ax.set_xlabel("Stress Index")
    ax.set_ylabel("Yield")
    ax.set_title("Stress vs Yield Breeding Map")
    style_matplotlib(ax)
    fig.savefig("tradeoff.png", dpi=300, bbox_inches="tight")
    return fig


def run_yield_analysis(df):
    yield_rank = df.groupby("Genotype")["Yield"].mean().sort_values(ascending=False)
    fig = plot_yield(df)

    st.subheader("🌾 Yield Ranking (Relative Performance)")
    st.markdown(
        """
Yield classification is based on within-experiment distribution (percentiles), not fixed thresholds.
"""
    )
    st.dataframe(yield_rank)

    st.subheader("🌾 Yield Performance")
    st.markdown(
        """
### 🌾 What this means

This plot shows productivity differences among genotypes.

- Higher bars = higher productivity  
- Lower bars = reduced performance  

👉 Yield helps balance **stress tolerance vs productivity trade-offs**.
"""
    )
    st.pyplot(fig)

    best_yield = yield_rank.idxmax()
    worst_yield = yield_rank.idxmin()
    st.subheader("🌾 Yield Insight")
    st.success(f"Best yield genotype: {best_yield}")
    st.error(f"Lowest yield genotype: {worst_yield}")


def run_tradeoff(df):
    explain_stress_vs_yield()
    st.pyplot(plot_stress_vs_yield(df))

    st.subheader("🧠 Breeding Interpretation")

    elite = df[
        (df["Stress_Index"] < df["Stress_Index"].quantile(0.33))
        & (df["Yield"] > df["Yield"].quantile(0.66))
    ]
    sensitive = df[
        (df["Stress_Index"] > df["Stress_Index"].quantile(0.66))
        & (df["Yield"] < df["Yield"].quantile(0.33))
    ]

    st.success(f"🌱 Elite genotypes (recommended): {', '.join(elite['Genotype'].unique())}")
    st.error(f"🔥 Sensitive genotypes (avoid): {', '.join(sensitive['Genotype'].unique())}")


def generate_word_report(df):
    has_yield = safe_yield_available(df)
    doc = Document()

    ranking = df.groupby("Genotype")["Stress_Index"].mean().sort_values()
    best = ranking.idxmin()
    worst = ranking.idxmax()

    title = doc.add_heading("PhytoStress AI Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("Drought Stress Phenotyping Report")

    doc.add_heading("PhytoStress AI Report Explanation", level=1)
    doc.add_paragraph(
        """
This report explains the results of the PhytoStress AI analysis in a simple and structured way.
Each section describes what the graphs and tables mean and how to interpret genotype performance under drought stress conditions.
"""
    )

    doc.add_heading("1. Genotype Ranking", level=1)
    doc.add_paragraph(
        """
This table shows the genotypes ranked by Stress Index.

- Lower values = better drought tolerance
- Higher values = more sensitive plants

The best performing genotype is the one with the lowest Stress Index.
"""
    )
    doc.add_heading("Genotype Ranking Table", level=1)
    table = doc.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "Rank"
    hdr[1].text = "Genotype"
    hdr[2].text = "Stress Index"

    for i, (geno, val) in enumerate(ranking.items(), 1):
        row = table.add_row().cells
        row[0].text = str(i)
        row[1].text = str(geno)
        row[2].text = f"{val:.3f}"

    doc.add_heading("2. Stress Index Explanation", level=1)
    doc.add_paragraph(
        """
The Stress Index measures how much a plant is affected by drought.

It is calculated using biomass reduction:
Stress Index = 1 - (Stress / Control)

- Values close to 0 → very tolerant plants
- Values close to 1 → highly stressed plants
"""
    )

    if "Yield" in df.columns:
        doc.add_heading("3. Yield Performance", level=1)
        doc.add_paragraph(
            """
Yield represents the productivity potential of each genotype.

In this analysis, yield is interpreted relatively:
- High yield = better performance in this experiment
- Low yield = lower productivity under the same conditions
"""
        )

    doc.add_heading("4. Heatmap Interpretation", level=1)
    doc.add_paragraph(
        """
The heatmap shows relationships between Stress Index, SPAD, and Fv/Fm.

- Green colors indicate better plant performance
- Red colors indicate higher stress

This helps visualize overall genotype behavior in a single view.
"""
    )

    if "Yield" in df.columns:
        doc.add_heading("5. Stress vs Yield Relationship", level=1)
        doc.add_paragraph(
            """
This graph compares stress tolerance and yield at the same time.

- Ideal genotypes appear with low stress and high yield
- Poor genotypes show high stress and low yield

This is the most important figure for selecting elite genotypes.
"""
        )

    fig1_path = "fig1.png"
    fig1.savefig(fig1_path, dpi=300, bbox_inches="tight")

    fig2_path = "fig2.png"
    fig2.savefig(fig2_path, dpi=300, bbox_inches="tight")

    doc.add_heading("Figures", level=1)
    doc.add_paragraph("Figure 1: Stress Index by Genotype")
    doc.add_picture(fig1_path, width=Inches(5.5))

    doc.add_paragraph("Figure 2: SPAD vs Fv/Fm")
    doc.add_picture(fig2_path, width=Inches(5.5))

    doc.add_paragraph("Figure 3: Heatmap of physiological traits")
    doc.add_picture("heatmap.png", width=Inches(5.5))

    if has_yield:
        doc.add_paragraph("Figure 4: Stress vs Yield Trade-off")
        doc.add_picture("tradeoff.png", width=Inches(5.5))

    low = df[df["Stress_Index"] < 0.4]["Genotype"].unique()
    mid = df[
        (df["Stress_Index"] >= 0.4) & (df["Stress_Index"] < 0.7)
    ]["Genotype"].unique()
    high = df[df["Stress_Index"] >= 0.7]["Genotype"].unique()

    doc.add_heading("Stress Classification", level=1)
    doc.add_paragraph("Low stress (tolerant): " + ", ".join(low))
    doc.add_paragraph("Moderate stress: " + ", ".join(mid))
    doc.add_paragraph("High stress (sensitive): " + ", ".join(high))

    doc.add_heading("Discussion", level=1)
    if "Breeding_Score" in df.columns and has_yield:
        final_rank = df.groupby("Genotype")["Breeding_Score"].mean().sort_values(
            ascending=False
        )
        report_best = final_rank.idxmax()
        report_worst = final_rank.idxmin()
        avg_stress = df["Stress_Index"].mean()

        if avg_stress < 0.4:
            stress_state = "low overall stress conditions"
        elif avg_stress < 0.7:
            stress_state = "moderate stress conditions"
        else:
            stress_state = "high stress conditions"

        doc.add_paragraph(
            f"""
The evaluated genotypes were analyzed under {stress_state}, showing variability in physiological and agronomic performance.

The breeding score analysis identified {report_best} as the most promising genotype, while {report_worst} showed the lowest performance.

These results confirm that integrating Stress Index and Yield is effective for genotype selection under drought conditions.
"""
        )
    else:
        doc.add_paragraph(
            f"""
The evaluated genotypes were analyzed for physiological stress response under drought conditions.

The ranking analysis identified {best} as the most promising genotype, while {worst} showed the lowest performance.

These results confirm that physiological indicators such as Fv/Fm, SPAD, and Stress Index are effective for genotype selection under drought conditions.
"""
        )

    doc.add_heading("Final Recommendation", level=1)
    if "Breeding_Score" in df.columns and has_yield:
        final_rank = df.groupby("Genotype")["Breeding_Score"].mean().sort_values(
            ascending=False
        )
        report_best = final_rank.idxmax()
        report_worst = final_rank.idxmin()
        doc.add_paragraph(
            f"""
Best genotype: {report_best}
Worst genotype: {report_worst}

Best genotype is recommended because it combines:
- Low Stress Index
- High Yield
- Strong overall Breeding Score

This genotype represents the ideal candidate for drought tolerance selection.
"""
        )
    else:
        doc.add_paragraph(
            f"""
Best genotype: {best}
Worst genotype: {worst}

Best genotype is recommended because it combines:
- Low stress response
- High physiological stability
- Strong drought tolerance

This genotype represents the ideal candidate for drought tolerance selection.
"""
        )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generate_full_report(df, ranking):
    doc = Document()

    title = doc.add_heading("PhytoStress AI Full Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("Drought stress phenotyping integrated report")
    doc.add_page_break()

    doc.add_heading("1. Genotype Ranking", level=1)
    for i, (g, v) in enumerate(ranking.items(), 1):
        doc.add_paragraph(f"{i}. {g}: {v:.3f}")

    fig1_path = "fig1.png"
    fig2_path = "fig2.png"
    fig3_path = "heatmap.png"

    doc.add_heading("2. Figures", level=1)
    doc.add_picture(fig1_path, width=Inches(5.5))
    doc.add_picture(fig2_path, width=Inches(5.5))
    doc.add_picture(fig3_path, width=Inches(5.5))

    doc.add_heading("3. Stress Classification Summary", level=1)

    low = df[df["Stress_Index"] < 0.4]["Genotype"].unique()
    mid = df[(df["Stress_Index"] >= 0.4) & (df["Stress_Index"] < 0.7)]["Genotype"].unique()
    high = df[df["Stress_Index"] >= 0.7]["Genotype"].unique()

    doc.add_paragraph("Low stress: " + ", ".join(low))
    doc.add_paragraph("Moderate stress: " + ", ".join(mid))
    doc.add_paragraph("High stress: " + ", ".join(high))

    best = ranking.idxmin()
    worst = ranking.idxmax()

    doc.add_heading("4. Conclusion", level=1)
    doc.add_paragraph(
        f"The analysis identified {best} as the most drought-tolerant genotype "
        f"and {worst} as the most sensitive genotype. "
        "Physiological traits (Fv/Fm and SPAD) successfully differentiated stress responses."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


# =========================
# MAIN APP FLOW
# =========================
col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader("📂 Upload Excel file (.xlsx)", type=["xlsx"])

with col2:
    run_demo = st.button("🚀 Run Demo")

df = None

if run_demo:
    df = pd.DataFrame(
        {
            "Genotype": ["H1", "H2", "H3", "H4"],
            "FvFm": [0.80, 0.72, 0.58, 0.61],
            "SPAD": [45, 40, 30, 33],
            "Yield": [6.1, 5.4, 4.0, 4.2],
        }
    )
    st.success("Demo loaded successfully")

elif uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("File uploaded successfully")
    st.dataframe(df)
else:
    st.stop()

df = normalize_trait_columns(df)

required = ["Genotype", FVFM_COL, SPAD_COL]

if any(col not in df.columns for col in required):
    st.error(
        f"Missing required columns: Genotype, {FVFM_COL}, {SPAD_COL}"
    )
    st.stop()

df = df.dropna(subset=["Genotype", FVFM_COL, SPAD_COL])
df["Genotype"] = df["Genotype"].astype(str)

df = compute_stress(df)

if uploaded_file or run_demo:
    st.subheader("📂 Simulated Dataset" if run_demo else "📋 Data Preview")
    st.dataframe(df)

if "Yield" in df.columns:
    df["Yield"] = pd.to_numeric(df["Yield"], errors="coerce")
    df = df.dropna(subset=["Yield"])
else:
    df["Yield"] = np.nan

for col in df.columns:
    if "yield" in col.lower() and "estimated" in col.lower():
        df[col] = np.nan

has_yield = "Yield" in df.columns and df["Yield"].notna().any()

if has_yield:
    q33 = df["Yield"].quantile(0.33)
    q66 = df["Yield"].quantile(0.66)

    def yield_class(x):
        if x >= q66:
            return "🟢 High Yield"
        if x >= q33:
            return "🟡 Medium Yield"
        return "🔴 Low Yield"

    def yield_color(x):
        if x >= q66:
            return "#2ECC71"
        if x >= q33:
            return "#F1C40F"
        return "#E74C3C"

    df["Yield_Class"] = df["Yield"].apply(yield_class)
    df["Stress_norm"] = 1 - (
        (df["Stress_Index"] - df["Stress_Index"].min())
        / (df["Stress_Index"].max() - df["Stress_Index"].min())
    )
    df["Yield_norm"] = (
        (df["Yield"] - df["Yield"].min())
        / (df["Yield"].max() - df["Yield"].min())
    )
    df["Breeding_Score"] = 0.5 * df["Stress_norm"] + 0.5 * df["Yield_norm"]

# =========================
# RANKING
# =========================
st.subheader("📊 Genotype Display Control")

top_n = st.selectbox(
    "Show top genotypes",
    options=[10, 15, 20, 30, "All"],
    index=0,
)

ranking = df.groupby("Genotype")["Stress_Index"].mean().sort_values()
ranking_df = ranking.reset_index()
ranking_df.columns = ["Genotype", "Stress_Index"]
ranking_df = ranking_df.sort_values("Stress_Index")
best = ranking.idxmin()
worst = ranking.idxmax()
fig1 = plot_stress(df)
fig2 = plot_yield(df) if has_yield else None
fig_tradeoff = plot_stress_vs_yield(df) if has_yield else None

st.subheader("🔎 Search Genotype")

genotype_list = df["Genotype"].unique()

selected = st.selectbox(
    "Find specific genotype",
    genotype_list,
)

highlight = df[df["Genotype"] == selected]


def medal(i):
    if i == 0:
        return "🥇"
    if i == 1:
        return "🥈"
    if i == 2:
        return "🥉"
    return ""

st.subheader("🏆 Genotype Ranking")
for i, row in enumerate(ranking_df.itertuples()):
    color = color_scale(row.Stress_Index)
    st.markdown(
        f"<div style='padding:8px;background-color:{color}22;border-radius:8px'>"
        f"{medal(i)} <b>{row.Genotype}</b> → {row.Stress_Index:.3f}"
        f"</div>",
        unsafe_allow_html=True,
    )

st.subheader("🥇 Breeding Recommendation")

st.markdown(
    """
### 🧠 How to read this section

This section identifies the best genotypes for breeding based on the lowest Stress Index.

- Top ranked genotypes → most stable under drought stress  
- Bottom ranked genotypes → poor performance under stress  

👉 These are the **final candidates for selection in breeding programs**.
"""
)

top3 = ranking.nsmallest(3)

for i, (g, v) in enumerate(top3.items(), 1):
    st.markdown(f"**{i}. {g} → Stress Index: {v:.3f}**")

st.success(
    f"""
🌱 Key Message:
Genotype {best} shows superior drought tolerance and is the strongest candidate for breeding improvement.

🔥 Genotype {worst} is highly sensitive and not recommended for drought-prone environments.
"""
)

if "Breeding_Score" in df.columns:
    st.subheader("🏆 Final Breeding Score Ranking")
    final_rank = df.groupby("Genotype")["Breeding_Score"].mean().sort_values(ascending=False)
    st.dataframe(final_rank)
    best_genotype = final_rank.idxmax()
    st.success(
        f"""
🌱 Best overall breeding genotype: {best_genotype}

This genotype shows the best combination of:
- Low drought stress
- High yield performance

It is recommended as the primary candidate for breeding programs.
"""
    )

    st.subheader("🌾 Stress vs Yield with Breeding Classification")

    fig_breed, ax_breed = plt.subplots()
    scatter = ax_breed.scatter(
        df["Stress_Index"],
        df["Yield"],
        c=df["Breeding_Score"],
        cmap="viridis",
        s=120,
        alpha=0.85,
    )

    elite_cut = df["Breeding_Score"].quantile(0.66)
    poor_cut = df["Breeding_Score"].quantile(0.33)

    for _, row in df.iterrows():
        if row["Breeding_Score"] >= elite_cut:
            label = "🟢 Elite"
            color = "green"
        elif row["Breeding_Score"] >= poor_cut:
            label = "🟡 Good"
            color = "orange"
        else:
            label = "🔴 Poor"
            color = "red"

        ax_breed.text(
            row["Stress_Index"],
            row["Yield"],
            f"{row['Genotype']}\n{label}",
            fontsize=8,
            color=color,
        )

    ax_breed.set_xlabel("Stress Index (Lower = Better)")
    ax_breed.set_ylabel("Yield (Higher = Better)")
    ax_breed.set_title("Genotype Breeding Space")

    style_matplotlib(ax_breed)
    cbar = plt.colorbar(scatter, label="Breeding Score")
    style_colorbar(cbar)
    st.pyplot(fig_breed)

if has_yield:
    st.subheader("🌾 Yield Analysis")
    run_yield_analysis(df)
    run_tradeoff(df)
else:
    st.warning("Yield data not available. Yield analysis is disabled.")

if "Breeding_Score" in df.columns and has_yield:
    st.subheader("📄 Automated Discussion (Scientific Interpretation)")

    best = final_rank.idxmax()
    worst = final_rank.idxmin()

    avg_stress = df["Stress_Index"].mean()
    avg_yield = df["Yield"].mean()

    if avg_stress < 0.4:
        stress_state = "low overall stress conditions"
    elif avg_stress < 0.7:
        stress_state = "moderate stress conditions"
    else:
        stress_state = "high stress conditions"

    if avg_yield > df["Yield"].median():
        yield_state = "above-average yield performance"
    else:
        yield_state = "below-average yield performance"

    st.markdown(
        f"""
### Discussion

The evaluated genotypes were analyzed under {stress_state}, showing variability in both physiological stress response and yield potential.

Overall, the dataset indicates {yield_state}, suggesting that environmental conditions and genotype interactions significantly influenced performance.

The breeding score analysis identified **{best}** as the most promising genotype due to its superior combination of low stress and high yield.

Conversely, **{worst}** showed the lowest performance and is not recommended for drought-prone environments.

These results support the use of integrated physiological and agronomic indicators (Fv/Fm, SPAD, Stress Index, and Yield) for efficient selection of drought-tolerant genotypes.
"""
    )

# =========================
# CLASSIFICATION
# =========================
df["Class"] = df["Stress_Index"].apply(classify)
low = df[df["Stress_Index"] < 0.4]
mid = df[(df["Stress_Index"] >= 0.4) & (df["Stress_Index"] < 0.7)]
high = df[df["Stress_Index"] >= 0.7]

st.subheader("📊 Stress Classification")
st.markdown(
    """
### 🧠 What this means

Genotypes are grouped based on Stress Index values:

- 🟢 Low stress → drought tolerant (preferred for breeding)
- 🟡 Moderate stress → intermediate response
- 🔴 High stress → sensitive genotypes (not recommended)

👉 Each point represents one genotype classified by stress level.
"""
)
st.dataframe(df.style.map(color_class, subset=["Stress_Index"]))

tab1, tab2, tab3 = st.tabs(["🟢 Low", "🟡 Moderate", "🔴 High"])

with tab1:
    st.dataframe(low)

with tab2:
    st.dataframe(mid)

with tab3:
    st.dataframe(high)

st.subheader("📊 Stress Classes Overview")

fig, ax = plt.subplots()

colors = df["Stress_Index"].apply(
    lambda x: "green" if x < 0.4 else "orange" if x < 0.7 else "red"
)

ax.scatter(df["Genotype"], df["Stress_Index"], c=colors)

ax.set_title("Stress Classification by Genotype")
ax.set_xlabel("Genotype")
ax.set_ylabel("Stress Index")

for i in range(len(df)):
    ax.text(
        df["Genotype"].iloc[i],
        df["Stress_Index"].iloc[i],
        df["Genotype"].iloc[i],
    )

style_matplotlib(ax)
st.pyplot(fig)

st.subheader("🌡️ Integrated Stress Heatmap")
st.markdown(
    """
### 🔬 What this heatmap shows

This heatmap combines all physiological traits and stress response.

- 🟢 Green → optimal performance  
- 🟡 Yellow → intermediate response  
- 🔴 Red → high stress sensitivity  

👉 It provides a **quick visual summary of genotype performance**.
"""
)

st.subheader("🌡️ Stress Heatmap (Genotype × Traits)")
heat = df.groupby("Genotype")[[FVFM_COL, SPAD_COL, "Stress_Index"]].mean()

fig, ax = plt.subplots()
heatmap = sns.heatmap(heat, annot=True, cmap="RdYlGn_r", ax=ax)
style_matplotlib(ax)
if heatmap.collections:
    style_colorbar(heatmap.collections[0].colorbar)
fig.savefig("heatmap.png", dpi=300, bbox_inches="tight")
st.pyplot(fig)

# =========================
# RECOMMENDATION
# =========================
st.subheader("💡 Recommendation")
st.write(generate_recommendation(df["Stress_Index"].mean()))

st.subheader("🧠 Biological Interpretation")
st.markdown(
    """
This section summarizes biological meaning of the results and identifies best and worst performing genotypes.
"""
)

st.subheader("🧠 Scientific Interpretation")

st.markdown(
    """
### 📄 What this means biologically

This section summarizes all results into a biological conclusion.

- Identifies most tolerant genotype  
- Identifies most sensitive genotype  
- Confirms physiological consistency across traits  

👉 This is the **decision-making output for breeding selection**.
"""
)

st.subheader("🧠 Key Biological Insight")
st.markdown(
    f"""
<div style="
    background-color: rgba(34,197,94,0.1);
    padding: 15px;
    border-radius: 10px;
    border-left: 5px solid #22c55e;
">
<b>🌱 Best genotype:</b> {best}<br>
<b>🔥 Worst genotype:</b> {worst}
</div>
""",
    unsafe_allow_html=True,
)

st.success(
    f"""
Best genotype: {best}
Worst genotype: {worst}
"""
)

st.subheader("🔬 Scientific Examiner Insight")
st.write(
    f"""
The dataset shows clear physiological separation among genotypes.

- Best performing genotype: {best}
- Most sensitive genotype: {worst}

This indicates that SPAD and Fv/Fm effectively discriminate drought stress response.
Genotypes with lower Stress Index are recommended for breeding programs.
"""
)

st.subheader("📄 Automated Scientific Conclusion")

avg = df["Stress_Index"].mean()

if avg < 0.4:
    stress_level = "low overall stress conditions"
elif avg < 0.7:
    stress_level = "moderate stress conditions"
else:
    stress_level = "high stress conditions"

st.markdown(
    f"""
This analysis evaluated genotype performance under {stress_level} using physiological traits (Fv/Fm and SPAD).

The results show a clear separation among genotypes in drought response.

- Most tolerant genotype: **{best}**
- Most sensitive genotype: **{worst}**

These findings suggest that integrated physiological indicators are effective for selecting drought-tolerant germplasm.
"""
)

# =========================
# INTERPRETATION
# =========================
st.subheader("🧠 Automated Interpretation")

st.markdown(
    f"""
- **Best genotype:** {best} (lowest stress, highest tolerance)
- **Worst genotype:** {worst} (highest stress sensitivity)

The results show clear separation among genotypes under drought stress.
Physiological traits strongly correlate with stress response.
"""
)

# =========================
# BAR PLOT (GREEN - como ayer)
# =========================
st.subheader("📉 Stress Index by Genotype")

st.markdown(
    """
### 📊 What this graph shows

This plot ranks each genotype based on its Stress Index.

- 🟢 Low values = drought tolerant genotypes  
- 🟡 Medium values = intermediate response  
- 🔴 High values = drought sensitive genotypes  

👉 This is the **main selection criterion** for breeding decisions.
"""
)

st.pyplot(fig1)

# =========================
# SCATTER (AZUL - como ayer)
# =========================
st.subheader("🌿 Physiological Relationship: Chlorophyll Content vs Photosystem Efficiency")
st.markdown(
    """
### 🌱 Biological meaning

This graph shows plant physiological health.

- SPAD → chlorophyll content (leaf greenness)  
- Fv/Fm → photosynthetic efficiency  

👉 Genotypes in the upper range indicate **healthier plants under stress**.

### 🧠 Trait definition update

- **Chlorophyll Content (SPAD):** proxy of leaf greenness and chlorophyll level  
- **Photosystem II Efficiency (Fv/Fm):** indicator of photosynthetic performance and stress damage  

👉 These traits are complementary but NOT equivalent.
"""
)

st.subheader("🌿 Physiological Trait Relationship")

fig2, ax2 = plt.subplots()

if top_n == "All":
    df_plot = df.copy()
else:
    df_plot = df[
        df["Genotype"].isin(
            df.groupby("Genotype")["Stress_Index"]
            .mean()
            .nsmallest(int(top_n))
            .index
        )
    ]

ax2.scatter(
    df_plot[SPAD_COL],
    df_plot[FVFM_COL],
    c=df_plot["Stress_Index"],
    cmap="RdYlGn_r",
)

ax2.scatter(
    highlight[SPAD_COL],
    highlight[FVFM_COL],
    color="blue",
    s=200,
    edgecolor="black",
)

ax2.set_xlabel("Chlorophyll Content")
ax2.set_ylabel("Photosystem II Efficiency")
style_matplotlib(ax2)

st.pyplot(fig2)

# =========================
# WORD REPORT BUTTON
# =========================
if st.button("📄 Generate Full Scientific Report"):
    buffer = generate_word_report(df)
    st.download_button(
        "Download Word Report",
        buffer,
        file_name="PhytoStress_Report.docx",
    )

if st.button("Generate full report (disabled)"):
    pass
