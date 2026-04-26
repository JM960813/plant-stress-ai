import io

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

plt.style.use("dark_background")

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

st.markdown(
    """
# 🌱 PhytoStress AI
### 🌾 Drought Stress Phenotyping & Breeding Intelligence System

"""
)

st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    """
<style>

/* Fondo general */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, #0f172a 0%, #111827 100%);
    color: white;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #0b1220;
}

/* Text color */
html, body, [class*="css"]  {
    color: white;
}

/* Cards estilo */
.block-container {
    padding: 2rem 3rem;
}

/* Dataframes */
div[data-testid="stDataFrame"] {
    background-color: #111827;
    border-radius: 12px;
    padding: 10px;
}

/* Buttons */
.stButton>button {
    background-color: #22c55e;
    color: white;
    border-radius: 10px;
    padding: 0.5rem 1rem;
    border: none;
}

.stButton>button:hover {
    background-color: #16a34a;
}

</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<style>
/* fondo general más estable */
.block-container {
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

st.info("This tool integrates physiological traits with stress modeling to support drought-tolerant genotype selection.")

st.markdown(
    """
## 📌 How to use this app

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

st.info(
    """
This tool is crop-independent.
It can be used for any plant species as long as Fv/Fm, SPAD, and yield data are provided.
"""
)

st.markdown(
    """
## 🧠 Interpretation Guide

### 🎨 Stress classification colors
- 🟢 **Green (Low Stress)** → Healthy plants, high drought tolerance  
- 🟡 **Yellow (Moderate Stress)** → Intermediate stress response  
- 🔴 **Red (High Stress)** → Sensitive plants, low drought tolerance  

---

## 🌿 What do the variables mean?

### 📊 Fv/Fm (Photosynthetic efficiency)
- Measures how efficiently the plant is doing photosynthesis  
- Higher values = healthier plant  
- Lower values = stress damage in photosystem II  

### 🍃 SPAD (Chlorophyll content)
- Estimates chlorophyll level in leaves  
- Higher SPAD = greener, healthier leaves  
- Lower SPAD = nutrient or stress limitation  

---

## 🧪 Stress Index (what the app calculates)
The app combines both variables:

**Higher Stress Index = more drought stress**

- Close to 0 → tolerant genotype  
- Close to 1 → sensitive genotype  
"""
)

st.info("The app automatically classifies genotypes into Low, Moderate, and High stress groups.")


# =========================
# TEMPLATE DOWNLOAD
# =========================
def convert_df(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


st.subheader("📥 Step 1: Download Template")

template = pd.DataFrame(
    {
        "Genotype": ["H1", "H1", "H2"],
        "FvFm": [0.75, 0.72, 0.60],
        "SPAD": [40, 38, 30],
        "Yield": [5.1, 4.9, 4.0],
    }
)

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    template.to_excel(writer, index=False, sheet_name="data")

st.download_button(
    label="⬇ Download Excel Template",
    data=buffer.getvalue(),
    file_name="PhytoStress_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# =========================
# CORE ANALYSIS
# =========================
def compute_stress(df):
    df = df.copy()

    df["FvFm_norm"] = df["FvFm"] / df["FvFm"].max()
    df["SPAD_norm"] = df["SPAD"] / df["SPAD"].max()
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
st.subheader("🚀 Step 2: Try Demo (No file needed)")

df = None

if st.button("Run Demo"):
    df = pd.DataFrame(
        {
            "Genotype": ["H1", "H2", "H3", "H4"],
            "FvFm": [0.78, 0.70, 0.55, 0.60],
            "SPAD": [42, 39, 28, 31],
            "Yield": [5.2, 4.8, 3.9, 4.1],
        }
    )
    st.success("Demo loaded")

if df is None:
    st.subheader("📂 Upload your dataset")
    file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

    if file is None:
        st.stop()

    df = pd.read_excel(file)
    st.success("File uploaded successfully")
    st.dataframe(df)

required = ["Genotype", "FvFm", "SPAD"]

if any(col not in df.columns for col in required):
    st.error("Missing required columns: Genotype, FvFm, SPAD")
    st.stop()

df = df.dropna(subset=["Genotype", "FvFm", "SPAD"])
df["Genotype"] = df["Genotype"].astype(str)

df = compute_stress(df)

if "Yield" not in df.columns:
    st.warning("Yield not found. Using estimated yield model.")
    df["Yield"] = 0.5 * df["FvFm"] + 0.5 * df["SPAD"]

if "Yield" in df.columns:
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
ranking = df.groupby("Genotype")["Stress_Index"].mean().sort_values()
ranking_df = ranking.reset_index()
ranking_df.columns = ["Genotype", "Stress_Index"]
ranking_df = ranking_df.sort_values("Stress_Index")


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

top3 = ranking.nsmallest(3)

st.markdown(
    """
The following genotypes are recommended for drought tolerance breeding programs based on lowest Stress Index values:
"""
)

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

    plt.colorbar(scatter, label="Breeding Score")
    st.pyplot(fig_breed)

if "Yield" in df.columns:
    yield_rank = df.groupby("Genotype")["Yield"].mean().sort_values(ascending=False)
    st.subheader("🌾 Yield Ranking (Relative Performance)")
    st.markdown(
        """
Yield classification is based on within-experiment distribution (percentiles), not fixed thresholds.
"""
    )
    st.dataframe(yield_rank)

    fig_yield, ax_yield = plt.subplots()
    ax_yield.bar(
        yield_rank.index,
        yield_rank.values,
        color=[yield_color(v) for v in yield_rank.values],
    )
    ax_yield.set_title("Yield Performance by Genotype")
    ax_yield.set_ylabel("Yield")

    for i, v in enumerate(yield_rank.values):
        ax_yield.text(i, v, f"{v:.2f}", ha="center", va="bottom")

    st.pyplot(fig_yield)

    best_yield = yield_rank.idxmax()
    worst_yield = yield_rank.idxmin()

    st.subheader("🌾 Yield Insight")
    st.success(f"Best yield genotype: {best_yield}")
    st.error(f"Lowest yield genotype: {worst_yield}")

    st.subheader("🌾 Stress vs Yield Trade-off (Breeding Decision Map)")

    fig_tradeoff, ax_tradeoff = plt.subplots()
    scatter = ax_tradeoff.scatter(
        df["Stress_Index"],
        df["Yield"],
        c=df["Stress_Index"],
        cmap="RdYlGn_r",
        s=100,
        alpha=0.8,
    )

    stress_cut = df["Stress_Index"].quantile(0.33)
    yield_cut = df["Yield"].quantile(0.66)

    ax_tradeoff.axvspan(
        0,
        stress_cut,
        ymin=0.5,
        ymax=1,
        color="green",
        alpha=0.08,
        label="Elite Zone (Low Stress)",
    )

    ax_tradeoff.axhline(
        yield_cut,
        color="green",
        linestyle="--",
        linewidth=1,
        label="High Yield Threshold",
    )

    ax_tradeoff.set_xlabel("Stress Index (Lower = Better)")
    ax_tradeoff.set_ylabel("Yield (Higher = Better)")
    ax_tradeoff.set_title("Genotype Performance Landscape")

    for _, row in df.iterrows():
        ax_tradeoff.text(row["Stress_Index"], row["Yield"], row["Genotype"], fontsize=8)

    plt.colorbar(scatter, label="Stress Level")
    ax_tradeoff.legend()
    fig_tradeoff.savefig("tradeoff.png", dpi=300, bbox_inches="tight")
    st.pyplot(fig_tradeoff)

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

if "Breeding_Score" in df.columns and "Yield" in df.columns:
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
st.dataframe(df.style.map(color_class, subset=["Stress_Index"]))

tab1, tab2, tab3 = st.tabs(["🟢 Low", "🟡 Moderate", "🔴 High"])

with tab1:
    st.dataframe(low)

with tab2:
    st.dataframe(mid)

with tab3:
    st.dataframe(high)

st.subheader("📊 Stress Classes Overview")
st.bar_chart(df["Class"].value_counts())

st.subheader("🌡️ Integrated Stress Profile")
st.markdown(
    """
Heatmap combining physiological traits and Stress Index.
Darker red = higher stress, green = healthier genotypes.
"""
)

st.subheader("🌡️ Stress Heatmap (Genotype × Traits)")
heat_df = df.copy()
heat_df = heat_df.groupby("Genotype")[["FvFm", "SPAD", "Stress_Index"]].mean()

fig_heat, ax_heat = plt.subplots()
sns.heatmap(
    heat_df,
    annot=True,
    cmap="RdYlGn_r",
    linewidths=0.5,
    cbar_kws={"label": "Stress Level"},
    ax=ax_heat,
)
ax_heat.set_title("Physiological and Stress Response Heatmap")
fig_heat.savefig("heatmap.png", dpi=300, bbox_inches="tight")
st.pyplot(fig_heat)

# =========================
# RECOMMENDATION
# =========================
st.subheader("💡 Recommendation")
st.write(generate_recommendation(df["Stress_Index"].mean()))

best = ranking.idxmin()
worst = ranking.idxmax()

st.subheader("🧠 Biological Interpretation")
st.markdown(
    """
This section summarizes biological meaning of the results and identifies best and worst performing genotypes.
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
# BAR PLOT (GREEN - como ayer)
# =========================
st.subheader("📊 Stress Index Distribution")
st.markdown(
    """
Each bar represents the average stress level per genotype.
Green indicates tolerance, red indicates sensitivity.
"""
)

st.subheader("📉 Stress Index by Genotype")

fig1, ax1 = plt.subplots()
bars = ax1.bar(
    ranking.index,
    ranking.values,
    color=[color_scale(v) for v in ranking.values],
)
ax1.set_title("Stress Index by Genotype")
ax1.set_ylabel("Stress Index")

for bar in bars:
    h = bar.get_height()
    ax1.text(
        bar.get_x() + bar.get_width() / 2,
        h,
        f"{h:.2f}",
        ha="center",
        va="bottom",
        fontsize=9,
    )

st.pyplot(fig1)

# =========================
# SCATTER (AZUL - como ayer)
# =========================
st.subheader("🌿 Physiological Trait Relationship")
st.markdown(
    """
This plot shows the relationship between chlorophyll content (SPAD) and photosynthetic efficiency (Fv/Fm).
Genotypes cluster based on physiological performance.
"""
)

st.subheader("🌿 SPAD vs Fv/Fm")

fig2, ax2 = plt.subplots()
for g in df["Genotype"].unique():
    sub = df[df["Genotype"] == g]
    ax2.scatter(
        sub["SPAD"],
        sub["FvFm"],
        label=g,
    )

ax2.set_title("Physiological Response by Genotype")
ax2.set_xlabel("SPAD")
ax2.set_ylabel("Fv/Fm")
ax2.legend()
st.pyplot(fig2)

# =========================
# WORD REPORT BUTTON
# =========================
if st.button("Generate Word Report"):
    doc = Document()

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

    if "Yield" in df.columns:
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
    if "Breeding_Score" in df.columns and "Yield" in df.columns:
        report_best = final_rank.idxmax()
        report_worst = final_rank.idxmin()
        avg_stress = df["Stress_Index"].mean()
        avg_yield = df["Yield"].mean()

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
    if "Breeding_Score" in df.columns and "Yield" in df.columns:
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

    st.download_button(
        "Download Word Report",
        buffer,
        file_name="PhytoStress_Report.docx",
    )

if st.button("📄 Generate Full Report (1 Click)"):
    report = generate_full_report(df, ranking)
    st.download_button(
        label="⬇ Download Word Report",
        data=report,
        file_name="PhytoStress_Full_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
