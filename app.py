import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

plt.style.use("seaborn-v0_8-whitegrid")

# ---------------------------
# CONFIG
# ---------------------------
st.set_page_config(page_title="PhytoStress AI", layout="wide")

st.markdown(
    """
# 🌱 PhytoStress AI  
### Plant stress phenotyping & breeding decision support tool
"""
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        border: 1px solid #E6E6E6;
        padding: 6px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.subheader("🧠 Interpretation Guide")
st.markdown(
    """
### 🎨 Color meaning
🟢 Green → Low stress / healthy plant  
🟡 Yellow → Moderate stress response  
🔴 Red → High stress / low tolerance  
"""
)

st.markdown(
    """
### 📊 Stress Index explanation

The Stress Index is calculated as:

**Stress Index = 1 - (Biomass under stress / Biomass under control)**

- Values close to **0** → plants are **highly tolerant** (low stress impact)  
- Values close to **1** → plants are **highly sensitive** (high stress impact)  

This index represents the overall reduction in plant growth due to drought stress.
"""
)
st.info(
    "Higher SPAD and Fv/Fm values indicate healthier plants. Higher Stress Index indicates lower drought tolerance."
)


# ---------------------------
# CORE FUNCTION
# ---------------------------
def compute_stress(df):
    df = df.copy()
    df["Stress_Index"] = 1 - (df["Biomass_stress"] / df["Biomass_control"])
    return df


def interpret(avg):
    if avg < 0.4:
        return "Low stress detected (high drought tolerance across genotypes)."
    if avg < 0.7:
        return "Moderate stress response observed."
    return "High stress levels detected (low tolerance)."


def classify_stress(x):
    if x < 0.4:
        return "Low"
    if x < 0.7:
        return "Moderate"
    return "High"


def generate_word_report(geno_rank, summary_text, df_result):
    fig1, ax1 = plt.subplots()
    df_plot = geno_rank.reset_index()
    df_plot.columns = ["Genotype", "Stress_Index"]

    ax1.bar(df_plot["Genotype"], df_plot["Stress_Index"], color="#4CAF50")
    ax1.set_title("Figure 1. Stress Index by Genotype")
    ax1.set_ylabel("Stress Index")
    plt.xticks(rotation=45)

    fig_path = "fig1_stress.png"
    plt.savefig(fig_path, dpi=300, bbox_inches="tight")
    plt.close()

    fig2, ax2 = plt.subplots()
    ax2.scatter(df_result["SPAD"], df_result["FvFm"], color="#4C78A8")
    ax2.set_title("Figure 2. SPAD vs Fv/Fm Relationship")
    ax2.set_xlabel("SPAD")
    ax2.set_ylabel("Fv/Fm")
    plt.savefig("fig2_physiology.png", dpi=300, bbox_inches="tight")
    plt.close()

    doc = Document()

    title = doc.add_heading("PhytoStress AI", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph(
        "Drought Stress Phenotyping and Breeding Decision Support System"
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("\n")
    doc.add_paragraph(
        "This report summarizes genotype performance under drought stress conditions "
        "using physiological and biomass-based indicators."
    )

    doc.add_page_break()

    doc.add_heading("Abstract", level=1)
    doc.add_paragraph(summary_text)

    doc.add_heading("1. Introduction", level=1)
    doc.add_paragraph(
        "Drought stress is one of the major limiting factors in crop productivity. "
        "This study evaluates genotype performance using a Stress Index derived from biomass reduction "
        "under controlled and stress conditions."
    )

    doc.add_heading("2. Methodology", level=1)
    doc.add_paragraph(
        "Stress Index was calculated as: 1 - (Biomass_stress / Biomass_control). "
        "Lower values indicate higher drought tolerance."
    )
    doc.add_paragraph(
        "Physiological traits such as SPAD and Fv/Fm were considered as complementary indicators "
        "of plant health status under stress conditions."
    )

    doc.add_heading("Figures", level=1)
    doc.add_paragraph("Figure 1. Stress Index by Genotype")
    doc.add_picture("fig1_stress.png", width=Inches(5.8))
    doc.add_paragraph("Figure 2. Relationship between SPAD and Fv/Fm")
    doc.add_picture("fig2_physiology.png", width=Inches(5.8))

    doc.add_heading("4. Genotype Ranking", level=1)
    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Genotype"
    hdr[1].text = "Stress Index"

    for geno, value in geno_rank.items():
        row = table.add_row().cells
        row[0].text = str(geno)
        row[1].text = f"{value:.3f}"

    doc.add_heading("5. Stress Classification", level=1)
    low = df_result[df_result["Stress_Index"] < 0.4]["Genotype"].unique()
    mid = df_result[
        (df_result["Stress_Index"] >= 0.4) & (df_result["Stress_Index"] < 0.7)
    ]["Genotype"].unique()
    high = df_result[df_result["Stress_Index"] >= 0.7]["Genotype"].unique()

    doc.add_paragraph("Low stress: " + ", ".join(low))
    doc.add_paragraph("Moderate stress response: " + ", ".join(mid))
    doc.add_paragraph("High stress (sensitive genotypes): " + ", ".join(high))

    doc.add_heading("6. Discussion", level=1)
    doc.add_paragraph(
        "Results indicate clear differentiation among genotypes under drought stress conditions. "
        "Genotypes with lower Stress Index values maintained higher biomass and are therefore "
        "strong candidates for drought tolerance breeding programs."
    )

    doc.add_heading("7. Conclusion", level=1)
    doc.add_paragraph(
        "The Stress Index effectively discriminated genotypic responses under drought conditions. "
        "This approach provides a reliable framework for selecting drought-tolerant germplasm."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


# ---------------------------
# COLOR STYLING FUNCTION
# ---------------------------
def style_rows(row):
    styles = [""] * len(row)

    if row["SPAD"] >= 42:
        styles[row.index.get_loc("SPAD")] = "background-color:#b6fcb6"
    elif row["SPAD"] >= 35:
        styles[row.index.get_loc("SPAD")] = "background-color:#fff3a6"
    else:
        styles[row.index.get_loc("SPAD")] = "background-color:#ffb3b3"

    if row["FvFm"] >= 0.80:
        styles[row.index.get_loc("FvFm")] = "background-color:#b6fcb6"
    elif row["FvFm"] >= 0.75:
        styles[row.index.get_loc("FvFm")] = "background-color:#fff3a6"
    else:
        styles[row.index.get_loc("FvFm")] = "background-color:#ffb3b3"

    if row["Stress_Index"] < 0.4:
        styles[row.index.get_loc("Stress_Index")] = "background-color:#b6fcb6"
    elif row["Stress_Index"] < 0.7:
        styles[row.index.get_loc("Stress_Index")] = "background-color:#fff3a6"
    else:
        styles[row.index.get_loc("Stress_Index")] = "background-color:#ffb3b3"

    return styles


def color_rank(val):
    if val < 0.35:
        return "background-color:#2ecc71; color:white; font-weight:bold"
    if val < 0.6:
        return "background-color:#f1c40f; color:black; font-weight:bold"
    return "background-color:#e74c3c; color:white; font-weight:bold"


def color_map(val):
    if val < 0.25:
        return "background-color: #BFEED3; color: #1b4332; font-weight: 600"
    if val < 0.4:
        return "background-color: #FFF4B3; color: #5c4b00; font-weight: 600"
    if val < 0.55:
        return "background-color: #FFD6A5; color: #5c2b00; font-weight: 600"
    return "background-color: #FFADAD; color: #5c0000; font-weight: 600"


def stress_color(val):
    if val < 0.4:
        return "background-color:#4CAF50; color:white; font-weight:bold"
    if val < 0.7:
        return "background-color:#FFC107; color:black; font-weight:bold"
    return "background-color:#F44336; color:white; font-weight:bold"


# ---------------------------
# DEMO MODE
# ---------------------------
st.subheader("🚀 Demo Mode")

if st.button("Run Full Analysis Demo"):
    data = pd.DataFrame(
        {
            "Genotype": ["G1", "G1", "G2", "G2", "G3", "G3", "G4", "G4", "G5", "G5"],
            "SPAD": [46, 44, 39, 37, 50, 48, 34, 36, 42, 40],
            "FvFm": [0.83, 0.81, 0.75, 0.73, 0.85, 0.84, 0.69, 0.71, 0.78, 0.76],
            "Biomass_control": [10] * 10,
            "Biomass_stress": [8.2, 7.9, 6.0, 5.5, 8.8, 8.5, 4.2, 4.8, 6.8, 6.2],
        }
    )

    st.write("### Input Data")
    st.dataframe(data)

    result = compute_stress(data)
    result["Stress_Class"] = result["Stress_Index"].apply(classify_stress)

    st.write("### Stress Results")
    st.dataframe(result.style.apply(style_rows, axis=1))

    avg = result["Stress_Index"].mean()

    st.metric("Average Stress Index", f"{avg:.3f}")
    st.write(interpret(avg))

    avg_stress = result["Stress_Index"].mean()

    st.subheader("📊 Stress Summary")
    st.write(f"Average Stress Index: {avg_stress:.2f}")

    if avg_stress < 0.4:
        st.success("Low stress detected – High drought tolerance across genotypes")
    elif avg_stress < 0.7:
        st.warning("Moderate stress detected – Mixed genotype responses")
    else:
        st.error("High stress detected – Sensitive population")

    st.subheader("📊 Stress Overview & Classification")
    st.dataframe(
        result.sort_values("Stress_Index").style.map(
            stress_color,
            subset=["Stress_Index"],
        ),
        use_container_width=True,
    )

    st.subheader("🧬 Stress Classification by Genotype")
    st.caption("Click each category to explore grouped genotypes")
    tab1, tab2, tab3 = st.tabs(
        [
            "🟢 Low Stress",
            "🟡 Moderate Stress",
            "🔴 High Stress",
        ]
    )

    low = result[result["Stress_Index"] < 0.4]
    mid = result[(result["Stress_Index"] >= 0.4) & (result["Stress_Index"] < 0.7)]
    high = result[result["Stress_Index"] >= 0.7]

    with tab1:
        st.dataframe(low)

    with tab2:
        st.dataframe(mid)

    with tab3:
        st.dataframe(high)

    # ---------------------------
    # RANKING (BREEDING BASE)
    # ---------------------------
    st.markdown("## 🏆 Genotype Ranking")
    st.markdown("### 🧬 Breeding Priority (Best → Worst)")
    geno_rank = result.groupby("Genotype")["Stress_Index"].mean().sort_values()
    rank_df = geno_rank.reset_index()
    rank_df.columns = ["Genotype", "Stress_Index"]
    rank_df = rank_df.sort_values("Stress_Index", ascending=True).reset_index(drop=True)
    rank_df.insert(0, "Rank", range(1, len(rank_df) + 1))
    styled_rank = rank_df.style.map(stress_color, subset=["Stress_Index"])
    st.dataframe(styled_rank, use_container_width=True)

    fig, ax = plt.subplots()
    df_plot = rank_df.sort_values("Stress_Index")

    colors = []
    for v in df_plot["Stress_Index"]:
        if v < 0.4:
            colors.append("#4CAF50")
        elif v < 0.7:
            colors.append("#FFC107")
        else:
            colors.append("#F44336")

    bars = ax.bar(df_plot["Genotype"], df_plot["Stress_Index"], color=colors)

    for bar in bars:
        h = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            h,
            f"{h:.2f}",
            ha="center",
            va="bottom",
            fontsize=9,
        )

    ax.set_title("Stress Index by Genotype")
    ax.set_ylabel("Stress Index")
    st.pyplot(fig)

    best = result.loc[result["Stress_Index"].idxmin()]["Genotype"]
    worst = result.loc[result["Stress_Index"].idxmax()]["Genotype"]

    st.markdown("## 🏁 Performance Summary")
    st.info(f"🌱 Best genotype: {best}")
    st.warning(f"⚠️ Worst genotype: {worst}")
    st.success("✔ Dataset successfully classified under drought stress system")

    st.subheader("📄 Scientific Summary")
    best_genotypes = geno_rank[geno_rank < 0.4].index.tolist()
    worst_genotypes = geno_rank[geno_rank >= 0.7].index.tolist()
    summary_text = "Based on physiological and biomass-based stress analysis, "

    if len(best_genotypes) > 0:
        summary_text += (
            f"genotypes {', '.join(best_genotypes)} showed the lowest stress index values, "
            "indicating high drought tolerance and strong potential for breeding programs. "
        )

    if len(worst_genotypes) > 0:
        summary_text += (
            f"In contrast, genotypes {', '.join(worst_genotypes)} exhibited high stress sensitivity "
            "and are not recommended for selection under drought conditions. "
        )

    summary_text += (
        "Overall, stress index successfully discriminated genotypic performance under water-limited conditions."
    )
    st.info(summary_text)
    st.subheader("📄 Download Full Report")
    word_file = generate_word_report(geno_rank, summary_text, result)
    st.download_button(
        label="⬇️ Download Word Report",
        data=word_file,
        file_name="PhytoStress_AI_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    # ---------------------------
    # BREEDING MODULE
    # ---------------------------
    st.subheader("🌱 Breeding Recommendation System")
    elite = result[result["Stress_Index"] < 0.4]
    intermediate = result[
        (result["Stress_Index"] >= 0.4) & (result["Stress_Index"] < 0.7)
    ]
    sensitive = result[result["Stress_Index"] >= 0.7]

    st.write("### 🟢 Elite Candidates")
    st.write(elite["Genotype"].tolist())

    st.write("### 🟡 Intermediate Candidates")
    st.write(intermediate["Genotype"].tolist())

    st.write("### 🔴 Sensitive Candidates")
    st.write(sensitive["Genotype"].tolist())

    elite_genotypes = elite["Genotype"].drop_duplicates().tolist()
    if elite_genotypes:
        st.success(f"Recommended breeding lines: {', '.join(elite_genotypes)}")

    # ---------------------------
    # VISUAL SUMMARY
    # ---------------------------
    st.subheader("📊 Stress Overview")

    fig, ax = plt.subplots()
    geno_rank.plot(kind="bar", ax=ax)
    st.pyplot(fig)

# ---------------------------
# CSV INPUT
# ---------------------------
st.subheader("📂 Upload Dataset")

file = st.file_uploader("Upload CSV", type=["csv"])

if file is not None:
    df = pd.read_csv(file)

    st.write("### Input Data")
    st.dataframe(df)

    result = compute_stress(df)
    result["Stress_Class"] = result["Stress_Index"].apply(classify_stress)

    st.write("### Results")
    st.dataframe(result.style.apply(style_rows, axis=1))

    avg_stress = result["Stress_Index"].mean()

    st.subheader("📊 Stress Summary")
    st.write(f"Average Stress Index: {avg_stress:.2f}")

    if avg_stress < 0.4:
        st.success("Low stress detected – High drought tolerance across genotypes")
    elif avg_stress < 0.7:
        st.warning("Moderate stress detected – Mixed genotype responses")
    else:
        st.error("High stress detected – Sensitive population")

    st.subheader("📊 Stress Overview & Classification")
    st.dataframe(
        result.sort_values("Stress_Index").style.map(
            stress_color,
            subset=["Stress_Index"],
        ),
        use_container_width=True,
    )

    st.subheader("🧬 Stress Classification by Genotype")
    st.caption("Click each category to explore grouped genotypes")
    tab1, tab2, tab3 = st.tabs(
        [
            "🟢 Low Stress",
            "🟡 Moderate Stress",
            "🔴 High Stress",
        ]
    )

    low = result[result["Stress_Index"] < 0.4]
    mid = result[(result["Stress_Index"] >= 0.4) & (result["Stress_Index"] < 0.7)]
    high = result[result["Stress_Index"] >= 0.7]

    with tab1:
        st.dataframe(low)

    with tab2:
        st.dataframe(mid)

    with tab3:
        st.dataframe(high)

    geno_rank = result.groupby("Genotype")["Stress_Index"].mean().sort_values()

    st.markdown("## 🏆 Genotype Ranking")
    st.markdown("### 🧬 Breeding Priority (Best → Worst)")
    rank_df = geno_rank.reset_index()
    rank_df.columns = ["Genotype", "Stress_Index"]
    rank_df = rank_df.sort_values("Stress_Index", ascending=True).reset_index(drop=True)
    rank_df.insert(0, "Rank", range(1, len(rank_df) + 1))
    styled_rank = rank_df.style.map(stress_color, subset=["Stress_Index"])
    st.dataframe(styled_rank, use_container_width=True)

    fig, ax = plt.subplots()
    df_plot = rank_df.sort_values("Stress_Index")

    colors = []
    for v in df_plot["Stress_Index"]:
        if v < 0.4:
            colors.append("#4CAF50")
        elif v < 0.7:
            colors.append("#FFC107")
        else:
            colors.append("#F44336")

    bars = ax.bar(df_plot["Genotype"], df_plot["Stress_Index"], color=colors)

    for bar in bars:
        h = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            h,
            f"{h:.2f}",
            ha="center",
            va="bottom",
            fontsize=9,
        )

    ax.set_title("Stress Index by Genotype")
    ax.set_ylabel("Stress Index")
    st.pyplot(fig)

    st.write("### Breeding Insight")

    best = result.loc[result["Stress_Index"].idxmin()]["Genotype"]
    worst = result.loc[result["Stress_Index"].idxmax()]["Genotype"]
    st.markdown("## 🏁 Performance Summary")
    st.info(f"🌱 Best genotype: {best}")
    st.warning(f"⚠️ Worst genotype: {worst}")
    st.success("✔ Dataset successfully classified under drought stress system")

    st.subheader("📄 Scientific Summary")
    best_genotypes = geno_rank[geno_rank < 0.4].index.tolist()
    worst_genotypes = geno_rank[geno_rank >= 0.7].index.tolist()
    summary_text = "Based on physiological and biomass-based stress analysis, "

    if len(best_genotypes) > 0:
        summary_text += (
            f"genotypes {', '.join(best_genotypes)} showed the lowest stress index values, "
            "indicating high drought tolerance and strong potential for breeding programs. "
        )

    if len(worst_genotypes) > 0:
        summary_text += (
            f"In contrast, genotypes {', '.join(worst_genotypes)} exhibited high stress sensitivity "
            "and are not recommended for selection under drought conditions. "
        )

    summary_text += (
        "Overall, stress index successfully discriminated genotypic performance under water-limited conditions."
    )
    st.info(summary_text)
    st.subheader("📄 Download Full Report")
    word_file = generate_word_report(geno_rank, summary_text, result)
    st.download_button(
        label="⬇️ Download Word Report",
        data=word_file,
        file_name="PhytoStress_AI_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------------------------
# EXPORT NOTE
# ---------------------------
st.subheader("📄 Export")
st.info("Use screenshot or browser export for submission.")
