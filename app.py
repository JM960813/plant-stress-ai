import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Plant Stress AI", layout="wide")

st.title("🌱 Plant Stress AI Analyzer")

st.markdown(
    """
**Drought stress analysis platform**

- Fv/Fm → Photosynthetic efficiency  
- SPAD → Chlorophyll content  
- Stress Index → Computed drought response indicator  
"""
)


def convert_df(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


st.subheader("📥 Download Required Excel Template")

st.markdown(
    """
Fill in this template with your experimental data:

- **Genotype** → e.g., H1, H2, H3  
- **FvFm** → Photosynthetic efficiency  
- **SPAD** → Chlorophyll content  
"""
)

template = pd.DataFrame(
    {
        "Genotype": ["H1", "H1", "H2"],
        "FvFm": [0.78, 0.75, 0.60],
        "SPAD": [42, 40, 30],
    }
)

st.download_button(
    label="Download Excel Template",
    data=convert_df(template),
    file_name="plant_stress_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


def generate_word_report(geno_rank, summary_text, df_result):
    # =========================
    # FIGURE 1: STRESS INDEX BARPLOT
    # =========================
    fig1, ax1 = plt.subplots()

    df_plot = geno_rank.reset_index()
    df_plot.columns = ["Genotype", "Stress_Index"]

    ax1.bar(df_plot["Genotype"], df_plot["Stress_Index"], color="#4CAF50")
    ax1.set_title("Figure 1. Stress Index by Genotype")
    ax1.set_ylabel("Stress Index")
    plt.xticks(rotation=45)

    fig1_path = "fig1.png"
    fig1.savefig(fig1_path, dpi=300, bbox_inches="tight")
    plt.close(fig1)

    # =========================
    # FIGURE 2: PHYSIOLOGY SCATTER
    # =========================
    fig2, ax2 = plt.subplots()

    ax2.scatter(df_result["SPAD"], df_result["FvFm"], color="#4C78A8")

    ax2.set_title("Figure 2. SPAD vs Fv/Fm Relationship")
    ax2.set_xlabel("SPAD")
    ax2.set_ylabel("Fv/Fm")

    fig2_path = "fig2.png"
    fig2.savefig(fig2_path, dpi=300, bbox_inches="tight")
    plt.close(fig2)

    # =========================
    # WORD DOCUMENT
    # =========================
    doc = Document()

    title = doc.add_heading("PhytoStress AI Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph(
        "Drought Stress Phenotyping and Genotype Evaluation"
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("\n")
    doc.add_paragraph(
        "This report presents a physiological and stress-based evaluation of genotypes "
        "under drought conditions using Fv/Fm, SPAD, and a derived Stress Index."
    )

    doc.add_page_break()

    doc.add_heading("Abstract", level=1)
    doc.add_paragraph(summary_text)

    doc.add_heading("1. Introduction", level=1)
    doc.add_paragraph(
        "Drought stress is one of the major limiting factors in crop productivity. "
        "Physiological traits such as chlorophyll content (SPAD) and photosynthetic efficiency (Fv/Fm) "
        "provide insight into plant performance under stress conditions."
    )

    doc.add_heading("2. Methodology", level=1)
    doc.add_paragraph(
        "The Stress Index was calculated as: 1 - (normalized Fv/Fm + normalized SPAD) / 2. "
        "Lower values indicate higher drought tolerance."
    )

    doc.add_heading("3. Results", level=1)
    doc.add_heading("3.1 Genotype Performance", level=2)

    for i, (geno, val) in enumerate(geno_rank.items(), 1):
        doc.add_paragraph(f"{i}. {geno}: {val:.3f}")

    doc.add_heading("Figures", level=1)
    doc.add_paragraph("Figure 1. Stress Index by Genotype")
    doc.add_picture(fig1_path, width=Inches(5.8))

    doc.add_paragraph("Figure 2. Relationship between SPAD and Fv/Fm")
    doc.add_picture(fig2_path, width=Inches(5.8))

    doc.add_heading("4. Stress Classification", level=1)

    low = df_result[df_result["Stress_Index"] < 0.4]["Genotype"].unique()
    mid = df_result[
        (df_result["Stress_Index"] >= 0.4) & (df_result["Stress_Index"] < 0.7)
    ]["Genotype"].unique()
    high = df_result[df_result["Stress_Index"] >= 0.7]["Genotype"].unique()

    doc.add_paragraph("Low stress (tolerant): " + ", ".join(low))
    doc.add_paragraph("Moderate stress: " + ", ".join(mid))
    doc.add_paragraph("High stress (sensitive): " + ", ".join(high))

    doc.add_heading("5. Discussion", level=1)
    doc.add_paragraph(
        "Genotypes with lower Stress Index values exhibited higher drought tolerance, "
        "indicating better maintenance of physiological performance under stress conditions."
    )

    doc.add_heading("6. Conclusion", level=1)
    doc.add_paragraph(
        "The Stress Index effectively discriminated genotype performance under drought stress. "
        "This framework provides a reliable tool for selection of drought-tolerant germplasm."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer

# =========================
# UPLOAD
# =========================
uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    required = ["Genotype", "FvFm", "SPAD"]
    missing = [c for c in required if c not in df.columns]

    if missing:
        st.error(f"Missing columns: {missing}")
        st.stop()

    # =========================
    # STRESS INDEX
    # =========================
    df["FvFm_norm"] = df["FvFm"] / df["FvFm"].max()
    df["SPAD_norm"] = df["SPAD"] / df["SPAD"].max()
    df["Stress_Index"] = 1 - (df["FvFm_norm"] + df["SPAD_norm"]) / 2

    geno_rank = df.groupby("Genotype")["Stress_Index"].mean().sort_values()
    summary_text = (
        "The objective of this study was to evaluate drought stress responses across genotypes "
        "using physiological indicators and a computed Stress Index derived from Fv/Fm and SPAD."
    )

    # =========================
    # TABLE
    # =========================
    st.subheader("📊 Genotype Ranking")
    st.dataframe(geno_rank)

    # =========================
    # BAR CHART (GREEN STYLE - como ayer)
    # =========================
    st.subheader("📉 Stress Index Comparison")

    fig, ax = plt.subplots()
    ax.bar(geno_rank.index, geno_rank.values, color="#4CAF50")
    ax.set_title("Stress Index by Genotype")
    ax.set_ylabel("Stress Index")
    plt.xticks(rotation=45)
    st.pyplot(fig)
    plt.close(fig)

    # =========================
    # SCATTER (BLUE STYLE - como ayer)
    # =========================
    st.subheader("🌿 Physiological Relationship")

    fig2, ax2 = plt.subplots()
    ax2.scatter(df["SPAD"], df["FvFm"], color="#4C78A8")
    ax2.set_xlabel("SPAD")
    ax2.set_ylabel("Fv/Fm")
    ax2.set_title("SPAD vs Fv/Fm")
    st.pyplot(fig2)
    plt.close(fig2)

    if st.button("Generate Word Report"):
        report = generate_word_report(geno_rank, summary_text, df)
        st.success("Report generated!")
        st.download_button(
            "Download Report",
            report,
            file_name="Plant_Stress_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
