import io

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

# =========================
# APP CONFIG
# =========================
st.set_page_config(page_title="Plant Stress AI", layout="wide")

st.title("🌱 Plant Stress AI Analyzer")

st.markdown(
    """
### Drought stress phenotyping tool

Upload your Excel file with:
- Genotype
- FvFm
- SPAD

The app will calculate Stress Index, rank genotypes, and generate a scientific report.
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


st.subheader("📥 Download Excel Template")

template = pd.DataFrame(
    {
        "Genotype": ["H1", "H1", "H2"],
        "FvFm": [0.78, 0.75, 0.60],
        "SPAD": [42, 40, 30],
    }
)

st.download_button(
    "Download Template",
    data=convert_df(template),
    file_name="plant_stress_template.xlsx",
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


# =========================
# MAIN APP FLOW
# =========================
st.subheader("🚀 Input Mode")

mode = st.radio("Choose data source:", ["📊 Demo Data", "📤 Upload Excel"])

if mode == "📊 Demo Data":
    df = pd.DataFrame(
        {
            "Genotype": ["H1", "H1", "H2", "H3"],
            "FvFm": [0.78, 0.75, 0.60, 0.50],
            "SPAD": [42, 40, 30, 25],
        }
    )
else:
    file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if file is None:
        st.stop()

    df = pd.read_excel(file)

required = ["Genotype", "FvFm", "SPAD"]

if any(col not in df.columns for col in required):
    st.error("Missing required columns: Genotype, FvFm, SPAD")
    st.stop()

df = compute_stress(df)

# =========================
# RANKING
# =========================
ranking = df.groupby("Genotype")["Stress_Index"].mean().sort_values()

st.subheader("🏆 Genotype Ranking (Lower = Better)")
st.dataframe(ranking)

# =========================
# CLASSIFICATION
# =========================
df["Class"] = df["Stress_Index"].apply(classify)

st.subheader("📊 Stress Classification")
st.dataframe(df)

# =========================
# RECOMMENDATION
# =========================
st.subheader("💡 Recommendation")
st.write(generate_recommendation(df["Stress_Index"].mean()))

# =========================
# BAR PLOT (GREEN - como ayer)
# =========================
st.subheader("📉 Stress Index by Genotype")

fig1, ax1 = plt.subplots()
ax1.bar(ranking.index, ranking.values, color="#4CAF50")
ax1.set_ylabel("Stress Index")
ax1.set_title("Genotype Performance")
plt.xticks(rotation=45)
st.pyplot(fig1)

# =========================
# SCATTER (AZUL - como ayer)
# =========================
st.subheader("🌿 SPAD vs Fv/Fm")

fig2, ax2 = plt.subplots()
ax2.scatter(df["SPAD"], df["FvFm"], color="#4C78A8")
ax2.set_xlabel("SPAD")
ax2.set_ylabel("Fv/Fm")
ax2.set_title("Physiological relationship")
st.pyplot(fig2)

# =========================
# WORD REPORT BUTTON
# =========================
if st.button("Generate Word Report"):
    doc = Document()

    title = doc.add_heading("PhytoStress AI Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("Drought Stress Phenotyping Report")

    doc.add_heading("Genotype Ranking", level=1)
    for i, (g, v) in enumerate(ranking.items(), 1):
        doc.add_paragraph(f"{i}. {g}: {v:.3f}")

    fig1_path = "fig1.png"
    fig1.savefig(fig1_path, dpi=300, bbox_inches="tight")
    doc.add_picture(fig1_path, width=Inches(5.8))

    fig2_path = "fig2.png"
    fig2.savefig(fig2_path, dpi=300, bbox_inches="tight")
    doc.add_picture(fig2_path, width=Inches(5.8))

    low = df[df["Stress_Index"] < 0.4]["Genotype"].unique()
    mid = df[
        (df["Stress_Index"] >= 0.4) & (df["Stress_Index"] < 0.7)
    ]["Genotype"].unique()
    high = df[df["Stress_Index"] >= 0.7]["Genotype"].unique()

    doc.add_heading("Stress Classification", level=1)
    doc.add_paragraph("Low stress (tolerant): " + ", ".join(low))
    doc.add_paragraph("Moderate stress: " + ", ".join(mid))
    doc.add_paragraph("High stress (sensitive): " + ", ".join(high))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        "Download Word Report",
        buffer,
        file_name="PhytoStress_Report.docx",
    )
