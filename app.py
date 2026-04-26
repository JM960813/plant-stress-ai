from io import BytesIO
import streamlit as st
import pandas as pd

st.title("Plant Stress AI Analyzer")

st.info(
    """
Your file must contain:
- Hybrid: genotype (H1, H2, etc.)
- FvFm: photosynthetic efficiency
- SPAD: chlorophyll content

The app will automatically calculate Stress Index from these variables.
"""
)


def convert_df(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


template = pd.DataFrame(
    {
        "Plant_ID": [1, 2],
        "FvFm": [0.78, 0.65],
        "SPAD": [42, 35],
    }
)

excel_file = convert_df(template)

st.download_button(
    label="Download Excel template",
    data=excel_file,
    file_name="template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

required_cols = ["Hybrid", "FvFm", "SPAD"]

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.write("Preview of data:")
    st.dataframe(df)

    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        st.error(f"Missing columns: {missing}")
        st.info("Required format: Hybrid, FvFm, SPAD")
        st.stop()
    else:
        st.success("File format is correct ✔️")

        df["FvFm_norm"] = df["FvFm"] / df["FvFm"].max()
        df["SPAD_norm"] = df["SPAD"] / df["SPAD"].max()
        df["Stress_Index"] = 1 - (df["FvFm_norm"] + df["SPAD_norm"]) / 2

        st.write("Basic stats:")
        st.write(df[required_cols].describe())

        st.write("Calculated stress results:")
        st.dataframe(df)

        if "Hybrid" in df.columns:
            summary = df.groupby("Hybrid")[["Stress_Index", "FvFm", "SPAD"]].mean()
            st.write("Hybrid summary:")
            st.dataframe(summary)
            ranking = summary.sort_values("Stress_Index")
            st.write("Hybrid ranking:")
            st.dataframe(ranking)
