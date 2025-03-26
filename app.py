import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="Vyhodnocení laboratorního deníku")
st.title("Vyhodnocení laboratorního deníku")

pdf_file = st.file_uploader("Nahraj laboratorní deník (PDF)", type="pdf")
xlsx_file = st.file_uploader("Nahraj PROMT.xlsx", type="xlsx")

def build_mapping(typy_row, stanice_row):
    mapping = {}
    for col in typy_row.index:
        typ = typy_row[col]
        stanice = stanice_row[col]
        if pd.notna(typ) and pd.notna(stanice):
            mapping[typ.strip()] = stanice.strip()
    return mapping

def count_tests(text, typ, staniceni):
    search = f"{typ.lower()} {staniceni.lower()}"
    return text.count(search)

def vypln_skutecnosti(df, lab_text, op1_mapping, op2_mapping):
    for i, row in df.iterrows():
        typ = row["Typ zásypu"]
        if pd.isna(typ):
            continue
        typ = typ.strip()
        if typ in op1_mapping:
            df.at[i, "Skutečnost OP1"] = count_tests(lab_text, typ, op1_mapping[typ])
        if typ in op2_mapping:
            df.at[i, "Skutečnost OP2"] = count_tests(lab_text, typ, op2_mapping[typ])
    return df

if pdf_file and xlsx_file:
    pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
    lab_text = "\n".join(page.get_text() for page in pdf).lower()

    xls = pd.ExcelFile(xlsx_file)
    pm_df = pd.read_excel(xls, sheet_name="PM")
    lm_df = pd.read_excel(xls, sheet_name="LM")
    op1_df = pd.read_excel(xls, sheet_name="seznam zkoušek PM+LM OP1 ")
    op2_df = pd.read_excel(xls, sheet_name="seznam zkoušek PM+LM OP2")

    op1_mapping = build_mapping(op1_df.iloc[0], op1_df.iloc[2])
    op2_mapping = build_mapping(op2_df.iloc[0], op2_df.iloc[2])

    pm_df = vypln_skutecnosti(pm_df, lab_text, op1_mapping, op2_mapping)
    lm_df = vypln_skutecnosti(lm_df, lab_text, op1_mapping, op2_mapping)

    st.subheader("Výsledky pro list PM")
    st.dataframe(pm_df)

    st.subheader("Výsledky pro list LM")
    st.dataframe(lm_df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pm_df.to_excel(writer, index=False, sheet_name="PM")
        lm_df.to_excel(writer, index=False, sheet_name="LM")
        op1_df.to_excel(writer, index=False, sheet_name="seznam zkoušek PM+LM OP1 ")
        op2_df.to_excel(writer, index=False, sheet_name="seznam zkoušek PM+LM OP2")

    st.download_button(
        label="📥 Stáhnout výsledný Excel",
        data=output.getvalue(),
        file_name="vyhodnoceni_vystup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
