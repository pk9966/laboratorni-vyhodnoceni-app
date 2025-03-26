import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="Vyhodnocení laboratorního deníku")
st.title("Vyhodnocení laboratorního deníku")

pdf_file = st.file_uploader("Nahraj laboratorní deník (PDF)", type="pdf")
xlsx_file = st.file_uploader("Nahraj soubor Klíč.xlsx", type="xlsx")

def count_matches(text, *terms):
    return sum(1 for line in text.splitlines() if all(term.lower() in line.lower() for term in terms))

def process_op_sheet(key_df, target_df, lab_text):
    for i, row in target_df.iterrows():
        typ = row.get("Typ zásypu")
        if pd.isna(typ):
            continue
        matches = key_df[key_df.get("typ zásypu") == typ]
        count = 0
        for _, match_row in matches.iterrows():
            konstrukce = match_row.get("konstrukční prvek", "")
            zkouska = match_row.get("druh zkoušky", "")
            staniceni = str(match_row.get("staničení", ""))
            if konstrukce and zkouska and staniceni:
                count += count_matches(lab_text, konstrukce, zkouska, staniceni)
        target_df.at[i, "D"] = count
        pozadovano = row.get("C")
        if pd.notna(pozadovano):
            if count >= pozadovano:
                target_df.at[i, "E"] = "Vyhovující"
            else:
                target_df.at[i, "E"] = f"Chybí {abs(int(pozadovano - count))} zk."
    return target_df

def process_cely_objekt_sheet(key_df, target_df, lab_text):
    for i, row in target_df.iterrows():
        material = row.get("materiál")
        zkouska = row.get("druh zkoušky")
        if pd.isna(material) or pd.isna(zkouska):
            continue
        count = count_matches(lab_text, material, zkouska)
        target_df.at[i, "C"] = count
        pozadovano = row.get("B")
        if pd.notna(pozadovano):
            if count >= pozadovano:
                target_df.at[i, "D"] = "Vyhovující"
            else:
                target_df.at[i, "D"] = f"Chybí {abs(int(pozadovano - count))} zk."
    return target_df

if pdf_file and xlsx_file:
    lab_text = "\n".join(page.get_text() for page in fitz.open(stream=pdf_file.read(), filetype="pdf"))

    xls = pd.ExcelFile(xlsx_file)

    op1_key = pd.read_excel(xls, sheet_name="seznam zkoušek PM+LM OP1")
    op2_key = pd.read_excel(xls, sheet_name="seznam zkoušek PM+LM OP2")
    cely_key = pd.read_excel(xls, sheet_name="seznam zkoušek Celý objekt")

    pm_op1 = pd.read_excel(xls, sheet_name="PM - OP1")
    lm_op1 = pd.read_excel(xls, sheet_name="LM - OP1")
    pm_op2 = pd.read_excel(xls, sheet_name="PM - OP2")
    lm_op2 = pd.read_excel(xls, sheet_name="LM - OP2")
    cely_objekt = pd.read_excel(xls, sheet_name="Celý objekt")

    pm_op1 = process_op_sheet(op1_key, pm_op1, lab_text)
    lm_op1 = process_op_sheet(op1_key, lm_op1, lab_text)
    pm_op2 = process_op_sheet(op2_key, pm_op2, lab_text)
    lm_op2 = process_op_sheet(op2_key, lm_op2, lab_text)
    cely_objekt = process_cely_objekt_sheet(cely_key, cely_objekt, lab_text)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pm_op1.to_excel(writer, index=False, sheet_name="PM - OP1")
        lm_op1.to_excel(writer, index=False, sheet_name="LM - OP1")
        pm_op2.to_excel(writer, index=False, sheet_name="PM - OP2")
        lm_op2.to_excel(writer, index=False, sheet_name="LM - OP2")
        cely_objekt.to_excel(writer, index=False, sheet_name="Celý objekt")

    st.success("Vyhodnocení dokončeno. Stáhni výsledný soubor níže.")
    st.download_button(
        label="📥 Stáhnout výsledný Excel",
        data=output.getvalue(),
        file_name="vyhodnoceni_vystup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
