import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="Vyhodnocení laboratorního deníku")
st.title("Vyhodnocení laboratorního deníku")

pdf_file = st.file_uploader("Nahraj laboratorní deník (PDF)", type="pdf")
xlsx_file = st.file_uploader("Nahraj soubor Klíč.xlsx", type="xlsx")

def count_matches_advanced(text, konstrukce, zkouska_raw, stanice_raw):
    druhy_zk = [z.strip().lower() for z in str(zkouska_raw).split(",") if z.strip()]
    staniceni = [s.strip().lower() for s in str(stanice_raw).split(",") if s.strip()]
    return sum(
        1 for line in text.splitlines()
        if konstrukce.lower() in line.lower()
        and any(z in line.lower() for z in druhy_zk)
        and any(s in line.lower() for s in staniceni)
    )

def process_op_sheet(key_df, target_df, lab_text):
    if "D" not in target_df.columns:
        target_df["D"] = 0
    if "E" not in target_df.columns:
        target_df["E"] = ""

    for i in range(1, len(target_df)):
        row = target_df.iloc[i]
        zasyp = str(row.iloc[0])
        if pd.isna(zasyp):
            continue
        matches = key_df[key_df.iloc[:, 0] == zasyp]
        total_count = 0
        for _, mrow in matches.iterrows():
            konstrukce = mrow.get("konstrukční prvek", "")
            zkouska = mrow.get("druh zkoušky", "")
            stanice = mrow.get("staničení", "")
            if konstrukce and zkouska and stanice:
                total_count += count_matches_advanced(lab_text, konstrukce, zkouska, stanice)
        target_df.at[i, "D"] = total_count
        pozadovano = row.get("C")
        if pd.notna(pozadovano):
            target_df.at[i, "E"] = "Vyhovující" if total_count >= pozadovano else f"Chybí {abs(int(pozadovano - total_count))} zk."
    return target_df

def process_cely_objekt_sheet(key_df, target_df, lab_text):
    for i, row in target_df.iterrows():
        material = row.get("materiál")
        zkouska = row.get("druh zkoušky")
        if pd.isna(material) or pd.isna(zkouska):
            continue
        count = count_matches_advanced(lab_text, material, zkouska, "")
        target_df.at[i, "C"] = count
        pozadovano = row.get("B")
        if pd.notna(pozadovano):
            target_df.at[i, "D"] = "Vyhovující" if count >= pozadovano else f"Chybí {abs(int(pozadovano - count))} zk."
    return target_df

if pdf_file and xlsx_file:
    lab_text = "\n".join(page.get_text() for page in fitz.open(stream=pdf_file.read(), filetype="pdf"))

    try:
        xls = pd.ExcelFile(xlsx_file)
        sheet_names = xls.sheet_names

        def load_sheet(name):
            if name in sheet_names:
                return pd.read_excel(xls, sheet_name=name)
            else:
                st.warning(f"Chybí list v Excelu: {name}")
                return pd.DataFrame()

        op1_key = load_sheet("seznam zkoušek PM+LM OP1")
        op2_key = load_sheet("seznam zkoušek PM+LM OP2")
        cely_key = load_sheet("seznam zkoušek Celý objekt")

        pm_op1 = load_sheet("PM - OP1")
        lm_op1 = load_sheet("LM - OP1")
        pm_op2 = load_sheet("PM - OP2")
        lm_op2 = load_sheet("LM - OP2")
        cely_objekt = load_sheet("Celý objekt")

        if not op1_key.empty and not pm_op1.empty:
            pm_op1 = process_op_sheet(op1_key, pm_op1, lab_text)
        if not op1_key.empty and not lm_op1.empty:
            lm_op1 = process_op_sheet(op1_key, lm_op1, lab_text)
        if not op2_key.empty and not pm_op2.empty:
            pm_op2 = process_op_sheet(op2_key, pm_op2, lab_text)
        if not op2_key.empty and not lm_op2.empty:
            lm_op2 = process_op_sheet(op2_key, lm_op2, lab_text)
        if not cely_key.empty and not cely_objekt.empty:
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

    except Exception as e:
        st.error(f"Chyba při zpracování souboru: {e}")
