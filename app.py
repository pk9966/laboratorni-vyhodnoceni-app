import streamlit as st
st.set_page_config(page_title="Vyhodnocení laboratorního deníku")
st.write("Streamlit import OK")
import pandas as pd
st.write("Pandas import OK")
import pdfplumber
st.write("pdfplumber import OK")
import io
st.write("io import OK")
from openpyxl import load_workbook
st.write("openpyxl import OK")
from difflib import SequenceMatcher

st.title("Vyhodnocení laboratorního deníku")

pdf_file = st.file_uploader("Nahraj laboratorní deník (PDF)", type="pdf")
xlsx_file = st.file_uploader("Nahraj soubor Klíč.xlsx", type="xlsx")

def extract_text_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def contains_similar(text, keyword, threshold=0.6):
    text = text.lower()
    keyword = keyword.lower()
    if keyword in text:
        return True
    return similar(text, keyword) >= threshold

def count_matches_advanced(text, konstrukce, zkouska_raw, stanice_raw):
    st.markdown(f"---
🔍 **Konstrukce:** `{konstrukce}`")
    st.markdown(f"🔍 **Zkoušky:** `{zkouska_raw}`")
    st.markdown(f"🔍 **Staničení:** `{stanice_raw}`")
    druhy_zk = [z.strip().lower() for z in str(zkouska_raw).split(",") if z.strip()]
    staniceni = [s.strip().lower() for s in str(stanice_raw).split(",") if s.strip()]
    match_count = 0
    for line in text.splitlines():
        line_lower = line.lower()
        konstrukce_ok = contains_similar(line, konstrukce)
        zkouska_ok = any(z in line_lower for z in druhy_zk)
        stanice_ok = any(s in line_lower for s in staniceni)
        if konstrukce_ok and zkouska_ok and stanice_ok:
            match_count += 1
            st.markdown(f"✅ **Shoda nalezena:** `{line.strip()}`")
    st.markdown(f"**Celkem nalezeno:** `{match_count}` záznamů")
    return match_count

if pdf_file and xlsx_file:
    lab_text = extract_text_from_pdf(pdf_file)
        
        # Ukázka prvních 15 řádků z PDF
        st.subheader("📄 Náhled textu z PDF")
        st.text("
".join(lab_text.splitlines()[:15]))

    try:
        xlsx_bytes = xlsx_file.read()
        workbook = load_workbook(io.BytesIO(xlsx_bytes))

        def load_sheet_df(name):
            return pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=name)

        sheet_names = workbook.sheetnames

        def sheet_exists(name):
            return name in sheet_names

        op1_key = load_sheet_df("seznam zkoušek PM+LM OP1") if sheet_exists("seznam zkoušek PM+LM OP1") else pd.DataFrame()
        op2_key = load_sheet_df("seznam zkoušek PM+LM OP2") if sheet_exists("seznam zkoušek PM+LM OP2") else pd.DataFrame()
        cely_key = load_sheet_df("seznam zkoušek Celý objekt") if sheet_exists("seznam zkoušek Celý objekt") else pd.DataFrame()

        st.subheader("Výsledky hledání zkoušek v PDF")

        for key_df, label in [
            (op1_key, "OP1"),
            (op2_key, "OP2"),
            (cely_key, "Celý objekt")
        ]:
            if not key_df.empty:
                st.markdown(f"### 🔎 Zpracovávám list: {label}")
                for _, row in key_df.iterrows():
                    konstrukce = row.get("konstrukční prvek", "")
                    zkouska = row.get("druh zkoušky", "")
                    stanice = row.get("staničení", "")
                    if konstrukce and zkouska:
                        count = count_matches_advanced(lab_text, konstrukce, zkouska, stanice)
                        st.write(f"➡ Počet shod: {count}")

    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")
