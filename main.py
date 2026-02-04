import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io

st.set_page_config(
    page_title="XLSX to XML Converter",
    layout="centered"
)

SECTION_CONFIG = {
    "Daftar Penyusutan": {
        "outer_tag": "ListOfDepreciation",
        "list_tag": "Depreciation"
    },
    "Daftar Amortisasi": {
        "outer_tag": "ListOfAmortization",
        "list_tag": "Amortization"
    }
}

COLUMN_MAPPING_A = {
    "KodeAset": "CodeOfAsset",
    "KelompokAset": "GroupOfAsset",
    "BulanPerolehan": "MonthOfAcquisition",
    "TahunPerolehan": "YearOfAcquisition",
    "HargaPerolehan": "AcquisitionPrice",
    "NilaiSisaBuku": "RemainingValue",
    "MetodeKomersial": "CommercialMethode",
    "MetodeFiskal": "FiscalMethode",
    "PenyusutanFiskal": "FiscalDepretiationThisYear",
    "Keterangan": "Notes"
}

HEADER_CONFIG_B = {
    "TIN": {
        "label": "NPWP SPT",
        "xml_tag": "TIN"
    },
    "TaxYear": {
        "label": "Tahun Pajak",
        "xml_tag": "TaxYear"
    }
}

COLUMN_MAPPING_B = {
    "NomorIdentitas": "IdentityNumber",
    "NamaPenerima": "Name",
    "Alamat": "Address",
    "Tanggal": "DateOfPromotion",
    "BentukJenisBiaya": "FormAndType",
    "Nilai": "AmountOfPromotion",
    "PPhDipotongDipungut": "AmountOfWitholding",
    "NomorBupot": "WitholdingSlipNumber",
    "Keterangan": "Description"
}

def read_excel_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name="DATA", header=None)
    return df.dropna(how="all")

def extract_header_values(df, header_config):
    result = {}
    for _, row in df.iterrows():
        for cfg in header_config.values():
            for i, cell in enumerate(row):
                if str(cell).strip() == cfg["label"]:
                    result[cfg["xml_tag"]] = (
                        str(row[i + 1]).strip()
                        if i + 1 < len(row)
                        else ""
                    )
    return result

def convert_tab_a(df):
    root = ET.Element(
        "DepreciationAmortization",
        attrib={"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"}
    )

    current_cfg = None
    outer_el = None
    header = None

    for _, row in df.iterrows():
        first_cell = str(row[0]).strip()

        if first_cell in SECTION_CONFIG:
            current_cfg = SECTION_CONFIG[first_cell]
            outer_el = ET.SubElement(root, current_cfg["outer_tag"])
            header = None
            continue

        if current_cfg and header is None and first_cell == "Kode Aset":
            header = [str(col).replace(" ", "") for col in row]
            continue

        if header and not row.isnull().all():
            record = ET.SubElement(outer_el, current_cfg["list_tag"])
            for col_name, value in zip(header, row):
                if pd.notna(value) and col_name in COLUMN_MAPPING_A:
                    ET.SubElement(
                        record,
                        COLUMN_MAPPING_A[col_name]
                    ).text = str(value)

    ET.indent(root, space="  ")
    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, encoding="utf-8")
    return buf.getvalue().decode("utf-8")

def convert_tab_b(df):
    header_values = extract_header_values(df, HEADER_CONFIG_B)

    root = ET.Element(
        "PromotionExpense",
        attrib={"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"}
    )

    ET.SubElement(root, "TIN").text = header_values.get("TIN", "")
    ET.SubElement(root, "TaxYear").text = header_values.get("TaxYear", "")

    expense_list = ET.SubElement(root, "PromotionExpenseList")

    header = None
    start_data = False

    for _, row in df.iterrows():
        first_cell = str(row[0]).strip()

        if first_cell == "Nomor Identitas":
            header = [
                str(col)
                .replace(" ", "")
                .replace("&", "")
                .replace("/", "")
                for col in row
            ]
            start_data = True
            continue

        if start_data and header and not row.isnull().all():
            item = ET.SubElement(expense_list, "List")
            for col_name, value in zip(header, row):
                if pd.notna(value) and col_name in COLUMN_MAPPING_B:
                    ET.SubElement(
                        item,
                        COLUMN_MAPPING_B[col_name]
                    ).text = str(value)

    ET.indent(root, space="  ")
    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, encoding="utf-8")
    return buf.getvalue().decode("utf-8")

def main():
    st.title("🔄 Konverter XLSX ke XML")

    tab_a, tab_b = st.tabs([
        "📄 Depreciation & Amortization (L9)",
        "🎁 Promotion Expense (L11)"
    ])

    with tab_a:
        st.markdown("Konversi XLSX ke XML **Depreciation & Amortization**")

        file_a = st.file_uploader(
            "Upload XLSX dengan SheetName = DATA",
            type=["xlsx", "xls"],
            key="file_a"
        )

        if file_a:
            df = read_excel_file(file_a)
            st.dataframe(df.head(), width="stretch")

            if st.button("🔄 Convert to XML", type="primary", key="btn_a", width="stretch"):
                xml = convert_tab_a(df)
                st.code(xml[:800] + "...", language="xml")

                st.download_button(
                    "💾 Download XML",
                    xml,
                    file_a.name.replace(".xlsx", ".xml"),
                    "application/xml",
                    width="stretch"
                )

    with tab_b:
        st.markdown("Konversi XLSX ke XML **Promotion Expense**")

        file_b = st.file_uploader(
            "Upload XLSX dengan SheetName = DATA",
            type=["xlsx", "xls"],
            key="file_b"
        )

        if file_b:
            df = read_excel_file(file_b)
            st.dataframe(df.head(), width="stretch")

            if st.button("🔄 Convert to XML", type="primary", key="btn_b", width="stretch"):
                xml = convert_tab_b(df)
                st.code(xml[:800] + "...", language="xml")

                st.download_button(
                    "💾 Download XML",
                    xml,
                    file_b.name.replace(".xlsx", ".xml"),
                    "application/xml",
                    width="stretch"
                )

if __name__ == "__main__":
    main()
