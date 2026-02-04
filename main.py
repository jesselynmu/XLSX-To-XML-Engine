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

COLUMN_MAPPING = {
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

def read_excel_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name="DATA", header=None)
    df = df.dropna(how="all")
    return df

def convert_to_xml(df):
    root = ET.Element(
        "DepreciationAmortization",
        attrib={
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
        }
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
                if pd.notna(value) and col_name in COLUMN_MAPPING:
                    ET.SubElement(
                        record,
                        COLUMN_MAPPING[col_name]
                    ).text = str(value)

    ET.indent(root, space="  ")
    buffer = io.BytesIO()
    ET.ElementTree(root).write(buffer, encoding="utf-8")
    return buffer.getvalue().decode("utf-8")

def main():
    st.title("🔄 Konverter XLSX ke XML")
    st.markdown(
        "Aplikasi untuk mengkonversi file XLSX (sheet **DATA**) "
        "menjadi format XML **Depreciation & Amortization**"
    )

    uploaded_file = st.file_uploader(
        "📁 Pilih file XLSX:",
        type=["xlsx", "xls"],
        help="Upload file Excel dengan sheet bernama 'DATA'"
    )

    if uploaded_file is not None:
        st.success(f"✅ File '{uploaded_file.name}' berhasil diupload!")

        with st.spinner("📖 Membaca file Excel..."):
            df = read_excel_file(uploaded_file)

        st.subheader("📊 Preview Data")
        st.dataframe(df.head(), use_container_width=True)
        st.info(f"📈 Total baris data: {len(df)}")

        if st.button(
            "🔄 Konversi ke XML",
            type="primary",
            use_container_width=True
        ):
            with st.spinner("⚙️ Mengkonversi ke XML..."):
                try:
                    xml_string = convert_to_xml(df)

                    st.success("✅ Konversi berhasil!")

                    st.subheader("📄 Preview XML")
                    st.code(
                        xml_string[:500] + "..."
                        if len(xml_string) > 500
                        else xml_string,
                        language="xml"
                    )

                    filename = (
                        uploaded_file.name
                        .replace(".xlsx", ".xml")
                        .replace(".xls", ".xml")
                    )

                    st.download_button(
                        label="💾 Download XML",
                        data=xml_string,
                        file_name=filename,
                        mime="application/xml",
                        type="primary",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"❌ Error saat konversi: {str(e)}")

    else:
        st.info("👆 Silakan upload file XLSX di atas untuk memulai konversi")

        st.subheader("📋 Format File yang Diharapkan")
        st.markdown("""
        File XLSX harus memiliki:
        - **Sheet bernama 'DATA'**
        - **Section di kolom pertama**:
          - Daftar Penyusutan
          - Daftar Amortisasi
        - **Header kolom**:
          - Kode Aset
          - Kelompok Aset
          - Bulan Perolehan
          - Tahun Perolehan
          - Harga Perolehan
          - Nilai Sisa Buku
          - Metode Komersial
          - Metode Fiskal
          - Penyusutan Fiskal
          - Keterangan
        """)

if __name__ == "__main__":
    main()
