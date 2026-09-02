import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import io

st.set_page_config(
    page_title="XLSX to XML Converter",
    layout="centered"
)

# ─────────────────────────────────────────────
# CONFIG TAB A (Depreciation & Amortization)
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
# CONFIG TAB B (Promotion Expense)
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
# CONFIG TAB D (Retail Invoice)
# ─────────────────────────────────────────────
HEADER_CONFIG_D = {
    "TIN": {
        "label": "NPWP",
        "xml_tag": "TIN"
    },
    "TaxPeriodMonth": {
        "label": "Masa Pajak",
        "xml_tag": "TaxPeriodMonth"
    },
    "TaxPeriodYear": {
        "label": "Tahun Pajak",
        "xml_tag": "TaxPeriodYear"
    }
}

COLUMN_MAPPING_D = {
    "TrxCode": "TrxCode",
    "BuyerName": "BuyerName",
    "BuyerIdOpt": "BuyerIdOpt",
    "BuyerIdNumber": "BuyerIdNumber",
    "GoodServiceOpt": "GoodServiceOpt",
    "SerialNo": "SerialNo",
    "TransactionDate": "TransactionDate",
    "TaxBaseSellingPrice": "TaxBaseSellingPrice",
    "OtherTaxBaseSellingPrice": "OtherTaxBaseSellingPrice",
    "VAT": "VAT",
    "STLG": "STLG",
    "Info": "Info"
}

# ─────────────────────────────────────────────
# HELPERS TAB A & B
# ─────────────────────────────────────────────
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

def convert_tab_d(df):
    header_values = extract_header_values(df, HEADER_CONFIG_D)

    root = ET.Element(
        "RetailInvoiceBulk",
        attrib={
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "xsi:noNamespaceSchemaLocation": "schema.xsd"
        }
    )

    ET.SubElement(root, "TIN").text = header_values.get("TIN", "")
    ET.SubElement(root, "TaxPeriodMonth").text = header_values.get("TaxPeriodMonth", "")
    ET.SubElement(root, "TaxPeriodYear").text = header_values.get("TaxPeriodYear", "")

    invoice_list = ET.SubElement(root, "ListOfRetailInvoice")

    header = None

    for _, row in df.iterrows():
        row_values = [str(cell).strip() for cell in row]

        # deteksi baris header: cari "TrxCode" di mana saja dalam baris
        if header is None and "TrxCode" in row_values:
            header = [str(col).replace(" ", "") for col in row]
            continue

        if header and not row.isnull().all():
            record = ET.SubElement(invoice_list, "RetailInvoice")
            for col_name, value in zip(header, row):
                if pd.notna(value) and col_name in COLUMN_MAPPING_D and str(value).strip() != "":
                    ET.SubElement(
                        record,
                        COLUMN_MAPPING_D[col_name]
                    ).text = str(value)

    ET.indent(root, space="  ")
    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, encoding="utf-8")
    return buf.getvalue().decode("utf-8")

# ─────────────────────────────────────────────
# HELPERS TAB C (Unifikasi / BpuBulk)
# ─────────────────────────────────────────────
def is_empty(value):
    """Cek apakah nilai kosong/nan."""
    if value is None:
        return True
    return str(value).strip() in ('', 'nan', 'NaN', 'None')

def read_excel_unifikasi(uploaded_file):
    """Baca Excel dengan skip 2 baris header (untuk Tab C)."""
    file_name = uploaded_file.name.lower()
    engine = 'xlrd' if file_name.endswith('.xls') else 'openpyxl'
    try:
        df = pd.read_excel(
            uploaded_file,
            sheet_name='DATA',
            engine=engine,
            skiprows=2,
            keep_default_na=False,
            # Kolom yang punya leading zero dibaca langsung sebagai string
            dtype={
                'ID TKU Pemotong': str,
                'ID TKU Penerima Penghasilan': str,
                'NPWP': str,
                'Nomor Dok. Referensi': str,
                'Nomor SP2D (IP)': str,
            }
        )
        df = df.dropna(how='all')
        # Fix tipe data object agar tidak OverflowError
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str)
        return df
    except Exception as e:
        st.error(f"Error membaca file Excel: {str(e)}")
        return None

def format_date(date_value):
    if pd.isna(date_value):
        return None
    if isinstance(date_value, str):
        try:
            return pd.to_datetime(date_value).strftime('%Y-%m-%d')
        except:
            return date_value
    elif isinstance(date_value, datetime):
        return date_value.strftime('%Y-%m-%d')
    else:
        try:
            return pd.to_datetime(date_value).strftime('%Y-%m-%d')
        except:
            return str(date_value)

def format_npwp(npwp_value):
    if is_empty(npwp_value):
        return "N/A"
    npwp_str = str(npwp_value).replace('.0', '').strip()
    if npwp_str.isdigit():
        return npwp_str.zfill(16)
    return npwp_str

def convert_tab_c(df, tin_pemotong):
    root = ET.Element("BpuBulk")
    root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    ET.SubElement(root, "TIN").text = tin_pemotong

    list_of_bpu = ET.SubElement(root, "ListOfBpu")

    for _, row in df.iterrows():
        npwp = row.get('NPWP', '')
        if is_empty(npwp):
            continue

        bpu = ET.SubElement(list_of_bpu, "Bpu")

        # TaxPeriodMonth
        masa = row.get('Masa Pajak', '')
        ET.SubElement(bpu, "TaxPeriodMonth").text = (
            str(int(float(masa))) if not is_empty(masa) else "1"
        )

        # TaxPeriodYear
        tahun = row.get('Tahun Pajak', '')
        ET.SubElement(bpu, "TaxPeriodYear").text = (
            str(int(float(tahun))) if not is_empty(tahun) else "2025"
        )

        # CounterpartTin
        ET.SubElement(bpu, "CounterpartTin").text = format_npwp(npwp)

        # IDPlaceOfBusinessActivityOfIncomeRecipient
        id_tku_penerima = str(row.get('ID TKU Penerima Penghasilan', '')).strip()
        ET.SubElement(bpu, "IDPlaceOfBusinessActivityOfIncomeRecipient").text = (
            "" if is_empty(id_tku_penerima) else id_tku_penerima
        )

        # TaxCertificate — jaga nilai N/A dari Excel
        fasilitas = str(row.get('Fasilitas', '')).strip()
        ET.SubElement(bpu, "TaxCertificate").text = (
            "N/A" if is_empty(fasilitas) else fasilitas
        )

        # TaxObjectCode
        kop = str(row.get('Kode Objek Pajak', '')).strip()
        ET.SubElement(bpu, "TaxObjectCode").text = (
            "" if is_empty(kop) else kop
        )

        # TaxBase
        dpp = row.get('DPP', '')
        ET.SubElement(bpu, "TaxBase").text = (
            str(int(float(dpp))) if not is_empty(dpp) else "0"
        )

        # Rate — biarkan float, jangan int agar 0.50 tidak jadi 0
        tarif = row.get('Tarif', '')
        ET.SubElement(bpu, "Rate").text = (
            str(tarif) if not is_empty(tarif) else "2"
        )

        # Document
        jenis_dok = str(row.get('Jenis Dok. Referensi', '')).strip()
        ET.SubElement(bpu, "Document").text = (
            "" if is_empty(jenis_dok) else jenis_dok
        )

        # DocumentNumber
        nomor_dok = str(row.get('Nomor Dok. Referensi', '')).strip()
        ET.SubElement(bpu, "DocumentNumber").text = (
            "" if is_empty(nomor_dok) else nomor_dok
        )

        # DocumentDate
        ET.SubElement(bpu, "DocumentDate").text = format_date(
            row.get('Tanggal Dok. Referensi')
        ) or ""

        # IDPlaceOfBusinessActivity — leading zero dijaga karena dtype=str saat baca
        id_tku = str(row.get('ID TKU Pemotong', '')).strip()
        if not is_empty(id_tku):
            id_tku_final = "00" + id_tku if len(id_tku) == 20 else id_tku
        else:
            id_tku_final = ""
        ET.SubElement(bpu, "IDPlaceOfBusinessActivity").text = id_tku_final

        # GovTreasurerOpt
        opsi = str(row.get('Opsi Pembayaran (IP)', '')).strip()
        ET.SubElement(bpu, "GovTreasurerOpt").text = (
            "N/A" if is_empty(opsi) else opsi
        )

        # SP2DNumber
        sp2d = ET.SubElement(bpu, "SP2DNumber")
        nomor_sp2d = str(row.get('Nomor SP2D (IP)', '')).strip()
        if is_empty(nomor_sp2d):
            sp2d.set("xsi:nil", "true")
        else:
            sp2d.text = nomor_sp2d

        # WithholdingDate
        ET.SubElement(bpu, "WithholdingDate").text = format_date(
            row.get('Tanggal Pemotongan')
        ) or ""

    return root

def prettify_xml(element):
    rough_string = ET.tostring(element, 'unicode')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="\t")[23:].strip()

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    st.title("🔄 Konverter XLSX ke XML")

    tab_a, tab_b, tab_c, tab_d = st.tabs([
        "📄 Depreciation & Amortization (L9)",
        "🎁 Promotion Expense (L11)",
        "📑 Unifikasi (BpuBulk)",
        "🛒 Retail"
    ])

    # ── TAB A ──────────────────────────────────
    with tab_a:
        st.markdown("Konversi XLSX ke XML **Depreciation & Amortization**")

        file_a = st.file_uploader(
            "Upload XLSX dengan SheetName = DATA",
            type=["xlsx", "xls"],
            key="file_a"
        )

        if file_a:
            df_a = pd.read_excel(file_a, sheet_name="DATA", header=None, keep_default_na=False)
            df_a = df_a.dropna(how="all")
            for col in df_a.columns:
                if df_a[col].dtype == 'object':
                    df_a[col] = df_a[col].astype(str)

            st.dataframe(df_a.head(), use_container_width=True)

            if st.button("🔄 Convert to XML", type="primary", key="btn_a", use_container_width=True):
                try:
                    xml = convert_tab_a(df_a)
                    st.code(xml[:800] + "...", language="xml")
                    st.download_button(
                        "💾 Download XML",
                        xml,
                        file_a.name.replace(".xlsx", ".xml").replace(".xls", ".xml"),
                        "application/xml",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Error saat konversi: {str(e)}")

    # ── TAB B ──────────────────────────────────
    with tab_b:
        st.markdown("Konversi XLSX ke XML **Promotion Expense**")

        file_b = st.file_uploader(
            "Upload XLSX dengan SheetName = DATA",
            type=["xlsx", "xls"],
            key="file_b"
        )

        if file_b:
            df_b = pd.read_excel(file_b, sheet_name="DATA", header=None, keep_default_na=False)
            df_b = df_b.dropna(how="all")
            for col in df_b.columns:
                if df_b[col].dtype == 'object':
                    df_b[col] = df_b[col].astype(str)

            st.dataframe(df_b.head(), use_container_width=True)

            if st.button("🔄 Convert to XML", type="primary", key="btn_b", use_container_width=True):
                try:
                    xml = convert_tab_b(df_b)
                    st.code(xml[:800] + "...", language="xml")
                    st.download_button(
                        "💾 Download XML",
                        xml,
                        file_b.name.replace(".xlsx", ".xml").replace(".xls", ".xml"),
                        "application/xml",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Error saat konversi: {str(e)}")

    # ── TAB C ──────────────────────────────────
    with tab_c:
        st.markdown("Konversi XLSX ke XML **Unifikasi (BpuBulk)**")

        tin_pemotong = st.text_input(
            "TIN Pemotong:",
            value="0013936760054000",
            help="Masukkan TIN Pemotong (16 digit)",
            key="tin_c"
        )

        file_c = st.file_uploader(
            "Upload XLSX dengan SheetName = DATA",
            type=["xlsx", "xls"],
            key="file_c"
        )

        if file_c:
            df_c = read_excel_unifikasi(file_c)

            if df_c is not None:
                st.subheader("📊 Preview Data")
                st.dataframe(df_c.head(10), use_container_width=True)
                st.info(f"📈 Total baris data: {len(df_c)}")

                if st.button("🔄 Convert to XML", type="primary", key="btn_c", use_container_width=True):
                    with st.spinner("⚙️ Mengkonversi ke XML..."):
                        try:
                            xml_root = convert_tab_c(df_c, tin_pemotong)
                            xml_string = prettify_xml(xml_root)
                            final_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_string

                            st.success("✅ Konversi berhasil!")
                            st.subheader("📄 Preview XML")
                            st.code(
                                final_xml[:2000] + "..." if len(final_xml) > 2000 else final_xml,
                                language="xml"
                            )

                            filename = file_c.name.replace('.xlsx', '.xml').replace('.xls', '.xml')
                            st.download_button(
                                label="💾 Download XML",
                                data=final_xml,
                                file_name=filename,
                                mime="application/xml",
                                type="primary",
                                use_container_width=True
                            )
                        except Exception as e:
                            st.error(f"❌ Error saat konversi: {str(e)}")

            # ── TAB D ──────────────────────────────────
    with tab_d:
        st.markdown("Konversi XLSX ke XML **Retail Invoice**")

        file_d = st.file_uploader(
            "Upload XLSX dengan SheetName = DATA",
            type=["xlsx", "xls"],
            key="file_d"
        )

        if file_d:
            df_d = pd.read_excel(file_d, sheet_name="DATA", header=None, keep_default_na=False)
            df_d = df_d.dropna(how="all")
            for col in df_d.columns:
                if df_d[col].dtype == 'object':
                    df_d[col] = df_d[col].astype(str)

            st.dataframe(df_d.head(), use_container_width=True)

            if st.button("🔄 Convert to XML", type="primary", key="btn_d", use_container_width=True):
                try:
                    xml = convert_tab_d(df_d)
                    st.code(xml[:800] + "...", language="xml")
                    st.download_button(
                        "💾 Download XML",
                        xml,
                        file_d.name.replace(".xlsx", ".xml").replace(".xls", ".xml"),
                        "application/xml",
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
            - **Kolom-kolom berikut** (mulai dari baris ke-3):
              - Masa Pajak, Tahun Pajak, NPWP
              - ID TKU Penerima Penghasilan, Fasilitas
              - Kode Objek Pajak, DPP, Tarif
              - Jenis Dok. Referensi, Nomor Dok. Referensi, Tanggal Dok. Referensi
              - ID TKU Pemotong, Opsi Pembayaran (IP)
              - Nomor SP2D (IP), Tanggal Pemotongan
            """)

if __name__ == "__main__":
    main()