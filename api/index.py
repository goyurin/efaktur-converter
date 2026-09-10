import io
import os
import warnings
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response, FileResponse
from fastapi.staticfiles import StaticFiles

# Suppress openpyxl Data Validation warnings
warnings.simplefilter(action='ignore', category=UserWarning)

import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

app = FastAPI()

def convert_efaktur_bytes_to_xml(file_bytes: bytes) -> bytes:
    """
    In-memory e-Faktur Excel to XML converter.
    Takes file bytes, processes 'Faktur' and 'DetailFaktur' sheets,
    and returns encoded pretty-printed XML bytes.
    """
    excel_buffer = io.BytesIO(file_bytes)
    
    # 1. Read global TIN/NPWP Penjual value from cell C1 (Row index 0, Column index 2)
    df_tin_raw = pd.read_excel(excel_buffer, sheet_name="Faktur", header=None, nrows=1)
    seller_tin = str(df_tin_raw.iloc[0, 2]).strip().split('.')[0]
    
    # Reset buffer position for reading remaining sheets
    excel_buffer.seek(0)
    
    # 2. Read 'Faktur' data block starting from row 3
    df_faktur = pd.read_excel(
        excel_buffer, 
        sheet_name="Faktur", 
        header=2,
        converters={
            'Baris': str,
            'ID TKU Penjual': str,
            'NPWP/NIK Pembeli': str,
            'ID TKU Pembeli': str
        }
    )
    
    excel_buffer.seek(0)
    
    # 3. Read 'DetailFaktur' sheet
    df_detail = pd.read_excel(
        excel_buffer, 
        sheet_name="DetailFaktur",
        converters={
            'Baris': str,
            'Kode Barang Jasa': str
        }
    )
    
    # Clean up whitespace and line reference strings
    df_faktur['Baris'] = df_faktur['Baris'].astype(str).str.strip()
    df_detail['Baris'] = df_detail['Baris'].astype(str).str.strip()
    
    # 4. Clean up Invoice Header data
    df_faktur_clean = df_faktur.dropna(subset=['Tanggal Faktur'])
    df_faktur_clean = df_faktur_clean[df_faktur_clean['Baris'].str.upper() != 'END']
    
    # 5. Build Root XML Node matching DJP specs
    root = ET.Element("TaxInvoiceBulk", {
        "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
    })
    ET.SubElement(root, "TIN").text = seller_tin
    
    list_of_tax_invoices = ET.SubElement(root, "ListOfTaxInvoice")
    
    # Helper to clean text and avoid "nan" outputs
    def clean_text(cell):
        if pd.isna(cell) or str(cell).strip().lower() == 'nan':
            return ""
        return str(cell).strip()

    # Numeric formatting rule (Keeps decimals for calculations)
    num_fmt = lambda val: f"{float(val):.2f}".rstrip('0').rstrip('.') if pd.notna(val) and str(val).strip().lower() != 'nan' else "0"

    # 6. Process Invoice Headers
    for _, f_row in df_faktur_clean.iterrows():
        current_baris_id = f_row['Baris']
        
        tax_invoice = ET.SubElement(list_of_tax_invoices, "TaxInvoice")
        
        # Clean Date component to standard YYYY-MM-DD
        raw_date = str(f_row['Tanggal Faktur']).split(" ")[0]
        
        # Format transaction codes (e.g., 4 becomes '04')
        trx_code = str(f_row['Kode Transaksi']).split('.')[0].zfill(2)

        # Map Invoice Header Tags
        ET.SubElement(tax_invoice, "TaxInvoiceDate").text = raw_date
        ET.SubElement(tax_invoice, "TaxInvoiceOpt").text = clean_text(f_row['Jenis Faktur'])
        ET.SubElement(tax_invoice, "TrxCode").text = trx_code
        ET.SubElement(tax_invoice, "AddInfo").text = clean_text(f_row['Keterangan Tambahan'])
        ET.SubElement(tax_invoice, "CustomDoc").text = clean_text(f_row['Dokumen Pendukung'])
        ET.SubElement(tax_invoice, "CustomDocMonthYear").text = clean_text(f_row['Period Dok Pendukung'])
        ET.SubElement(tax_invoice, "RefDesc").text = clean_text(f_row['Referensi'])
        ET.SubElement(tax_invoice, "FacilityStamp").text = clean_text(f_row['Cap Fasilitas'])
        
        ET.SubElement(tax_invoice, "SellerIDTKU").text = clean_text(f_row['ID TKU Penjual'])
        ET.SubElement(tax_invoice, "BuyerTin").text = clean_text(f_row['NPWP/NIK Pembeli'])
        
        ET.SubElement(tax_invoice, "BuyerDocument").text = clean_text(f_row['Jenis ID Pembeli'])
        ET.SubElement(tax_invoice, "BuyerCountry").text = clean_text(f_row['Negara Pembeli'])
        ET.SubElement(tax_invoice, "BuyerDocumentNumber").text = clean_text(f_row['Nomor Dokumen Pembeli'])
        ET.SubElement(tax_invoice, "BuyerName").text = clean_text(f_row['Nama Pembeli'])
        ET.SubElement(tax_invoice, "BuyerAdress").text = clean_text(f_row['Alamat Pembeli'])
        ET.SubElement(tax_invoice, "BuyerEmail").text = clean_text(f_row['Email Pembeli'])
        ET.SubElement(tax_invoice, "BuyerIDTKU").text = clean_text(f_row['ID TKU Pembeli'])
        
        # 7. Extract and Nest matching line items
        list_of_good_service = ET.SubElement(tax_invoice, "ListOfGoodService")
        matched_items = df_detail[df_detail['Baris'] == current_baris_id]
        
        for _, d_row in matched_items.iterrows():
            good_service = ET.SubElement(list_of_good_service, "GoodService")
            
            ET.SubElement(good_service, "Opt").text = clean_text(d_row['Barang/Jasa'])
            
            good_code = clean_text(d_row['Kode Barang Jasa'])
            ET.SubElement(good_service, "Code").text = good_code if good_code else "000000"
            
            ET.SubElement(good_service, "Name").text = clean_text(d_row['Nama Barang/Jasa'])
            ET.SubElement(good_service, "Unit").text = clean_text(d_row['Nama Satuan Ukur'])
            
            ET.SubElement(good_service, "Price").text = num_fmt(d_row['Harga Satuan'])
            ET.SubElement(good_service, "Qty").text = num_fmt(d_row['Jumlah Barang Jasa'])
            ET.SubElement(good_service, "TotalDiscount").text = num_fmt(d_row['Total Diskon'])
            ET.SubElement(good_service, "TaxBase").text = num_fmt(d_row['DPP'])
            ET.SubElement(good_service, "OtherTaxBase").text = num_fmt(d_row['DPP Nilai Lain'])
            ET.SubElement(good_service, "VATRate").text = str(int(float(d_row['Tarif PPN']))) if pd.notna(d_row['Tarif PPN']) and str(d_row['Tarif PPN']).strip().lower() != 'nan' else "0"
            ET.SubElement(good_service, "VAT").text = num_fmt(d_row['PPN'])
            ET.SubElement(good_service, "STLGRate").text = str(int(float(d_row['Tarif PPnBM']))) if pd.notna(d_row['Tarif PPnBM']) and str(d_row['Tarif PPnBM']).strip().lower() != 'nan' else "0"
            ET.SubElement(good_service, "STLG").text = num_fmt(d_row['PPnBM'])

    # 8. Render XML tree structure
    raw_xml_bytes = ET.tostring(root, encoding="utf-8")
    
    # 9. Format with minidom pretty-printing
    parsed_xml = minidom.parseString(raw_xml_bytes)
    return parsed_xml.toprettyxml(indent="  ", encoding="utf-8")


# --- ROUTING ENDPOINTS ---

@app.get("/")
async def serve_frontend():
    return FileResponse(os.path.join(os.path.dirname(__file__), "..", "public", "index.html"))


@app.post("/api/convert")
async def convert_excel_endpoint(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed.")

    try:
        file_bytes = await file.read()
        pretty_xml_bytes = convert_efaktur_bytes_to_xml(file_bytes)
        
        output_filename = file.filename.rsplit('.', 1)[0] + ".xml"

        return Response(
            content=pretty_xml_bytes,
            media_type="application/xml",
            headers={"Content-Disposition": f"attachment; filename={output_filename}"}
        )
    except KeyError as ke:
        raise HTTPException(
            status_code=400, 
            detail=f"Missing sheet or column in Excel file: {str(ke)}. Make sure sheets 'Faktur' and 'DetailFaktur' exist."
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


# Mount public directory for static assets if folder exists
if os.path.exists("public"):
    app.mount("/public", StaticFiles(directory="public"), name="public")