import pandas as pd
import os
from openpyxl.utils import get_column_letter

nama_file_input = 'ExAcc.xls'
nama_file_output = 'ExAcc_clean_a_temp.xlsx'

ciri_bulan = ['Jan', 'Peb', 'Feb', 'Mar', 'Apr', 'Mei', 'May', 'Jun', 'Jul', 'Agu', 'Ags', 'Aug', 'Sep', 'Okt', 'Oct', 'Nov', 'Nop', 'Des', 'Dec']

def perbaiki_tanggal(date_str):
    if pd.isna(date_str): return None
    date_str = str(date_str).strip()
    ganti_bulan = {
        'Peb': 'Feb', 'Mei': 'May', 'Agu': 'Aug', 'Ags': 'Aug',
        'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec'
    }
    for indo, eng in ganti_bulan.items():
        if indo in date_str:
            date_str = date_str.replace(indo, eng)
    return date_str

def deteksi_kolom(df_sample):
    mapping = {}
    
    idx_tanggal = -1
    
    for col_idx in df_sample.columns:
        sample_data = df_sample[col_idx].astype(str).values
        match_count = 0
        for val in sample_data[:20]:
            if any(bulan in val for bulan in ciri_bulan):
                match_count += 1
        
        if match_count >= 3:
            idx_tanggal = col_idx
            break
    
    if idx_tanggal == -1:
        raise ValueError("--> Kolom Tanggal tidak ditemukan! Cek apakah file memiliki nama bulan (Jan, Peb, dll).")

    print(f"--> Kolom Tanggal ditemukan di index ke-{idx_tanggal}")
    mapping['Tanggal'] = idx_tanggal
    mapping['No Nota'] = idx_tanggal - 2 
    mapping['Nama Pelanggan'] = idx_tanggal + 1
    start_angka = idx_tanggal + 2 
    mapping['Nilai Faktur'] = start_angka
    mapping['DISC INV'] = start_angka + 1
    mapping['SUB TOTAL'] = start_angka + 2
    mapping['Jumlah Pajak'] = start_angka + 3
    mapping['Tax Date'] = start_angka + 4
    
    return mapping

if os.path.exists(nama_file_input):
    try:
        try:
            df = pd.read_excel(nama_file_input, header=None)
        except ValueError:
            df = pd.read_csv(nama_file_input, header=None, encoding='latin1')
            
        df = df.dropna(how='all')
        
        peta_kolom = deteksi_kolom(df)
        
        target_cols = [
            peta_kolom['No Nota'],
            peta_kolom['Tanggal'],
            peta_kolom['Nama Pelanggan'],
            peta_kolom['Nilai Faktur'],
            peta_kolom['DISC INV'],
            peta_kolom['SUB TOTAL'],
            peta_kolom['Jumlah Pajak'],
            peta_kolom['Tax Date']
        ]
        
        df_bersih = df.iloc[:, target_cols].copy()
        df_bersih.columns = [
            "No Nota", "Tanggal", "Nama Pelanggan", 
            "Nilai Faktur", "DISC INV", "SUB TOTAL", 
            "Jumlah Pajak", "Tax Date"
        ]

        df_bersih = df_bersih.dropna(subset=["No Nota"])

        print("--> Sedang memproses data...")
        
        for col in ["Tanggal", "Tax Date"]:
            df_bersih[col] = df_bersih[col].apply(perbaiki_tanggal)
            df_bersih[col] = pd.to_datetime(df_bersih[col], dayfirst=True, errors='coerce')
            df_bersih[col] = df_bersih[col].dt.strftime('%Y-%m-%d')

        kolom_angka = ["Nilai Faktur", "SUB TOTAL", "Jumlah Pajak"]
        for col in kolom_angka:
            df_bersih[col] = df_bersih[col].astype(str).str.replace(r'\.0$', '', regex=True)
            df_bersih[col] = df_bersih[col].str.replace('.', '', regex=False)
            df_bersih[col] = pd.to_numeric(df_bersih[col], errors='coerce')

        df_bersih["No Nota"] = pd.to_numeric(df_bersih["No Nota"], errors='coerce').fillna(0).astype(int).astype(str)
        df_bersih = df_bersih[df_bersih["No Nota"] != '0']
        writer = pd.ExcelWriter(nama_file_output, engine='openpyxl')
        df_bersih.to_excel(writer, index=False, sheet_name='Sheet1')
        
        worksheet = writer.sheets['Sheet1']
        for i, column in enumerate(worksheet.columns):
            max_length = 0
            column_letter = get_column_letter(i + 1)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            worksheet.column_dimensions[column_letter].width = max_length + 2

        writer.close()
        print(f"--> Sukses! File '{nama_file_output}' berhasil dibuat dengan deteksi kolom otomatis.")

    except Exception as e:
        print(f"--> Terjadi Error: {e}")