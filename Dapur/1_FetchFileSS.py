import pandas as pd

sheet_id = "SPREADSHEETS ID DI SINI"
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

print("--> Sedang menghubungi Google Sheets dan membaca data...")

all_sheets = pd.read_excel(url, sheet_name=None)
daftar_nama_sheet = list(all_sheets.keys())

print(f"\n--> Ditemukan {len(daftar_nama_sheet)} sheet dalam file ini:")
for i, nama in enumerate(daftar_nama_sheet):
    print(f"{i + 1}. {nama}")

try:
    pilihan = int(input("\n--> Masukkan NOMOR sheet yang ingin diunduh: "))
    
    if 1 <= pilihan <= len(daftar_nama_sheet):
        nama_sheet_terpilih = daftar_nama_sheet[pilihan - 1]
        df_terpilih = all_sheets[nama_sheet_terpilih]
        
        print(f"\n--> Memproses sheet: '{nama_sheet_terpilih}'...")

        output_filename = f"--> Nota_terbaru_temp.xlsx" 
        
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df_terpilih.to_excel(writer, sheet_name=nama_sheet_terpilih, index=False)
            worksheet = writer.sheets[nama_sheet_terpilih]
            
            for i, col in enumerate(df_terpilih.columns): 
                column_len = df_terpilih[col].astype(str).map(len).max()
                header_len = len(str(col))
                max_len = max(column_len, header_len) + 2
                worksheet.set_column(i, i, max_len)

        print(f"--> Selesai! Data berhasil disimpan di file: '{output_filename}'")
        
    else:
        print(f"--> Error: Nomor harus antara 1 sampai {len(daftar_nama_sheet)}.")

except ValueError:
    print("--> Error: Harap masukkan angka yang valid.")