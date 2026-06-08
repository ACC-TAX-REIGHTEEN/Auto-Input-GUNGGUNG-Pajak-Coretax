import pandas as pd
import os

files_to_merge = ["ExAcc_clean_temp.xlsx", "Nota_terbaru_up_temp.xlsx"]
output_filename = "Merged_all_temp.xlsx"

print(f"--> Sedang memproses penggabungan ke '{output_filename}'...")

with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
    
    for file_name in files_to_merge:
        try:
            df = pd.read_excel(file_name, sheet_name=0)
            sheet_name_baru = os.path.splitext(file_name)[0]
            df.to_excel(writer, sheet_name=sheet_name_baru, index=False)
            worksheet = writer.sheets[sheet_name_baru]
            
            for i, col in enumerate(df.columns):
                max_len_data = df[col].astype(str).map(len).max()
                len_header = len(str(col))
                final_len = max(max_len_data, len_header) + 2
                worksheet.set_column(i, i, final_len)
                
            print(f"--> '{file_name}' --> Sheet: '{sheet_name_baru}' (Auto-fit selesai)")
            
        except FileNotFoundError:
            print(f"--> File '{file_name}' tidak ditemukan. Lewati.")
        except Exception as e:
            print(f"--> Gagal memproses '{file_name}': {e}")

print("\n--> Selesai! File berhasil digabungkan dan dirapikan.")