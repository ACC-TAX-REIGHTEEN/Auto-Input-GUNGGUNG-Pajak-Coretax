import pandas as pd
import numpy as np
import sys
import os
import configparser

def get_config():
    config = configparser.ConfigParser()
    nama_conf = 'Helper.conf'
    
    if not os.path.exists(nama_conf):
        config['SETTINGS'] = {'mode': 'Match'}
        with open(nama_conf, 'w') as configfile:
            config.write(configfile)
        print(f"--> File {nama_conf} tidak ditemukan. Membuat default: Match")
        return "Match"
    
    try:
        config.read(nama_conf)
        return config.get('SETTINGS', 'mode', fallback='Match')
    except Exception as e:
        print(f"--> Gagal membaca config: {e}. Default ke Match.")
        return "Match"

def proses_pajak_bulat(nama_file):
    mode_helper = get_config()
    print(f"--> Mode Helper terdeteksi: '{mode_helper}'")

    print(f"--> Membaca file Excel: {nama_file}...")
    
    try:
        df_exacc = pd.read_excel(nama_file, sheet_name='ExAcc_clean_temp')
        df_nota = pd.read_excel(nama_file, sheet_name='Nota_terbaru_up_temp')
    except Exception as e:
        print(f"--> Error membaca file excel: {e}")
        return

    df_exacc['No Nota'] = df_exacc['No Nota'].astype(int)
    df_nota_clean = df_nota.dropna(subset=['No Nota']).copy()
    df_nota_clean['No Nota'] = df_nota_clean['No Nota'].astype(int)
    
    lookup_dict = dict(zip(df_nota_clean['No Nota'], df_nota_clean['Jumlah Pajak']))
    df_exacc['Matching'] = df_exacc['No Nota'].map(lookup_dict)

    def cek_match(row):
        val_match = row['Matching']
        val_asal = row['Jumlah Pajak']
        
        if pd.isna(val_match): 
            return "Tidak Ada Data"
        
        selisih = abs(val_match - val_asal)
        
        if mode_helper == "Losdol":
            return "Sesuai"
            
        elif mode_helper == "Toleransi":
            if selisih <= 1: 
                return "Sesuai"
            else: 
                return "Tidak Sesuai"
                
        else:
            if selisih == 0:
                return "Sesuai"
            else:
                return "Tidak Sesuai"

    df_exacc['Ket Matching'] = df_exacc.apply(cek_match, axis=1)

    if mode_helper == "Losdol":
        mask_valid = df_exacc['Matching'].notna()
        mask_diff = abs(df_exacc.loc[mask_valid, 'Matching'] - df_exacc.loc[mask_valid, 'Jumlah Pajak']) > 0
        
        df_beda_losdol = df_exacc.loc[mask_valid][mask_diff]
        
        if not df_beda_losdol.empty:
            print("\n" + "="*50)
            print("--> PERINGATAN MODE LOSDOL: Data berikut tidak sama persis")
            print("--> Namun proses tetap dilanjutkan (dianggap Sesuai).")
            print("="*50)
            cols_to_show = ['No Nota', 'Jumlah Pajak', 'Matching']
            print(df_beda_losdol[cols_to_show].to_string(index=False))
            print("="*50 + "\n")
        else:
            print("\n--> Mode Losdol: Semua data sama persis.\n")

    df_error = df_exacc[df_exacc['Ket Matching'] != "Sesuai"]

    if not df_error.empty:
        print(f"--> Ditemukan {len(df_error)} data tidak cocok (Mode: {mode_helper}). Proses dihentikan.")
        nama_output = 'Hasil_Matching_Tidak_Cocok_temp.xlsx'
        df_exacc.to_excel(nama_output, index=False)
        sys.exit(1) 
    else:
        print(f"--> Verifikasi Data Berhasil (Mode: {mode_helper}). Melakukan perhitungan...")

        df_exacc['Tax Base'] = np.ceil(df_exacc['Nilai Faktur'] / 1.11)

        df_exacc['OTB'] = np.ceil(df_exacc['Tax Base'] * (11 / 12))

        df_exacc['VAT'] = np.ceil(df_exacc['OTB'] * 0.12)

        df_exacc['Jml Pajak - VAT'] = df_exacc['Jumlah Pajak'] - df_exacc['VAT']

        nama_output = 'Pre-moving_temp.xlsx'
        
        with pd.ExcelWriter(nama_output, engine='xlsxwriter') as writer:
            df_exacc.to_excel(writer, sheet_name='ExAcc_clean_temp', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['ExAcc_clean_temp']
            
            fmt_desimal = workbook.add_format({'num_format': '#,##0.00'})
            
            try:
                idx_k = df_exacc.columns.get_loc('Tax Base')
                idx_l = df_exacc.columns.get_loc('OTB')
                idx_m = df_exacc.columns.get_loc('VAT')
                idx_n = df_exacc.columns.get_loc('Jml Pajak - VAT')
            except KeyError:
                idx_k, idx_l, idx_m, idx_n = -1, -1, -1, -1

            for i, col in enumerate(df_exacc.columns):
                try:
                    max_len = max(df_exacc[col].astype(str).map(len).max(), len(col)) + 2
                except:
                    max_len = 15 

                if i in [idx_k, idx_l, idx_m, idx_n]:
                    worksheet.set_column(i, i, max_len, fmt_desimal)
                else:
                    worksheet.set_column(i, i, max_len)

    print(f"--> Selesai. File disimpan sebagai: {nama_output}")

if __name__ == "__main__":
    proses_pajak_bulat('Merged_all_temp.xlsx')