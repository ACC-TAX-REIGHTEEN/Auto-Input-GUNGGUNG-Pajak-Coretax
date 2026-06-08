import pandas as pd
import os
import configparser

nama_file_input = 'ExAcc_clean_a_temp.xlsx'
nama_file_output = 'ExAcc_clean_temp.xlsx'
nama_file_config = 'Hapus.conf'

def get_config():
    config = configparser.ConfigParser()
    
    if not os.path.exists(nama_file_config):
        print(f"--> File '{nama_file_config}' tidak ditemukan. Membuat file default...")
        config['FILTER'] = {'mode': 'lengkap'}
        with open(nama_file_config, 'w') as configfile:
            config.write(configfile)
        print(f"--> File '{nama_file_config}' berhasil dibuat. Default mode: LENGKAP.")
        return "lengkap"
    
    try:
        config.read(nama_file_config)
        mode = config.get('FILTER', 'mode', fallback='lengkap')
        return mode.lower().strip()
    except Exception as e:
        print(f"-->  Gagal membaca config: {e}. Menggunakan mode default.")
        return "lengkap"

def proses_filter_data():
    mode_operasi = get_config()
    print(f"--> STATUS: Mode Operasi '{mode_operasi.upper()}' ---")

    if not os.path.exists(nama_file_input):
        print(f"--> File input '{nama_file_input}' tidak ditemukan.")
        return

    print(f"--> Membaca data dari: {nama_file_input} ...")
    try:
        df = pd.read_excel(nama_file_input, dtype=str)
        df = df.replace({'nan': '', 'NaN': '', float('nan'): ''})
            
    except Exception as e:
        print(f"--> Tidak bisa membuka file Excel: {e}")
        return

    kolom_target = "Tax Date"
    if kolom_target not in df.columns:
        print(f"--> Kolom '{kolom_target}' tidak ditemukan dalam data.")
        return

    if mode_operasi == 'hapus':
        dates = pd.to_datetime(df[kolom_target], errors='coerce')
        periode = dates.dt.to_period('M')
        
        if periode.mode().empty:
            print("--> Tidak ada data tanggal valid. Tidak ada yang dihapus.")
            df_hasil = df.copy()
        else:
            periode_mayoritas = periode.mode()[0]
            total_awal = len(df)
            mask_keep = (periode == periode_mayoritas)
            df_hasil = df[mask_keep].copy()
            total_akhir = len(df_hasil)
            dibuang = total_awal - total_akhir
            
            if dibuang > 0:
                print(f"--> Mayoritas data berada di bulan: {periode_mayoritas}")
                print(f"--> Menghapus {dibuang} baris data yang berbeda bulan (Nyasar).")
            else:
                print("--> Data sudah bersih (semua tanggal seragam).")
                
    else:
        print("--> Tidak ada penghapusan data (Mode Lengkap).")
        df_hasil = df.copy()

    try:
        with pd.ExcelWriter(nama_file_output, engine='openpyxl') as writer:
            df_hasil.to_excel(writer, index=False, sheet_name='Sheet1')
            print(f"--> SUKSES! File disimpan ke: '{nama_file_output}'")
    except Exception as e:
        print(f"--> Gagal menyimpan file: {e}")

if __name__ == "__main__":
    proses_filter_data()