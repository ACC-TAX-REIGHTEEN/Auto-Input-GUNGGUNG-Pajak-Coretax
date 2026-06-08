import os
import shutil
import sys
import subprocess
import pandas as pd
import xlwings as xw
from datetime import datetime
import glob
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DAPUR_DIR = os.path.join(BASE_DIR, "Dapur")
EXACC_SOURCE = os.path.join(BASE_DIR, "ExAcc.xls")
EXACC_DEST = os.path.join(DAPUR_DIR, "ExAcc.xls")

REQUIRED_FILES_IN_DAPUR = [
    "__init__.py",
    "1_FetchFileSS.py",
    "2_Acc_Cleaner.py",
    "2_Acc_CleanerDel.py",
    "3_MovDt2MtchFrFetch.py",
    "4_Merged_all.py",
    "5_Processing.py",
    "6_Premov2template.py",
    "Hapus.conf",
    "Helper.conf",
    "TEMPLATE-GUNG.xlsx"
]

def print_step(message):
    print(f"\n{'='*50}\n--> {message}\n{'='*50}")

def exit_with_pause(message=None):
    if message:
        print(message)
    print("\n" + "!"*50)
    input("--> Terjadi kesalahan. Tekan Enter untuk keluar...")
    sys.exit(1)

def initial_cleanup():
    print_step("--> Pembersihan Awal Folder Dapur")
    
    exacc_dapur = os.path.join(DAPUR_DIR, "ExAcc.xls")
    if os.path.exists(exacc_dapur):
        try:
            os.remove(exacc_dapur)
            print(f"--> Berhasil menghapus file lama: {exacc_dapur}")
        except Exception as e:
            print(f"--> Gagal menghapus {exacc_dapur}. Error: {e}")

    temp_files = glob.glob(os.path.join(DAPUR_DIR, "*temp.xlsx"))
    if temp_files:
        for f in temp_files:
            try:
                os.remove(f)
                filename = os.path.basename(f)
                print(f"--> Berhasil menghapus temp file: {filename}")
            except Exception as e:
                print(f"--> Gagal menghapus {f}: {e}")
    else:
        print("--> Tidak ditemukan file temp sisa.")

    print("--> Folder Dapur bersih dan siap digunakan.")

def check_requirements():
    print_step("--> Memeriksa kelengkapan File dan Folder")
    
    if not os.path.exists(EXACC_SOURCE):
        exit_with_pause(f"--> File tidak ditemukan: {EXACC_SOURCE}")
    
    if not os.path.exists(DAPUR_DIR):
        exit_with_pause(f"--> Folder tidak ditemukan: {DAPUR_DIR}")

    missing_files = []
    for filename in REQUIRED_FILES_IN_DAPUR:
        file_path = os.path.join(DAPUR_DIR, filename)
        if not os.path.exists(file_path):
            missing_files.append(filename)
    
    if missing_files:
        print(f"--> File berikut tidak ditemukan di dalam folder 'Dapur':")
        for f in missing_files:
            print(f" - {f}")
        exit_with_pause("--> Proses digagalkan karena file tidak lengkap.")
    
    print("--> Semua file dan folder lengkap.")

def run_scripts():
    print_step("--> Menyalin ExAcc.xls ke folder Dapur")
    try:
        shutil.copy(EXACC_SOURCE, EXACC_DEST)
        print(f"--> {EXACC_SOURCE} disalin ke {EXACC_DEST}")
    except Exception as e:
        exit_with_pause(f"--> Gagal menyalin file: {e}")

    scripts = [
        "1_FetchFileSS.py",
        "2_Acc_Cleaner.py",
        "2_Acc_CleanerDel.py",
        "3_MovDt2MtchFrFetch.py",
        "4_Merged_all.py",
        "5_Processing.py",
        "6_Premov2template.py"
    ]

    print_step("--> Mengeksekusi Script Python di folder Dapur")
    
    original_cwd = os.getcwd()
    os.chdir(DAPUR_DIR)

    try:
        for script in scripts:
            print(f"\n--> Menjalankan: {script}...")
            
            if script == "1_FetchFileSS.py":
                print("\n*** PERHATIAN: Silakan cek layar untuk memilih Sheet ***")
            
            subprocess.run([sys.executable, script], check=True)
            print(f"-->  {script} selesai.")
            
    except subprocess.CalledProcessError as e:
        os.chdir(original_cwd) 
        exit_with_pause(f"\n--> Terjadi kesalahan saat menjalankan {script}.\nProses dihentikan.")
    except Exception as e:
        os.chdir(original_cwd)
        exit_with_pause(f"\n--> Error tidak terduga di run_scripts: {e}")
    
    os.chdir(original_cwd) 

def process_excel_template():
    print_step("--> Memproses Excel dengan Xlwings")
    
    source_data_path = os.path.join(DAPUR_DIR, "Pre_template_temp.xlsx")
    template_path = os.path.join(DAPUR_DIR, "TEMPLATE-GUNG.xlsx")
    new_template_path = os.path.join(DAPUR_DIR, "TEMPLATE-GUNG-NEW.xlsx")

    if not os.path.exists(source_data_path):
        exit_with_pause(f"[ERROR] File {source_data_path} tidak ditemukan.")

    try:
        shutil.copy(template_path, new_template_path)
        print(f"--> Template disalin ke {new_template_path}")
    except Exception as e:
        exit_with_pause(f"--> Gagal menyalin template: {e}")

    try:
        df = pd.read_excel(source_data_path)
        
        if "TransactionDate" not in df.columns:
            exit_with_pause("--> Kolom 'TransactionDate' tidak ditemukan.")
        
        df['TransactionDate'] = pd.to_datetime(df['TransactionDate'])
        
        if not df.empty:
            mode_month = df['TransactionDate'].dt.month.mode()[0]
            mode_year = df['TransactionDate'].dt.year.mode()[0]
            date_for_filename = datetime(year=mode_year, month=mode_month, day=1) 
        else:
            print("--> Data kosong. Menggunakan tanggal hari ini untuk nama file.")
            date_for_filename = datetime.now()
            mode_month = date_for_filename.month
            mode_year = date_for_filename.year
        
    except Exception as e:
        exit_with_pause(f"--> Gagal membaca data source: {e}")

    app = xw.App(visible=False)
    try:
        wb = app.books.open(new_template_path)
        sheet = wb.sheets["DATA"]

        num_rows_data = len(df)
        rows_to_insert = num_rows_data - 1 

        print(f"--> Menambahkan {rows_to_insert} baris baru.")

        if rows_to_insert > 0:
            sheet.range(f"6:{6+rows_to_insert-1}").insert(shift='down')

        print("--> Menempel data mulai dari cell B5...")
        sheet.range("B5").options(index=False, header=False).value = df

        sheet.range("B2").value = mode_month
        sheet.range("B3").value = mode_year

        wb.save()
        wb.close()
        
        return date_for_filename

    except Exception as e:
        if 'app' in locals(): app.quit() 
        exit_with_pause(f"--> Xlwings Error: {e}")
    finally:
        app.quit()

def finalize_and_cleanup(date_obj):
    print_step("--> Finalisasi dan Pembersihan Akhir")
    
    filename_str = f"GUNGGUNG-{date_obj.strftime('%b').upper()}-{date_obj.strftime('%y')}.xlsx"
    
    source_path = os.path.join(DAPUR_DIR, "TEMPLATE-GUNG-NEW.xlsx")
    dest_path = os.path.join(BASE_DIR, filename_str)
    
    try:
        if os.path.exists(dest_path):
            os.remove(dest_path)
            print(f"--> File lama {filename_str} ditimpa.")
            
        shutil.move(source_path, dest_path)
        print(f"--> File hasil disimpan sebagai: {filename_str}")
    except Exception as e:
        print(f"--> Gagal memindahkan file hasil: {e}")

    print("--> Membersihkan file sementara...")
    temp_files = glob.glob(os.path.join(DAPUR_DIR, "*temp.xlsx"))
    for f in temp_files:
        try:
            os.remove(f)
        except:
            pass
            
    exacc_dapur = os.path.join(DAPUR_DIR, "ExAcc.xls")
    if os.path.exists(exacc_dapur):
        try:
            os.remove(exacc_dapur)
            print("--> File ExAcc.xls berhasil dihapus dari folder Dapur.")
        except Exception as e:
            pass

    print("\n" + "="*50)
    print("--> SEMUA PROSES SELESAI DENGAN SUKSES!")
    print("="*50)
    input("\n--> Tekan Enter untuk keluar. Pesan moral: Jangan pernah percaya kepada siapapun, bahkan jika itu dirimu sendiri. Selalu lakukan cek ulang, karena mesin juga bisa salah.")

def main():
    try:
        initial_cleanup()
        check_requirements()
        run_scripts()
        date_result = process_excel_template() 
        finalize_and_cleanup(date_result)
    except Exception as e:
        print("\n" + "!"*50)
        print(f"--> Terjadi kesalahan tak terduga:\n{e}")
        print("-" * 20)
        traceback.print_exc()
        exit_with_pause("--> Program berhenti mendadak.")

if __name__ == "__main__":
    main()
