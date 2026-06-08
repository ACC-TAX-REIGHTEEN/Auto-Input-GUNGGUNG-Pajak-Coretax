import pandas as pd
import os

def create_pre_template():
    input_file = "Pre-moving_temp.xlsx"
    input_sheet = "ExAcc_clean_temp"
    output_file = "Pre_template_temp.xlsx"

    if not os.path.exists(input_file):
        print(f"--> Error: File '{input_file}' tidak ditemukan.")
        return

    print("--> Membaca file source...")
    try:
        df_source = pd.read_excel(input_file, sheet_name=input_sheet)
    except Exception as e:
        print(f"--> Error saat membaca file: {e}")
        return

    df_output = pd.DataFrame()

    try:
        df_output['SerialNo'] = df_source['No Nota']         
        df_output['TransactionDate'] = df_source['Tanggal']     
        df_output['TaxBaseSellingPrice'] = df_source['Tax Base']
        df_output['OtherTaxBaseSellingPrice'] = df_source['OTB']
        df_output['VAT'] = df_source['VAT']                    
    except KeyError as e:
        print(f"--> Error: Kolom {e} tidak ditemukan di file sumber. Pastikan nama header sesuai.")
        return

    row_count = len(df_output)
    
    df_output['TrxCode'] = "Normal"
    df_output['BuyerName'] = "-"
    df_output['BuyerIdOpt'] = "NIK"
    df_output['BuyerIdNumber'] = "0000000000000000"
    df_output['GoodServiceOpt'] = "A"
    df_output['STLG'] = "0"
    df_output['Info'] = "ok"

    final_columns = [
        'TrxCode',
        'BuyerName',
        'BuyerIdOpt',
        'BuyerIdNumber',
        'GoodServiceOpt',
        'SerialNo',
        'TransactionDate',
        'TaxBaseSellingPrice',
        'OtherTaxBaseSellingPrice',
        'VAT',
        'STLG',
        'Info'
    ]

    df_output = df_output[final_columns]

    print("--> Menyimpan file output...")
    df_output.to_excel(output_file, index=False)
    
    print(f"--> Berhasil! File '{output_file}' telah dibuat.")

if __name__ == "__main__":
    create_pre_template()