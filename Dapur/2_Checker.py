import pandas as pd
import os
from openpyxl.utils import get_column_letter

def cek_data_excel(nama_file):
    print(f"--> Memulai proses pengecekan untuk file: {nama_file}")
    
    peta_gudang = {}
    if os.path.exists('config.conf'):
        with open('config.conf', 'r') as file:
            for line in file:
                line = line.strip()
                if "=" in line and not line.startswith("["):
                    kunci, nilai = line.split("=", 1)
                    kunci = kunci.strip()
                    nilai = nilai.strip()
                    
                    if kunci in peta_gudang:
                        peta_gudang[kunci].append(nilai)
                    else:
                        peta_gudang[kunci] = [nilai]
    
    df = pd.read_excel(nama_file)
    data_salah = []
    
    print(f"--> Memeriksa total {len(df)} baris data...")
    
    for index, row in df.iterrows():
        dept = str(row['Nama Dept.']).strip()
        kategori_aktual = str(row['Kategori']).strip()
        gudang_aktual = str(row['Gudang']).strip()
        
        keterangan_error = []
        
        if " - " in dept:
            kategori_harapan = dept.split(" - ")[0].strip()
        else:
            kategori_harapan = dept
            
        if kategori_aktual != kategori_harapan:
            keterangan_error.append("Tidak sesuai Kategori")
            
        gudang_harapan_list = peta_gudang.get(dept)
        
        if gudang_harapan_list is not None and gudang_aktual not in gudang_harapan_list:
            keterangan_error.append("Tidak sesuai Gudang")
            
        if keterangan_error:
            baris_baru = row.to_dict()
            baris_baru['Keterangan'] = " dan ".join(keterangan_error)
            data_salah.append(baris_baru)
            
    if data_salah:
        df_laporan = pd.DataFrame(data_salah)
        nama_file_baru = nama_file.replace("_temp", "_report")
        
        with pd.ExcelWriter(nama_file_baru, engine='openpyxl') as writer:
            df_laporan.to_excel(writer, index=False, sheet_name='Laporan')
            worksheet = writer.sheets['Laporan']
            
            for index_kolom, nama_kolom in enumerate(df_laporan.columns):
                panjang_isi_kolom = df_laporan[nama_kolom].astype(str).map(len).max()
                panjang_judul_kolom = len(str(nama_kolom))
                lebar_maksimal = max(panjang_isi_kolom, panjang_judul_kolom) + 2
                
                huruf_kolom = get_column_letter(index_kolom + 1)
                worksheet.column_dimensions[huruf_kolom].width = lebar_maksimal

        print(f"--> Laporan berhasil dibuat: {nama_file_baru} ({len(data_salah)} data salah ditemukan)")
    else:
        print(f"--> Pengecekan selesai. Tidak ditemukan data yang salah pada {nama_file}.")
    
    print("--> --------------------------------------------------")

daftar_file = ["PTMUS_temp.xlsx", "PTMSU_temp.xlsx"]

for file_excel in daftar_file:
    if os.path.exists(file_excel):
        cek_data_excel(file_excel)
    else:
        print(f"--> File {file_excel} tidak ditemukan di dalam folder")
