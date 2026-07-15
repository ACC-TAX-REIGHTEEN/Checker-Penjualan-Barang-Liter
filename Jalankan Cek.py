import os
import glob
import shutil
import subprocess
import sys

def jalankan_otomatisasi():
    folder_dapur = "Dapur"
    file_syarat = ["1_CleaningAcc.py", "2_Checker.py", "config.conf"]
    
    if not os.path.exists(folder_dapur) or not os.path.isdir(folder_dapur):
        print("--> Folder Dapur tidak ditemukan.")
        input("--> Tekan enter untuk keluar.")
        return

    for file in file_syarat:
        jalur_file = os.path.join(folder_dapur, file)
        if not os.path.isfile(jalur_file):
            print(f"--> File {file} tidak ditemukan di dalam folder Dapur.")
            input("--> Tekan enter untuk keluar.")
            return

    file_excel_lama = glob.glob(os.path.join(folder_dapur, "*.xls*"))
    for file in file_excel_lama:
        os.remove(file)

    file_sumber = ["PTMSU.xls", "PTMUS.xls"]
    ada_file_dipindah = False
    for file in file_sumber:
        if os.path.isfile(file):
            shutil.copy2(file, os.path.join(folder_dapur, file))
            ada_file_dipindah = True

    if not ada_file_dipindah:
        print("--> File PTMSU.xls atau PTMUS.xls tidak ditemukan untuk diproses.")
        input("--> Tekan enter untuk keluar.")
        return

    print("--> Memulai eksekusi 1_CleaningAcc.py")
    subprocess.run([sys.executable, "1_CleaningAcc.py"], cwd=folder_dapur)

    print("--> Memulai eksekusi 2_Checker.py")
    subprocess.run([sys.executable, "2_Checker.py"], cwd=folder_dapur)

    file_laporan = glob.glob(os.path.join(folder_dapur, "*report.xlsx"))
    for laporan in file_laporan:
        nama_file = os.path.basename(laporan)
        shutil.move(laporan, nama_file)

    file_excel_sisa = glob.glob(os.path.join(folder_dapur, "*.xls*"))
    for file in file_excel_sisa:
        os.remove(file)

    print("--> Semua proses telah selesai dijalankan.")
    input("--> Tekan enter untuk keluar.")

if __name__ == "__main__":
    jalankan_otomatisasi()