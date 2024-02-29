import os
import openpyxl
from tkinter import filedialog
from tkinter import Tk

# Klasör seçme penceresini açın
root = Tk()
root.withdraw()
klasor_yolu = filedialog.askdirectory(title="Klasör Seç")

if not klasor_yolu:
    print("Klasör seçilmedi. İşlem iptal edildi.")
else:
    excel_dosyasi_yolu = filedialog.askopenfilename(title="Excel Dosyasını Seç", filetypes=[("Excel Dosyaları", "*.xlsx")])

    if not excel_dosyasi_yolu:
        print("Excel dosyası seçilmedi. İşlem iptal edildi.")
    else:
        excel_dosyasi = openpyxl.load_workbook(excel_dosyasi_yolu)
        sayfa = excel_dosyasi.active

        dosya_adlari = [sayfa.cell(row=i, column=4).value for i in range(2, sayfa.max_row + 1) if sayfa.cell(row=i, column=4).value is not None]

        klasordeki_dosyalar = os.listdir(klasor_yolu)

        for dosya_adi in klasordeki_dosyalar:
            if dosya_adi not in dosya_adlari:
                dosya_yolu = os.path.join(klasor_yolu, dosya_adi)
                if os.path.exists(dosya_yolu):
                    os.remove(dosya_yolu)
                    print(f"{dosya_adi} onaylanmadı.")

        excel_dosyasi.close()
