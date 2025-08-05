# app.py

import tkinter
import customtkinter
from tkinter import filedialog, messagebox
import pandas as pd
import os
import numpy as np
from datetime import datetime

# Mengatur tema dasar untuk aplikasi
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Aplikasi Perbandingan Excel Mahasiswa")
        self.geometry("700x450")
        self.file_master_path = ""
        self.file_pembanding_path = ""
        self.output_folder_path = ""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")  
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.title_label = customtkinter.CTkLabel(self.main_frame, text="Pilih File untuk Dibandingkan", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.master_button = customtkinter.CTkButton(self.main_frame, text="1. Pilih File Master", command=self.pilih_file_master)
        self.master_button.grid(row=1, column=0, padx=20, pady=10)
        self.master_label = customtkinter.CTkLabel(self.main_frame, text="File belum dipilih", text_color="gray")
        self.master_label.grid(row=2, column=0, padx=20, pady=(0, 10))
        self.pembanding_button = customtkinter.CTkButton(self.main_frame, text="2. Pilih File Pembanding (Survey)", command=self.pilih_file_pembanding)
        self.pembanding_button.grid(row=3, column=0, padx=20, pady=10)
        self.pembanding_label = customtkinter.CTkLabel(self.main_frame, text="File belum dipilih", text_color="gray")
        self.pembanding_label.grid(row=4, column=0, padx=20, pady=(0, 10))
        self.output_button = customtkinter.CTkButton(self.main_frame, text="3. Pilih Lokasi Simpan Hasil", command=self.pilih_folder_output)
        self.output_button.grid(row=5, column=0, padx=20, pady=10)
        self.output_label = customtkinter.CTkLabel(self.main_frame, text="Folder belum dipilih", text_color="gray")
        self.output_label.grid(row=6, column=0, padx=20, pady=(0, 20))
        self.run_button = customtkinter.CTkButton(self.main_frame, text="JALANKAN PERBANDINGAN", height=40, font=customtkinter.CTkFont(size=16, weight="bold"), command=self.jalankan_perbandingan)
        self.run_button.grid(row=7, column=0, padx=20, pady=20, sticky="ew")

    def pilih_file_master(self):
        file_path = filedialog.askopenfilename(title="Pilih File Excel Master", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.file_master_path = file_path
            self.master_label.configure(text=os.path.basename(file_path), text_color="white")

    def pilih_file_pembanding(self):
        file_path = filedialog.askopenfilename(title="Pilih File Excel Pembanding", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.file_pembanding_path = file_path
            self.pembanding_label.configure(text=os.path.basename(file_path), text_color="white")
            
    def pilih_folder_output(self):
        folder_path = filedialog.askdirectory(title="Pilih Folder untuk Menyimpan Hasil")
        if folder_path:
            self.output_folder_path = folder_path
            self.output_label.configure(text=folder_path, text_color="white")

    def jalankan_perbandingan(self):
        if not self.file_master_path or not self.file_pembanding_path or not self.output_folder_path:
            messagebox.showerror("Error", "Harap pilih semua file dan folder yang dibutuhkan terlebih dahulu!")
            return

        try:
            # --- TAHAP 1: BACA, GABUNGKAN, DAN FILTER DATA ---
            df_master = pd.read_excel(self.file_master_path, dtype={'nim': str})
            df_pembanding = pd.read_excel(self.file_pembanding_path, dtype={'NIM': str})
            df_pembanding.rename(columns={'NIM': 'nim'}, inplace=True)
            
            kolom_status = next((col for col in df_pembanding.columns if df_pembanding[col].astype(str).str.lower().isin(['hadir', 'tidak hadir']).any()), None)
            
            if not kolom_status:
                messagebox.showerror("Error", "Kolom status 'Hadir'/'Tidak Hadir' tidak ditemukan!")
                return
            
            df_merged = pd.merge(df_master, df_pembanding[['nim', kolom_status]], on='nim', how='left', indicator=True)
            
            data_hadir = df_merged[df_merged[kolom_status].str.lower() == 'hadir']
            data_tidak_hadir = df_merged[df_merged[kolom_status].str.lower() == 'tidak hadir']
            data_tidak_ditemukan = df_merged[df_merged['_merge'] == 'left_only']

            # --- TAHAP 2: PERSIAPAN DATA FINAL SEBELUM DITULIS ---
            kolom_output_final = ['no', 'kd_fak', 'prodi', 'kd_strata', 'nim', 'nama_lengkap']
            
            datasets = {
                "Hadir": data_hadir.copy(),
                "Tidak Hadir": data_tidak_hadir.copy(),
                "Tidak Ditemukan": data_tidak_ditemukan.copy()
            }
            
            for df in datasets.values():
                if not df.empty:
                    df.sort_values(by='kd_fak', inplace=True)
                    df.insert(0, 'no', range(1, len(df) + 1))

            # --- TAHAP 3: TULIS KE EXCEL DENGAN FORMATING LANJUTAN ---
            # --- PERUBAHAN: Format Nama File dengan Jam ---
            timestamp = datetime.now().strftime("%H%M-%d-%m-%Y")
            nama_file_dinamis = f"data-compare-{timestamp}.xlsx"
            output_filename = os.path.join(self.output_folder_path, nama_file_dinamis)
            
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                for nama_sheet, df_final in datasets.items():
                    # Pilih dan urutkan kolom sesuai format akhir
                    df_to_write = df_final[kolom_output_final] if not df_final.empty else pd.DataFrame(columns=kolom_output_final)
                    
                    # Tulis DataFrame tanpa header, mulai di baris ke-5 (indeks 4)
                    df_to_write.to_excel(writer, sheet_name=nama_sheet, index=False, header=False, startrow=4)
                    
                    # Ambil objek workbook dan worksheet untuk manipulasi
                    workbook  = writer.book
                    worksheet = writer.sheets[nama_sheet]
                    
                    # --- Definisikan semua format yang dibutuhkan ---
                    title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
                    count_format = workbook.add_format({'bold': True, 'font_size': 11})
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    cell_format_bold = workbook.add_format({'bold': True, 'valign': 'vcenter'})
                    
                    # --- Menulis Judul Utama dan Total Data ---
                    main_title = f"LAPORAN DATA MAHASISWA {nama_sheet.upper()}"
                    worksheet.merge_range('A1:F1', main_title, title_format) # Merge A1 sampai F1 untuk judul
                    
                    count_title = f"Total Data: {len(df_final)}"
                    worksheet.write('A2', count_title, count_format)

                    # --- Menulis Header Tabel Secara Manual ---
                    for col_num, value in enumerate(df_to_write.columns.values):
                        worksheet.write(3, col_num, value, header_format)

                    # --- Logika untuk Merge Sel 'kd_fak' ---
                    # Cari indeks kolom 'kd_fak'
                    kd_fak_col_index = df_to_write.columns.get_loc('kd_fak')
                    merge_start = 4 # Data dimulai dari baris ke-5 (indeks 4)
                    
                    for i in range(1, len(df_to_write)):
                        # Cek jika nilai saat ini berbeda dengan nilai sebelumnya
                        if df_to_write.iloc[i, kd_fak_col_index] != df_to_write.iloc[i-1, kd_fak_col_index]:
                            if merge_start < i + 3: # Hanya merge jika ada lebih dari 1 baris
                                worksheet.merge_range(merge_start, kd_fak_col_index, i + 3, kd_fak_col_index, df_to_write.iloc[i-1, kd_fak_col_index], cell_format_bold)
                            merge_start = i + 4
                    
                    # Merge grup terakhir
                    if merge_start < len(df_to_write) + 3:
                        worksheet.merge_range(merge_start, kd_fak_col_index, len(df_to_write) + 3, kd_fak_col_index, df_to_write.iloc[-1, kd_fak_col_index], cell_format_bold)

                    # --- Mengatur Lebar Kolom dan Freeze Panes ---
                    worksheet.set_column('A:A', 5)  # Lebar kolom No
                    worksheet.set_column('B:B', 15) # Lebar kolom kd_fak
                    worksheet.set_column('C:C', 25) # Lebar kolom prodi
                    worksheet.set_column('D:D', 10) # Lebar kolom kd_strata
                    worksheet.set_column('E:E', 15) # Lebar kolom nim
                    worksheet.set_column('F:F', 35) # Lebar kolom nama_lengkap
                    
                    # Bekukan 4 baris teratas
                    worksheet.freeze_panes(4, 0)

            messagebox.showinfo("Sukses", f"Proses perbandingan selesai!\nFile berhasil disimpan di:\n{output_filename}")

        except Exception as e:
            messagebox.showerror("Terjadi Error", f"Sebuah error terjadi:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()