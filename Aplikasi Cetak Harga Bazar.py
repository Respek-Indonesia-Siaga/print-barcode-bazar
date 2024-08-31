import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os
import webbrowser
import requests
from io import BytesIO

try:
    import win32print
    import win32ui
    import win32con
    from PIL import ImageWin
except ImportError:
    messagebox.showerror("Error", "Pastikan modul pywin32 sudah terinstall. Jalankan 'pip install pywin32'.")

# Fungsi untuk mengunggah file CSV
def upload_file():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            data = pd.read_csv(file_path)
            create_labels(data)
    except Exception as e:
        messagebox.showerror("Error", f"Gagal mengunggah file: {e}")

# Fungsi untuk mengunduh template CSV
def download_template():
    template_data = {
        "SKU": ["SG-63"],
        "Nama Produk": ["ABSOLUTE FMNM HYGN CHAMOMILE 60ML"],
        "Harga": ["22000"],
        "Jumlah": ["2"]
    }
    template_df = pd.DataFrame(template_data)
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    if file_path:
        template_df.to_csv(file_path, index=False)
        messagebox.showinfo("Sukses", "Template berhasil diunduh")

# Fungsi untuk membungkus teks agar muat dalam lebar yang ditentukan
def wrap_text(draw, text, font, max_width):
    lines = []
    words = text.split()
    while words:
        line = ''
        while words and draw.textbbox((0, 0), line + words[0], font=font)[2] <= max_width:
            line += (words.pop(0) + ' ')
        lines.append(line.strip())
    return '\n'.join(lines)

# Fungsi untuk membuat label dan mencetaknya
def create_labels(data):
    try:
        for index, row in data.iterrows():
            sku = row['SKU']
            name = row['Nama Produk']
            quantity = int(row['Jumlah'])
            price = row['Harga']
            
            for _ in range(quantity):
                # Membuat QR Code
                qr = qrcode.QRCode(version=1, box_size=10, border=5)
                qr.add_data(sku)
                qr.make(fit=True)
                qr_img = qr.make_image(fill='black', back_color='white')
                
                # Membuat gambar label
                label_img = Image.new('RGB', (600, 300), color='white')
                draw = ImageDraw.Draw(label_img)
                
                # Menambahkan QR Code ke label
                qr_img = qr_img.resize((150, 150))
                label_img.paste(qr_img, (10, 10))
                
                # Menambahkan teks ke label
                font_path = "arial.ttf"  # Ganti dengan path ke font TrueType Anda
                font_large = ImageFont.truetype(font_path, 24)
                font_medium = ImageFont.truetype(font_path, 20)
                font_bold = ImageFont.truetype(font_path, 28)
                
                draw.text((150, 10), sku, font=font_large, fill="black")
                
                # Membungkus teks Nama Produk agar sesuai dengan lebar label
                max_width = 350  # Maksimal lebar teks
                wrapped_name = wrap_text(draw, name, font_medium, max_width)
                draw.multiline_text((150, 50), wrapped_name, font=font_medium, fill="black")
                
                draw.text((150, 170), f"Rp {price}", font=font_bold, fill="black")
                
                # Menambahkan kredit
                draw.text((10, 270), "Made by Ghasali from Jasapress", font=font_large, fill="black")
                
                # Menyimpan gambar label sementara
                temp_filename = f"label_{index}.png"
                label_img.save(temp_filename)

                # Mencetak label
                print_label(temp_filename)

                # Menghapus gambar label sementara
                os.remove(temp_filename)
        
        messagebox.showinfo("Selesai", "Label berhasil dicetak")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal membuat label: {e}")

# Fungsi untuk mencetak label
def print_label(filename):
    try:
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(win32print.GetDefaultPrinter())
        hdc.StartDoc(filename)
        hdc.StartPage()

        # Load image
        bmp = Image.open(filename)
        dib = ImageWin.Dib(bmp)
        dib.draw(hdc.GetHandleOutput(), (0, 0, bmp.width, bmp.height))
        
        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()
    except Exception as e:
        messagebox.showerror("Error", f"Gagal mencetak label: {e}")

# Fungsi untuk membuka URL di browser default
def open_url(event):
    webbrowser.open_new("https://jasapress.com")

# Inisialisasi Tkinter
root = tk.Tk()
root.title("JASAPRESS PRICE BARCODE APP")
root.geometry("500x500")

# Menambahkan style untuk ttk
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12), padding=10)
style.configure('TLabel', font=('Helvetica', 12))

# Menambahkan frame untuk menempatkan widget
frame = ttk.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

# Menambahkan label judul
title_label = ttk.Label(frame, text="APLIKASI PRINT BARCODE HARGA", font=('Helvetica', 16, 'bold'))
title_label.pack(pady=10)

# Tombol Unggah File
upload_button = ttk.Button(frame, text="Unggah File CSV", command=upload_file)
upload_button.pack(pady=10)

# Tombol Download Template
download_template_button = ttk.Button(frame, text="Download Template", command=download_template)
download_template_button.pack(pady=10)

# Menambahkan label kredit yang bisa diklik
credits_label = ttk.Label(frame, text="Made by Ghasali from Jasapress", font=('Helvetica', 10), foreground='blue', cursor="hand2")
credits_label.pack(side=tk.BOTTOM, pady=10)
credits_label.bind("<Button-1>", open_url)

# Mendapatkan logo dari URL
logo_url = "https://jasapress.com/wp-content/uploads/2024/02/Jasapress.com-black-text-logo-e1708339730752.webp"
response = requests.get(logo_url)
logo_image = Image.open(BytesIO(response.content))
logo_image = logo_image.resize((150, 37), Image.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)

# Menambahkan label untuk logo
logo_label = tk.Label(frame, image=logo_photo, cursor="hand2")
logo_label.image = logo_photo  # Menyimpan referensi ke image agar tidak terhapus oleh garbage collector
logo_label.pack(side=tk.BOTTOM, pady=10)
logo_label.bind("<Button-1>", open_url)

root.mainloop()
