import os
import sys
import shutil
import pandas as pd
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from PIL import Image, ImageEnhance
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox

# ضبط مسار Tesseract ليعمل داخل EXE
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # عند تشغيل EXE
else:
    base_path = os.path.dirname(os.path.abspath(__file__))  # عند التشغيل العادي

tesseract_path = os.path.join(base_path, "Tesseract-OCR", "tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# واجهة المستخدم
root = tk.Tk()
root.title("Barcode Reader - Anwer Ghallab Saeed")
root.geometry("600x400")

# متغيرات المسارات
excel_path_var = tk.StringVar()
images_path_var = tk.StringVar()
output_path_var = tk.StringVar()

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    excel_path_var.set(file_path)

def browse_images_folder():
    folder_path = filedialog.askdirectory()
    images_path_var.set(folder_path)

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    output_path_var.set(folder_path)

# تحسين جودة الصور
def enhance_image(image_path):
    try:
        image = Image.open(image_path)
        image = ImageEnhance.Sharpness(image).enhance(2.0)
        image = ImageEnhance.Contrast(image).enhance(1.5)
        image.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
        return image
    except:
        return None

# استخراج الباركود من الصور
def extract_barcode(image_path):
    image = cv2.imread(image_path)
    if image is None:
        return None
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    barcodes = decode(gray)
    return barcodes[0].data.decode("utf-8") if barcodes else None

# استخراج النصوص (OCR)
def extract_text(image_path):
    try:
        return pytesseract.image_to_string(Image.open(image_path)).strip()
    except:
        return ""

def start_search():
    excel_path = excel_path_var.get()
    images_path = images_path_var.get()
    output_path = output_path_var.get()

    if not (excel_path and images_path and output_path):
        messagebox.showerror("خطأ", "يرجى اختيار جميع المسارات.")
        return

    try:
        df = pd.read_excel(excel_path, dtype=str)
    except Exception as e:
        messagebox.showerror("خطأ", f"لا يمكن قراءة ملف Excel:\n{str(e)}")
        return

    if "External Code" not in df.columns:
        messagebox.showerror("خطأ", "لا يوجد عمود باسم 'External Code' في ملف Excel.")
        return

    valid_barcodes = df["External Code"].dropna().astype(str).tolist()

    processed, matched = 0, 0
    not_found = []
    found_data = []
    duplicate_data = []

    for root_folder, sub_folders, files in os.walk(images_path):
        found = False
        for file in files:
            file_path = os.path.join(root_folder, file)
            enhanced_img = enhance_image(file_path)
            if enhanced_img:
                tmp_path = file_path + "_enhanced.jpg"
                enhanced_img.save(tmp_path, quality=90)
                barcode = extract_barcode(tmp_path) or extract_text(tmp_path)
                os.remove(tmp_path)
            else:
                barcode = extract_barcode(file_path) or extract_text(file_path)

            if barcode and barcode in valid_barcodes:
                row_data = df[df["External Code"] == barcode].iloc[0]
                folder_name = "_".join(map(str, row_data.values))
                new_folder_path = os.path.join(output_path, folder_name)

                if os.path.exists(new_folder_path):
                    duplicate_data.append(row_data.tolist())
                else:
                    shutil.move(root_folder, new_folder_path)
                    found_data.append(row_data.tolist())

                matched += 1
                found = True
                break

        if not found:
            not_found.append(root_folder)

        processed += 1

    # إنشاء تقرير Excel
    report_file = os.path.join(output_path, "search_report.xlsx")
    with pd.ExcelWriter(report_file) as writer:
        pd.DataFrame(found_data).to_excel(writer, sheet_name="Found", index=False)
        pd.DataFrame(duplicate_data).to_excel(writer, sheet_name="Duplicate", index=False)
        pd.DataFrame(not_found, columns=["Not Found"]).to_excel(writer, sheet_name="Not Found", index=False)

    messagebox.showinfo(
        "النتائج",
        f"تمت معالجة {processed} مجلد.\n"
        f"تم نقل {matched} مجلد مطابق.\n"
        f"لم يتم العثور على {len(not_found)} مجلد."
    )

# أزرار الواجهة
tk.Button(root, text="اختيار ملف Excel", command=browse_excel_file).pack()
tk.Button(root, text="اختيار مجلد الصور", command=browse_images_folder).pack()
tk.Button(root, text="اختيار مجلد الإخراج", command=browse_output_folder).pack()
tk.Button(root, text="بدء البحث", command=start_search).pack()

root.mainloop()