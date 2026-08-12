import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from config_loader import load_config
from file_manager import archive_report
from pdf_report import generate_pdf_report
from excel_reader import read_excel_folder
from validator import validate_dataframe
from data_cleaner import clean_data
from kpi_calculator import calculate_kpis
from report_generator import generate_report
from logger import setup_logger


# -------------------------------------------------
# Rapor oluşturma
# -------------------------------------------------

def run_report(input_folder: Path):
    logger = setup_logger()

    project_root = Path(__file__).resolve().parent.parent
    config = load_config()

    logger.info("GUI üzerinden rapor oluşturma başladı.")

    df = read_excel_folder(input_folder)

    if df.empty:
        raise ValueError(
            "Seçilen klasörde işlenebilecek Excel veya CSV verisi bulunamadı."
        )

    validate_dataframe(df)

    df = clean_data(df)

    if df.empty:
        raise ValueError(
            "Veri temizleme sonrasında işlenecek kayıt kalmadı."
        )

    kpis, region_summary, product_summary = calculate_kpis(df)

    generate_report(
        kpis,
        region_summary,
        product_summary,
        df,
        config,
    )

    generate_pdf_report(kpis, config)

    report_file = project_root / config["output_file"]
    archive_folder = project_root / "data" / "archive"

    archive_report(report_file, archive_folder)

    logger.info("GUI üzerinden rapor başarıyla oluşturuldu.")

    return kpis, report_file


# -------------------------------------------------
# Klasör seçme
# -------------------------------------------------

def select_folder():
    folder = filedialog.askdirectory(
        title="Veri klasörünü seçin"
    )

    if folder:
        folder_var.set(folder)
        status_var.set("Klasör seçildi. Rapor oluşturmaya hazır.")


# -------------------------------------------------
# Rapor oluşturma
# -------------------------------------------------

def generate():
    folder = folder_var.get().strip()

    if not folder:
        messagebox.showwarning(
            "Klasör seçilmedi",
            "Lütfen Excel ve CSV dosyalarının bulunduğu klasörü seçin."
        )
        return

    input_folder = Path(folder)

    if not input_folder.exists() or not input_folder.is_dir():
        messagebox.showerror(
            "Geçersiz klasör",
            "Seçilen klasör bulunamadı veya geçerli bir klasör değil."
        )
        return

    supported_files = list(input_folder.glob("*.xlsx"))
    supported_files += list(input_folder.glob("*.xls"))
    supported_files += list(input_folder.glob("*.csv"))

    if not supported_files:
        messagebox.showwarning(
            "Veri bulunamadı",
            "Seçilen klasörde Excel veya CSV dosyası bulunamadı."
        )
        return

    generate_button.config(state="disabled")
    browse_button.config(state="disabled")

    status_var.set("Rapor oluşturuluyor...")
    root.update_idletasks()

    try:
        kpis, report_file = run_report(input_folder)

        config = load_config()
        currency = config.get("currency", "")

        result_text.set(
            "Rapor başarıyla oluşturuldu.\n\n"
            f"Toplam Satış: {currency}{kpis['Total Revenue']:,.2f}\n"
            f"Toplam Maliyet: {currency}{kpis['Total Cost']:,.2f}\n"
            f"Toplam Kâr: {currency}{kpis['Total Profit']:,.2f}\n"
            f"Kâr Marjı: {kpis['Profit Margin']:.2f}%\n\n"
            f"Dosya:\n{report_file}"
        )

        status_var.set("Rapor başarıyla oluşturuldu.")

        open_report_button.config(state="normal")

        messagebox.showinfo(
            "Başarılı",
            "Excel ve PDF raporları başarıyla oluşturuldu."
        )

    except Exception as error:
        result_text.set(
            "Rapor oluşturulamadı.\n\n"
            "Lütfen seçilen verileri ve dosyaları kontrol edin."
        )

        status_var.set("Rapor oluşturulurken hata oluştu.")

        messagebox.showerror(
            "Hata",
            f"Rapor oluşturulurken hata oluştu:\n\n{error}"
        )

    finally:
        generate_button.config(state="normal")
        browse_button.config(state="normal")


# -------------------------------------------------
# Rapor klasörünü aç
# -------------------------------------------------

def open_report_folder():
    try:
        config = load_config()

        project_root = Path(__file__).resolve().parent.parent
        report_file = project_root / config["output_file"]

        if not report_file.exists():
            messagebox.showwarning(
                "Dosya bulunamadı",
                "Oluşturulmuş bir rapor bulunamadı."
            )
            return

        os.startfile(report_file.parent)

    except Exception as error:
        messagebox.showerror(
            "Hata",
            f"Rapor klasörü açılamadı:\n\n{error}"
        )


# -------------------------------------------------
# GUI
# -------------------------------------------------

root = tk.Tk()

root.title("Excel Reporter")
root.geometry("700x520")
root.resizable(False, False)


# -------------------------------------------------
# Başlık
# -------------------------------------------------

title_label = tk.Label(
    root,
    text="Excel Reporter",
    font=("Arial", 24, "bold"),
)

title_label.pack(pady=(30, 5))


subtitle_label = tk.Label(
    root,
    text="Excel ve CSV verilerinden otomatik rapor oluşturma",
    font=("Arial", 11),
)

subtitle_label.pack(pady=(0, 30))


# -------------------------------------------------
# Klasör seçimi
# -------------------------------------------------

folder_var = tk.StringVar()

folder_frame = tk.Frame(root)
folder_frame.pack(
    padx=40,
    fill="x",
)


folder_entry = tk.Entry(
    folder_frame,
    textvariable=folder_var,
    width=60,
    font=("Arial", 10),
)

folder_entry.pack(
    side="left",
    fill="x",
    expand=True,
)


browse_button = tk.Button(
    folder_frame,
    text="Klasör Seç",
    command=select_folder,
    width=12,
)

browse_button.pack(
    side="left",
    padx=(10, 0),
)


# -------------------------------------------------
# Rapor oluştur butonu
# -------------------------------------------------

generate_button = tk.Button(
    root,
    text="Rapor Oluştur",
    command=generate,
    width=25,
    height=2,
    font=("Arial", 11, "bold"),
)

generate_button.pack(pady=25)


# -------------------------------------------------
# Durum
# -------------------------------------------------

status_var = tk.StringVar(
    value="Rapor oluşturmak için bir veri klasörü seçin."
)

status_label = tk.Label(
    root,
    textvariable=status_var,
    font=("Arial", 10),
)

status_label.pack(
    pady=(0, 15)
)


# -------------------------------------------------
# Sonuç
# -------------------------------------------------

result_text = tk.StringVar()

result_label = tk.Label(
    root,
    textvariable=result_text,
    justify="left",
    anchor="w",
    font=("Arial", 10),
    wraplength=620,
)

result_label.pack(
    padx=40,
    fill="x",
)


# -------------------------------------------------
# Rapor klasörünü aç
# -------------------------------------------------

open_report_button = tk.Button(
    root,
    text="Rapor Klasörünü Aç",
    command=open_report_folder,
    width=20,
    state="disabled",
)

open_report_button.pack(
    pady=15
)


# -------------------------------------------------
# Uygulamayı başlat
# -------------------------------------------------

root.mainloop()