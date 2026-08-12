import sys
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


def get_application_root():
    """
    Uygulamanın çalıştığı ana dizini döndürür.

    Normal Python çalıştırmasında:
        Proje kökü

    PyInstaller EXE çalıştırmasında:
        EXE'nin bulunduğu klasör
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent.parent


def run_report(input_folder: Path):
    logger = setup_logger()

    application_root = get_application_root()
    config = load_config()

    output_folder = application_root / "data" / "output"
    archive_folder = application_root / "data" / "archive"

    output_folder.mkdir(parents=True, exist_ok=True)
    archive_folder.mkdir(parents=True, exist_ok=True)

    logger.info("GUI üzerinden rapor oluşturma başladı.")

    df = read_excel_folder(input_folder)

    validate_dataframe(df)

    df = clean_data(df)

    kpis, region_summary, product_summary = calculate_kpis(df)

    generate_report(
        kpis,
        region_summary,
        product_summary,
        df,
        config,
    )

    generate_pdf_report(kpis, config)

    report_file = application_root / config["output_file"]

    archive_report(
        report_file,
        archive_folder,
    )

    logger.info("GUI üzerinden rapor başarıyla oluşturuldu.")

    return kpis, report_file


def select_folder():
    folder = filedialog.askdirectory(
        title="Veri klasörünü seçin"
    )

    if folder:
        folder_var.set(folder)


def generate():
    folder = folder_var.get()

    if not folder:
        messagebox.showwarning(
            "Klasör seçilmedi",
            "Lütfen veri klasörünü seçin."
        )
        return

    input_folder = Path(folder)

    if not input_folder.exists():
        messagebox.showerror(
            "Hata",
            "Seçilen klasör bulunamadı."
        )
        return

    try:
        kpis, report_file = run_report(input_folder)

        result_text.set(
            f"Rapor başarıyla oluşturuldu.\n\n"
            f"Toplam Satış: {kpis['Total Revenue']:,.2f}\n"
            f"Toplam Maliyet: {kpis['Total Cost']:,.2f}\n"
            f"Toplam Kâr: {kpis['Total Profit']:,.2f}\n"
            f"Kâr Marjı: {kpis['Profit Margin']:.2f}%\n\n"
            f"Dosya:\n{report_file}"
        )

        messagebox.showinfo(
            "Başarılı",
            "Rapor başarıyla oluşturuldu."
        )

    except Exception as error:
        result_text.set(
            "Rapor oluşturulamadı.\n\n"
            "Lütfen seçilen verileri ve dosyaları kontrol edin."
        )

        messagebox.showerror(
            "Hata",
            f"Rapor oluşturulurken hata oluştu:\n\n{error}"
        )


# -------------------------------------------------
# GUI
# -------------------------------------------------

root = tk.Tk()

root.title("Excel Reporter")
root.geometry("650x450")
root.resizable(False, False)


title_label = tk.Label(
    root,
    text="Excel Reporter",
    font=("Arial", 22, "bold"),
)

title_label.pack(pady=(30, 5))


subtitle_label = tk.Label(
    root,
    text="Excel ve CSV verilerinden otomatik rapor oluşturma",
    font=("Arial", 11),
)

subtitle_label.pack(pady=(0, 30))


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


generate_button = tk.Button(
    root,
    text="Rapor Oluştur",
    command=generate,
    width=25,
    height=2,
)

generate_button.pack(
    pady=30
)


result_text = tk.StringVar()


result_label = tk.Label(
    root,
    textvariable=result_text,
    justify="left",
    anchor="w",
    font=("Arial", 10),
)

result_label.pack(
    padx=40,
    fill="x",
)


root.mainloop()