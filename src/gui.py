import sys
import json
import os
import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

STRINGS = {
    "TR": {
        "app_title":        "PartLedger — Satış Rapor Sistemi",
        "subtitle":         "Excel ve CSV verilerinden otomatik rapor oluşturma",
        "select_folder":    "Klasör Seç",
        "folder_hint":      "Veri klasörü seçilmedi",
        "generate":         "Rapor Oluştur",
        "settings":         "Ayarlar",
        "processing":       "İşleniyor, lütfen bekleyin...",
        "files_found":      "dosya bulundu",
        "no_folder":        "Lütfen önce bir veri klasörü seçin.",
        "result_prefix":    "Toplam Ciro",
        "result_profit":    "Kâr",
        "result_margin":    "Kâr Marjı",
        "result_qty":       "Adet",
        "lang_btn":         "EN",
        "err_no_files":     "Seçilen klasörde hiç Excel veya CSV dosyası bulunamadı.\nLütfen doğru klasörü seçtiğinizden emin olun.",
        "err_missing_col":  "Veri dosyasında beklenen kolon bulunamadı.\nGerekli kolonlar: Date, Product, Region, Quantity, UnitPrice, UnitCost",
        "err_negative":     "Veri dosyasında negatif değerler tespit edildi.\nLütfen kaynak veriyi kontrol edin.",
        "err_invalid_date": "Geçersiz tarih formatı tespit edildi.\nTarihler YYYY-MM-DD formatında olmalıdır.",
        "err_generic":      "Rapor oluşturulurken bir hata meydana geldi.\nDetay: ",
        "settings_title":   "Ayarlar",
        "company_name":     "Şirket Adı",
        "report_title":     "Rapor Başlığı",
        "currency":         "Para Birimi",
        "logo_path":        "Logo Dosyası",
        "choose_logo":      "Seç",
        "save":             "Kaydet",
        "cancel":           "İptal",
    },
    "EN": {
        "app_title":        "PartLedger — Sales Report System",
        "subtitle":         "Automated reporting from Excel and CSV data",
        "select_folder":    "Select Folder",
        "folder_hint":      "No data folder selected",
        "generate":         "Generate Report",
        "settings":         "Settings",
        "processing":       "Processing, please wait...",
        "files_found":      "files found",
        "no_folder":        "Please select a data folder first.",
        "result_prefix":    "Total Revenue",
        "result_profit":    "Profit",
        "result_margin":    "Profit Margin",
        "result_qty":       "Quantity",
        "lang_btn":         "TR",
        "err_no_files":     "No Excel or CSV files found in the selected folder.\nPlease make sure you selected the correct folder.",
        "err_missing_col":  "A required column is missing from the data file.\nRequired columns: Date, Product, Region, Quantity, UnitPrice, UnitCost",
        "err_negative":     "Negative values were detected in the data.\nPlease review your source data.",
        "err_invalid_date": "Invalid date format detected.\nDates must be in YYYY-MM-DD format.",
        "err_generic":      "An error occurred while generating the report.\nDetails: ",
        "settings_title":   "Settings",
        "company_name":     "Company Name",
        "report_title":     "Report Title",
        "currency":         "Currency",
        "logo_path":        "Logo File",
        "choose_logo":      "Browse",
        "save":             "Save",
        "cancel":           "Cancel",
    },
}


def get_application_root():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def load_config():
    from config_loader import load_config as _load
    return _load()


def save_config(data):
    root = get_application_root()
    config_file = root / "config.json"
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def open_file(path):
    if path.exists():
        os.startfile(str(path))


def classify_error(exc, lang):
    s = STRINGS[lang]
    msg = str(exc).lower()
    etype = type(exc).__name__.lower()
    if "no excel" in msg or "no files" in msg or "empty" in msg or "no such" in msg:
        return s["err_no_files"]
    if "keyerror" in etype or "column" in msg or "missing" in msg:
        return s["err_missing_col"]
    if "negative" in msg:
        return s["err_negative"]
    if "date" in msg and ("invalid" in msg or "parse" in msg or "format" in msg):
        return s["err_invalid_date"]
    return s["err_generic"] + str(exc)


def count_data_files(folder):
    exts = {".xlsx", ".xls", ".csv"}
    return sum(1 for f in folder.iterdir() if f.suffix.lower() in exts and f.is_file())


class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent_app = parent
        self._lang = parent.lang
        s = STRINGS[self._lang]
        self.title(s["settings_title"])
        self.geometry("480x380")
        self.resizable(False, False)
        self.grab_set()
        config = load_config()
        pad = {"padx": 20, "pady": 6}
        ctk.CTkLabel(self, text=s["company_name"], anchor="w").pack(fill="x", **pad)
        self._company_var = ctk.StringVar(value=config.get("company_name", ""))
        ctk.CTkEntry(self, textvariable=self._company_var, width=420).pack(**pad)
        ctk.CTkLabel(self, text=s["report_title"], anchor="w").pack(fill="x", **pad)
        self._title_var = ctk.StringVar(value=config.get("dashboard_title", ""))
        ctk.CTkEntry(self, textvariable=self._title_var, width=420).pack(**pad)
        ctk.CTkLabel(self, text=s["currency"], anchor="w").pack(fill="x", **pad)
        self._currency_var = ctk.StringVar(value=config.get("currency", "TL"))
        cf = ctk.CTkFrame(self, fg_color="transparent")
        cf.pack(fill="x", padx=20, pady=4)
        for sym in ["TL", "$", "EUR", "GBP"]:
            ctk.CTkRadioButton(cf, text=sym, variable=self._currency_var, value=sym).pack(side="left", padx=10)
        ctk.CTkLabel(self, text=s["logo_path"], anchor="w").pack(fill="x", **pad)
        lf = ctk.CTkFrame(self, fg_color="transparent")
        lf.pack(fill="x", padx=20, pady=4)
        self._logo_var = ctk.StringVar(value=config.get("logo_path", ""))
        ctk.CTkEntry(lf, textvariable=self._logo_var, width=340).pack(side="left")
        ctk.CTkButton(lf, text=s["choose_logo"], width=70, command=self._browse_logo).pack(side="left", padx=8)
        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(fill="x", padx=20, pady=16)
        ctk.CTkButton(bf, text=s["save"], command=self._save, width=120).pack(side="right", padx=4)
        ctk.CTkButton(
            bf, text=s["cancel"], command=self.destroy,
            width=100, fg_color="gray30", hover_color="gray40",
        ).pack(side="right", padx=4)

    def _browse_logo(self):
        path = filedialog.askopenfilename(
            title="Logo seç",
            filetypes=[("Resim", "*.png *.jpg *.jpeg"), ("Tümü", "*.*")],
        )
        if path:
            self._logo_var.set(path)

    def _save(self):
        config = load_config()
        config["company_name"] = self._company_var.get()
        config["dashboard_title"] = self._title_var.get()
        config["currency"] = self._currency_var.get()
        config["logo_path"] = self._logo_var.get()
        save_config(config)
        self.destroy()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.lang = "TR"
        self.folder_var = ctk.StringVar()
        self._app_root = get_application_root()
        self.title(STRINGS[self.lang]["app_title"])
        self.geometry("720x540")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        s = STRINGS[self.lang]
        # Top bar
        top_bar = ctk.CTkFrame(self, height=44, corner_radius=0, fg_color="#1a1a2e")
        top_bar.pack(fill="x")
        top_bar.pack_propagate(False)
        ctk.CTkLabel(
            top_bar, text="PartLedger",
            font=ctk.CTkFont(size=15, weight="bold"), text_color="#4fc3f7",
        ).pack(side="left", padx=14)
        self._lang_btn = ctk.CTkButton(
            top_bar, text=s["lang_btn"], width=60, height=28,
            command=self._toggle_lang, fg_color="#16213e", hover_color="#0f3460", corner_radius=6,
        )
        self._lang_btn.pack(side="right", padx=12, pady=8)
        self._settings_btn = ctk.CTkButton(
            top_bar, text=s["settings"], width=100, height=28,
            command=self._open_settings, fg_color="#16213e", hover_color="#0f3460", corner_radius=6,
        )
        self._settings_btn.pack(side="right", padx=4, pady=8)
        # Hero
        hero = ctk.CTkFrame(self, fg_color="#16213e", corner_radius=0, height=80)
        hero.pack(fill="x")
        hero.pack_propagate(False)
        self._subtitle_lbl = ctk.CTkLabel(
            hero, text=s["subtitle"], font=ctk.CTkFont(size=13), text_color="#b0bec5",
        )
        self._subtitle_lbl.pack(expand=True)
        # Content
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=40, pady=20)
        # Folder card
        folder_card = ctk.CTkFrame(content, corner_radius=12)
        folder_card.pack(fill="x", pady=(0, 12))
        folder_inner = ctk.CTkFrame(folder_card, fg_color="transparent")
        folder_inner.pack(fill="x", padx=16, pady=12)
        self._folder_entry = ctk.CTkEntry(
            folder_inner, textvariable=self.folder_var,
            placeholder_text=s["folder_hint"], height=36, corner_radius=8,
        )
        self._folder_entry.pack(side="left", fill="x", expand=True)
        self._browse_btn = ctk.CTkButton(
            folder_inner, text=s["select_folder"], width=120, height=36,
            command=self._select_folder, corner_radius=8,
        )
        self._browse_btn.pack(side="left", padx=(8, 0))
        self._file_count_lbl = ctk.CTkLabel(
            folder_card, text="", font=ctk.CTkFont(size=11), text_color="#78909c",
        )
        self._file_count_lbl.pack(padx=16, pady=(0, 8), anchor="w")
        # Generate button
        self._gen_btn = ctk.CTkButton(
            content, text=s["generate"], height=48,
            font=ctk.CTkFont(size=15, weight="bold"), corner_radius=12,
            command=self._generate,
        )
        self._gen_btn.pack(fill="x", pady=(0, 12))
        # Progress bar
        self._progress = ctk.CTkProgressBar(content, mode="indeterminate", height=6)
        self._progress.pack(fill="x", pady=(0, 8))
        self._progress.set(0)
        # Result card
        self._result_card = ctk.CTkFrame(content, corner_radius=12)
        self._result_card.pack(fill="both", expand=True)
        self._result_lbl = ctk.CTkLabel(
            self._result_card, text="",
            font=ctk.CTkFont(size=13), justify="left", anchor="nw", wraplength=580,
        )
        self._result_lbl.pack(padx=16, pady=12, fill="both", expand=True)

    def _toggle_lang(self):
        self.lang = "EN" if self.lang == "TR" else "TR"
        s = STRINGS[self.lang]
        self.title(s["app_title"])
        self._lang_btn.configure(text=s["lang_btn"])
        self._settings_btn.configure(text=s["settings"])
        self._subtitle_lbl.configure(text=s["subtitle"])
        self._browse_btn.configure(text=s["select_folder"])
        self._folder_entry.configure(placeholder_text=s["folder_hint"])
        self._gen_btn.configure(text=s["generate"])

    def _select_folder(self):
        folder = filedialog.askdirectory(title="Veri klasörünü seçin")
        if folder:
            self.folder_var.set(folder)
            count = count_data_files(Path(folder))
            s = STRINGS[self.lang]
            if count > 0:
                self._file_count_lbl.configure(
                    text=f"  {count} {s['files_found']}", text_color="#66bb6a",
                )
            else:
                self._file_count_lbl.configure(
                    text=s["err_no_files"].split("\n")[0], text_color="#ef5350",
                )

    def _open_settings(self):
        win = SettingsWindow(self)
        win.focus()

    def _generate(self):
        folder = self.folder_var.get().strip()
        s = STRINGS[self.lang]
        if not folder:
            self._show_result(s["no_folder"], success=False)
            return
        input_folder = Path(folder)
        if not input_folder.exists():
            self._show_result(s["err_no_files"], success=False)
            return
        self._gen_btn.configure(state="disabled", text=s["processing"])
        self._progress.start()
        self._result_lbl.configure(text="")
        t = threading.Thread(target=self._run_report, args=(input_folder,), daemon=True)
        t.start()

    def _run_report(self, input_folder):
        from config_loader import load_config as _load_config
        from file_manager import archive_report
        from pdf_report import generate_pdf_report
        from excel_reader import read_excel_folder
        from validator import validate_dataframe
        from data_cleaner import clean_data
        from kpi_calculator import calculate_kpis
        from report_generator import generate_report
        from logger import setup_logger
        try:
            logger = setup_logger()
            config = _load_config()
            output_folder = self._app_root / "data" / "output"
            archive_folder = self._app_root / "data" / "archive"
            output_folder.mkdir(parents=True, exist_ok=True)
            archive_folder.mkdir(parents=True, exist_ok=True)
            logger.info("GUI: rapor oluşturma başladı.")
            df = read_excel_folder(input_folder)
            validate_dataframe(df)
            df = clean_data(df)
            kpis, region_summary, product_summary, monthly_summary = calculate_kpis(df)
            generate_report(kpis, region_summary, product_summary, monthly_summary, df, config)
            generate_pdf_report(kpis, region_summary, product_summary, monthly_summary, config)
            report_file = self._app_root / config["output_file"]
            pdf_file = self._app_root / "data" / "output" / "Report.pdf"
            archive_report(report_file, archive_folder)
            logger.info("GUI: rapor başarıyla oluşturuldu.")
            self.after(0, lambda: self._on_success(kpis, report_file, pdf_file))
        except Exception as exc:
            err_msg = classify_error(exc, self.lang)
            self.after(0, lambda: self._on_error(err_msg))

    def _on_success(self, kpis, report_file, pdf_file):
        s = STRINGS[self.lang]
        currency = load_config().get("currency", "TL")
        rev = kpis['Total Revenue']
        prf = kpis['Total Profit']
        margin = kpis['Profit Margin']
        qty = kpis['Total Quantity']
        result_text = (
            f"{s['result_prefix']}: {currency}{rev:,.0f}\n"
            f"{s['result_profit']}: {currency}{prf:,.0f}\n"
            f"{s['result_margin']}: {margin:.2f}%\n"
            f"{s['result_qty']}: {qty:,.0f}\n\n"
            f"{report_file}"
        )
        self._show_result(result_text, success=True)
        self._stop_progress()
        open_file(report_file)
        open_file(pdf_file)

    def _on_error(self, message):
        self._show_result(message, success=False)
        self._stop_progress()

    def _show_result(self, text, success):
        color = "#66bb6a" if success else "#ef5350"
        self._result_lbl.configure(text=text, text_color=color)

    def _stop_progress(self):
        s = STRINGS[self.lang]
        self._progress.stop()
        self._progress.set(0)
        self._gen_btn.configure(state="normal", text=s["generate"])


if __name__ == "__main__":
    app = App()
    app.mainloop()
