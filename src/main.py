import argparse
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


def main():
    parser = argparse.ArgumentParser(
        description="Excel Reporter"
    )

    parser.add_argument(
        "--month",
        nargs="*",
        help="Read only selected Excel files (without .xlsx)"
    )

    args = parser.parse_args()

    logger = setup_logger()

    try:
        logger.info("Program başlatıldı.")

        project_root = Path(__file__).resolve().parent.parent
        config = load_config()

        input_folder = project_root / "data" / "input"

        if not input_folder.exists():
            raise FileNotFoundError(
                f"Input klasörü bulunamadı: {input_folder}"
            )

        df = read_excel_folder(
            input_folder,
            selected_months=args.month,
        )

        logger.info("Excel dosyaları başarıyla okundu.")

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

        report_file = project_root / config["output_file"]
        archive_folder = project_root / "data" / "archive"

        archive_report(
            report_file,
            archive_folder,
        )

        print("\n=== İlk 5 Satır ===")
        print(df.head())

        print("\n=== Veri Bilgisi ===")
        df.info()

        print("\n=== KPI Results ===")
        for key, value in kpis.items():
            if isinstance(value, float):
                print(f"{key}: {value:.2f}")
            else:
                print(f"{key}: {value}")

        logger.info("Program başarıyla tamamlandı.")

    except FileNotFoundError as error:
        logger.error(f"Dosya veya klasör bulunamadı: {error}")
        print(f"\nHATA: {error}")

    except ValueError as error:
        logger.error(f"Geçersiz veri: {error}")
        print(f"\nHATA: {error}")

    except Exception as error:
        logger.exception(
            f"Beklenmeyen bir hata oluştu: {error}"
        )
        print(
            "\nHATA: Beklenmeyen bir hata oluştu. "
            "Detaylar logs/app.log dosyasına yazıldı."
        )


if __name__ == "__main__":
    main()