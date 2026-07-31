from pathlib import Path
from shutil import copy2
from datetime import datetime


def archive_report(report_file: Path, archive_folder: Path):
    archive_folder.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    archive_file = archive_folder / f"Report_{timestamp}.xlsx"

    copy2(report_file, archive_file)

    print(f"Archived: {archive_file.name}")