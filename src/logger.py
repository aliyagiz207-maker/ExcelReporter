from pathlib import Path
from loguru import logger


def setup_logger():
    """
    Programın log ayarlarını yapar.
    """

    project_root = Path(__file__).resolve().parent.parent

    log_folder = project_root / "logs"

    log_folder.mkdir(exist_ok=True)

    logger.remove()

    logger.add(
        log_folder / "app.log",
        rotation="1 MB",
        retention="10 days",
        level="INFO",
        encoding="utf-8"
    )

    logger.add(
        lambda msg: print(msg, end=""),
        level="INFO"
    )

    return logger