# Excel Reporter

Excel Reporter is a Python application that automates business Excel reporting.

The application reads one or multiple Excel files, validates and cleans the data, calculates KPIs, generates a professional Excel dashboard, creates a PDF summary report, and archives every generated report automatically.

---

## Features

- Read one or multiple Excel files
- Select specific months from the command line
- Automatic data validation
- Automatic data cleaning
- KPI calculation
- Revenue by Region analysis
- Top Products analysis
- Professional Excel Dashboard
- Detail Data sheet with:
  - Auto Filter
  - Freeze Header
  - Auto column width
- Company logo support
- JSON configuration
- Automatic report archiving
- PDF report generation
- Logging with Loguru

---

## Technologies

- Python 3
- Pandas
- OpenPyXL
- ReportLab
- Loguru

---

## Project Structure

```text
ExcelReporter/
│
├── assets/
├── data/
│   ├── archive/
│   ├── input/
│   └── output/
├── logs/
├── src/
├── config.json
├── requirements.txt
└── README.md
```

---

## Installation

```bash
git clone https://github.com/aliyagiz207-maker/ExcelReporter.git

cd ExcelReporter

python -m venv .venv

.venv\Scripts\activate

pip install -r requirements.txt
```

---

## Usage

Run with all Excel files:

```bash
python src/main.py
```

Run only selected months:

```bash
python src/main.py --month january
```

```bash
python src/main.py --month january march
```

---

## Output

The application generates:

- Excel Dashboard (.xlsx)
- PDF Report (.pdf)
- Archived Excel Reports
- Application Logs

---

## Configuration

All application settings can be changed from:

```
config.json
```

Example:

```json
{
    "company_name": "Honda Terakki",
    "dashboard_title": "Monthly Sales Dashboard",
    "currency": "$",
    "logo_path": "assets/logo.png",
    "output_file": "data/output/Honda_Report.xlsx"
}
```

---

## Author

Ali Yağız