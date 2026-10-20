# Emission Summary

Windows automation that builds an emissions summary workbook from the APTL "Emissions Utilities" web app.

## What it does

1. Opens Emissions Utilities in Edge or Chrome (Selenium) and waits for sign-in.
2. Fills the Reprint Report query (VIN, requestor, start date, sort order) and downloads the result spreadsheet.
3. Compares it with the latest workbook in `Emission_Summary_Result/` and lists new and updated tests.
4. Either rebuilds the workbook from the template (`fresh`) or applies only the differences (`append`).
5. Reads the emission PDFs in `Emission_Result/` and fills the summary columns (`AZ:BV`) of `QueryResults`.
6. Matches the monitor schedule (SharePoint workbook) to each test and writes the `Monitor` column.
7. Downloads missing ATI `*Combined*.mf4` files into `APTL_ATI/<testid>/` and links them in the workbook.
8. Rebuilds the Base/OBD plot sheets (`Emis_Plot_Base`, `Emis_Plot_OBD`) and the `CATMN` sheet, then opens the result.

Progress is shown in an auto-refreshing status window (`run_status.html` / `run_status.hta`).

## Requirements

- Windows with Microsoft Excel, plus Edge or Chrome
- Python 3.10+
- Access to the Ford network (and the SharePoint monitor schedule)

```powershell
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## Usage

```powershell
run_em_util.bat [options]
# or
python EM_SUM.py [options]
```

| Option | Description |
| --- | --- |
| `--browser {auto,edge,chrome}` | Browser to use (default `auto`, Edge first) |
| `--headless` | Run without a visible browser window |
| `--timeout SECONDS` | Wait time for sign-in and page load (default 300) |
| `--close-on-success` | Close the browser after access is confirmed |
| `--click-target LABEL` | Menu entry to click after login (default `1.1.1`) |
| `--run-mode {fresh,append}` | Skip the prompt and choose full rebuild or append |
| `--vehicle-vin`, `--test-requestor`, `--start-date`, `--sort-by` | Override the query form values |
| `--populate-workbook` | Fill `AZ:BV` from the PDFs in `--pdf-dir` |
| `--workbook-path PATH` | Use a specific workbook instead of the dated copy of the template |
| `--pdf-dir DIR` | PDF folder (default `Emission_Result`) |

Default query values come from `query_settings.json`. Only one run can be active at a time.

## Layout

| Path | Purpose |
| --- | --- |
| `Emission_Summary.py` | Main application |
| `EM_SUM.py` | Entry point |
| `run_em_util.bat`, `run_em_util.ps1` | Launchers (prefer `.venv`) |
| `query_settings.json` | Saved query form values |
| `Emission_Summary_Result/` | Output workbooks (`YYYYMMDD_Emission_Summary.xlsx`) and `Archive/` |
| `Emission_Result/` | Emission summary PDFs |
| `Imported_spreadsheet/` | Downloaded query spreadsheets |
| `APTL_ATI/` | Downloaded ATI MF4 files |

## Not in this repository

Data and generated files are git-ignored: `Imported_spreadsheet/`, `APTL_ATI/`, `Backup/`, `Emission_Result/`, `Emission_Summary_Result/`, `*.mf4`, `*.xls`, `*.xlsx`. This includes the workbook template `Emission summary_automation_templete_2.xlsx`, which must be placed next to the scripts before running.
