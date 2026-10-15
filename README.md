# AFP Data Engineering 
 
## Overview 
Python and R workflows for Acute Flaccid Paralysis (AFP) surveillance data engineering, processing, validation, and analysis. 
 
## Objectives 
- Automate AFP surveillance data extraction. 
- Support Microsoft Access MDB and ACCDB databases. 
- Standardize data cleaning and transformation. 
- Implement automated data quality checks. 
- Generate analytical datasets and surveillance indicators. 
- Support reproducible reporting and analytics. 
 
## Data Security 
Surveillance datasets and credentials are excluded from Git tracking. 

## SharePoint collection

The existing AFRO line list fetcher is integrated at
`src/extraction/fetch_afro_line_list.py`. It uses Microsoft Graph app credentials,
validates the XLSX archive, backs up the previous file under `data/raw/backups/`,
and atomically replaces `data/raw/afro_line_list.xlsx` after a successful download.
Failures return a nonzero exit code and leave the previous workbook intact.

Install dependencies: `python -m pip install -r requirements.txt`.
Copy `config/secrets.env.example` to `config/secrets.env` and supply the existing
three credentials locally, or provide them as environment variables.
Credentials are never copied automatically or committed.

Run from any working directory using the absolute path to `run_pipeline.py`:

```powershell
python run_pipeline.py
# Reuse the original credential file without copying its contents:
python run_pipeline.py --secrets-env "C:/Users/TOURE/Documents/Epidemiology/afro_line_list_fetcher/config/secrets.env"
# Refresh first, then run the existing vaccination workflow:
python run_pipeline.py --reconcile
```

`--share-url` overrides the original line list link. A failed collection prevents
reconciliation from running. The line list is a separate source: reconciliation
still reads the Access database and vaccination workbook configured in
`src/transformation/reconcile_vaccination.py`. The extraction stub
`src/extraction/extract_afp_mdb.py` remains empty.

For unattended runs, configure Windows Task Scheduler to run your Python
executable with the absolute path to `run_pipeline.py` and `--secrets-env` as
arguments. Use an account with access to the credential file and project, and
choose “Do not start a new instance” for overlapping runs. No recurring task
is installed by this integration; choose a collection frequency first.
