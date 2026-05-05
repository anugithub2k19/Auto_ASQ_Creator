# QRaie Ticket Creator (Automation)

Automates logging into QRaie and creating tickets from an Excel file.

## Prerequisites

- Windows 10/11
- Python 3.10+ (3.11+ recommended)

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install
```

## Configure

Copy the example config and fill in values:

```powershell
copy config.example.yaml config.yaml
```

## Prepare Excel input

Create an Excel file (e.g. `tickets.xlsx`) with a sheet named `Tickets` and these columns:

- `Tenant`
- `Project`
- `Title`
- `Module`
- `Severity`
- `Priority`
- `Issue Category`
- `Sub Category`
- `Owner`
- `Authorized Closer`
- `Description`

Optional columns:

- `Run` (Y/N; default Y if missing)
- `Attachment Path` (file path to upload; optional)

## Run

```powershell
python run.py --config config.yaml --excel tickets.xlsx
```

Outputs:

- `output/results.xlsx` (success/failure per row)
- `output/artifacts/` (screenshots/traces on errors)

## Notes

- If your login flow includes MFA/captcha/SSO, the tool may need a one-time “storage state” capture. See `config.example.yaml`.

