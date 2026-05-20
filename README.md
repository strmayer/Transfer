# Transfer

## Overview

This repository contains `get-sqlparameters.ps1`, a PowerShell script that collects SQL Server environment and configuration details from an Autodesk Vault source instance.

The script is designed for Vault migration planning from 2024 to 2026 and generates:

- a text report (`.txt`)
- a structured JSON export (`.json`)

## Usage

1. Open PowerShell in the repository folder.
2. Run the script:

   ```powershell
   .\get-sqlparameters.ps1
   ```

3. Enter the requested values:
   - SQL Server instance name (default: `COMPUTERNAME\AUTODESKVAULT`)
   - authentication mode
   - SQL login and password (for SQL Authentication)
   - output file name base

4. After the script completes, the report files are saved in the chosen output path.

## Output

- `*.txt` — readable migration inventory report
- `*.json` — structured data export for machine processing

## Requirements

- PowerShell 5.1 or later
- Windows environment with network access to the SQL Server instance
- SQL permissions to query metadata and configuration

## Notes

The script does not require external PowerShell modules.

## Author

- JM Consulting

## Creation Date

- 2026-05-12