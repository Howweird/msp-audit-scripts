# MSP Audit Scripts

Universal Windows audit scripts for MSP use. Run via ScreenConnect Toolbox as SYSTEM.

## Scripts

| Script | Target | Output |
|--------|--------|--------|
| `server_audit.ps1` | Windows Server 2012-2025 | `C:\Temp\server_audit_YYYY-MM-DD.txt` + `.json` |
| `workstation_audit.ps1` | Windows 10/11 | `C:\Temp\workstation_audit_YYYY-MM-DD.txt` + `.json` |

## ScreenConnect Toolbox Commands

### Server Audit
```powershell
#!ps
#maxlength=500000
#timeout=600000
Set-ExecutionPolicy Bypass -Scope Process -Force
New-Item -Path C:\Temp -ItemType Directory -Force | Out-Null
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Howweird/msp-audit-scripts/master/server_audit.ps1" -OutFile "C:\Temp\server_audit.ps1" -UseBasicParsing
. C:\Temp\server_audit.ps1
```

### Workstation Audit
```powershell
#!ps
#maxlength=500000
#timeout=600000
Set-ExecutionPolicy Bypass -Scope Process -Force
New-Item -Path C:\Temp -ItemType Directory -Force | Out-Null
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Howweird/msp-audit-scripts/master/workstation_audit.ps1" -OutFile "C:\Temp\workstation_audit.ps1" -UseBasicParsing
. C:\Temp\workstation_audit.ps1
```

## Requirements

- Must run as SYSTEM or local Administrator
- Output directory: `C:\Temp` (created automatically)
- No dependencies beyond built-in Windows PowerShell modules

## Output

- `.txt` — Full transcript (human-readable)
- `.json` — Structured data for AI/automated processing
