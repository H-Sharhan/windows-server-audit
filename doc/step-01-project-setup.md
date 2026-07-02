# Step 1 — Project Setup

## Goal

Create the directory structure, virtual environment, and install all dependencies.

## Dependencies

| Package | Purpose |
|---------|---------|
| `pywinrm` | WinRM remote PowerShell execution |
| `pywinrm[credssp]` | CredSSP for double-hop authentication |
| `openpyxl` | Excel report generation |
| `cryptography` | Fernet symmetric encryption for passwords |
| `requests-credssp` | CredSSP auth support for requests |

## Directory Structure

```
windows-server-audit/
├── src/           # Python modules and scripts
├── config/        # servers.txt, passwords.enc
├── output/        # Generated Excel reports
├── doc/           # Documentation
├── requirements.txt
└── README.md
```

## Implementation

Create `requirements.txt`:

```
pywinrm>=0.4.3
pywinrm[credssp]>=0.4.3
openpyxl>=3.1.2
cryptography>=41.0.0
requests-credssp>=2.0.0
```

Create empty `src/__init__.py` to make it a package.

Create `config/` and `output/` directories.

## WinRM Pre-requisite on Target Servers

Each target Windows Server must have WinRM configured:

```powershell
winrm quickconfig
Enable-PSRemoting -Force
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
```

If using HTTP (5985), the client must also allow unencrypted traffic:

```powershell
Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true
```

> Note: These are one-time server-side configurations, not part of the Python scripts.
