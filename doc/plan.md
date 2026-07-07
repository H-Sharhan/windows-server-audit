# Windows Server Audit & Administrator Password Update — Implementation Plan

## Overview

Two Python scripts that manage ~30 Windows Servers remotely:
1. **Audit Script** — read-only collection of local users, groups, scheduled tasks.
2. **Password Change Script** — change local Administrator password only.

All credentials stored in an encrypted file. Communication via WinRM (pywinrm). Output is a formatted Excel report.

## Steps

| Step | File | Description |
|------|------|-------------|
| 1 | `step-01-project-setup.md` | Directory layout, virtualenv, dependencies |
| 2 | `step-02-encrypted-passwords.md` | Encrypted credential storage (`crypto.py`) |
| 3 | `step-03-server-connection.md` | Server list loader + WinRM connection (`server_connection.py`) |
| 4 | `step-04-local-users.md` | Enumerate local users and privileges |
| 5 | `step-05-local-groups.md` | Enumerate local groups and members |
| 6 | `step-06-scheduled-tasks.md` | Enumerate and classify scheduled tasks |
| 7 | `step-07-excel-report.md` | Generate formatted Excel workbook |
| 8 | `step-08-audit-script.md` | Main audit orchestrator (`audit.py`) |
| 9 | `step-09-password-change.md` | Password change script (`change_password.py`) |
| 10 | `step-10-example-files.md` | Sample `servers.txt`, `passwords.enc`, how-to-run |

## Architecture

```
windows-server-audit/
├── doc/
│   ├── plan.md
│   └── step-*.md
├── src/
│   ├── crypto.py              # Fernet encrypt/decrypt (standalone + imported)
│   ├── audit.py               # Read-only audit: connection, users, groups,
│   │                          # tasks, Excel report, orchestrator
│   └── change_password.py     # Password change: connection, Excel report,
│                              # change logic
├── config/
│   ├── servers.txt
│   └── passwords.enc
├── output/
│   └── (generated reports)
├── requirements.txt
└── README.md
```

## Design Decisions

- **WinRM** over WMI — modern, supported, works with PowerShell.
- **pywinrm[credssp]** — required for double-hop scenarios and HTTP WinRM.
- **openpyxl** — Excel generation with formatting, filters, auto-width.
- **cryptography.fernet** — symmetric encryption for password file.
- Each script is self-contained (only `crypto.py` is shared between the two entry points).
- Error collection per-server, per-operation — never abort entire run on one failure.
