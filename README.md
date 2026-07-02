# Windows Server Audit

Remotely audit ~30 Windows Servers — enumerate local users, groups, and scheduled tasks — and optionally rotate the local Administrator password. All communication is over WinRM, credentials are encrypted with Fernet, and output is a formatted Excel report.

## Features

- **Read-only audit** — collects local users (with privilege classification), local groups (with member details), and scheduled tasks. Never modifies anything.
- **Password rotation** — changes the local Administrator password on every server to a single new value.
- **Credential fallback** — tries `administrator` then `admin` against every encrypted password, first with CredSSP transport then NTLM.
- **Per-server error resilience** — a single server failure never aborts the entire run. Errors are collected per-operation and recorded in the report.
- **Encrypted credentials** — passwords are Fernet-encrypted at rest using `cryptography`. The encryption key and encrypted payload are stored separately.
- **Formatted Excel reports** — styled worksheets with auto-filter, frozen headers, and auto-width columns.

## Directory layout

```
windows-server-audit/
├── config/
│   ├── servers.txt            # one hostname/IP per line
│   ├── secret.key             # Fernet encryption key (git-ignored)
│   └── passwords.enc          # encrypted passwords (git-ignored)
├── output/                    # generated reports (git-ignored)
├── src/
│   ├── crypto.py              # key generation, encrypt/decrypt credentials
│   ├── server_connection.py   # server list loader + WinRM session factory
│   ├── audit_users.py         # local user enumeration & privilege classification
│   ├── audit_groups.py        # local group enumeration & member analysis
│   ├── audit_tasks.py         # scheduled task enumeration & classification
│   ├── report.py              # Excel workbook builder
│   ├── audit.py               # read-only audit entry point
│   └── change_password.py     # password change entry point
├── doc/                       # detailed implementation docs (step-*.md)
├── requirements.txt
├── AGENTS.md
└── README.md
```

## Prerequisites

- Python 3.10+
- Windows targets must have WinRM enabled and configured for CredSSP (or NTLM as fallback).
- The account used must have local administrator privileges on target servers.

## Setup

```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

## Quick start

### 1. Generate encryption key and store passwords

```powershell
python src/crypto.py --init
```

This creates `config/secret.key`. Then:

```powershell
python src/crypto.py --set
```

You will be prompted interactively to enter one or more passwords. These are Fernet-encrypted and written to `config/passwords.enc`.

### 2. Prepare server list

Edit `config/servers.txt` — one hostname or IP per line; blank lines and `#` comments are ignored.

### 3. Run an audit

```powershell
python src/audit.py
```

Produces `output/Audit_Report_YYYYMMDD_HHMMSS.xlsx` with these worksheets:

| Sheet | Contents |
|---|---|
| Summary | Per-server counts of users, admins, privileged users, groups, tasks |
| Server Login Status | Connection success/failure per server |
| Local Users and Privileges | Every user with privilege level, group memberships, enabled status |
| Local Groups and Members | Every group, its members, and account type classification |
| Scheduled Tasks | All tasks with actions, triggers, state, classification |
| Errors | Per-operation errors (only present if errors occurred) |

### 4. Change local Administrator password

```powershell
python src/change_password.py
```

Prompts for a new password (with confirmation), then connects to every server and updates the local Administrator account. Produces `output/Password_Change_YYYYMMDD_HHMMSS.xlsx`.

## Command-line options

Both `audit.py` and `change_password.py` accept:

| Flag | Default | Description |
|---|---|---|
| `--servers` | `config/servers.txt` | Path to server list |
| `--key` | `config/secret.key` | Path to encryption key |
| `--passwords` | `config/passwords.enc` | Path to encrypted passwords |
| `--output` | `output` | Output directory for reports |

## Security

- **Encryption key** (`secret.key`) and **encrypted passwords** (`passwords.enc`) are listed in `.gitignore` — never commit them.
- Passwords are encrypted with Fernet (symmetric AES-based) before touching disk.
- The `audit.py` entrypoint is strictly read-only — no `Set-LocalUser`, `Add-LocalGroupMember`, or destructive PowerShell commands are executed.
- `change_password.py` targets **only** the local Administrator account and nothing else.

## Error handling

- If a server is unreachable or authentication fails, the script logs the error and continues to the next server.
- If user/group/task collection fails on a specific server, that operation is skipped and the error is recorded per-server.
- The run never aborts for a single server failure.
- All errors appear in the Errors worksheet of the report.

## Dependencies

- `pywinrm[credssp]` — WinRM client with CredSSP transport support
- `openpyxl` — Excel workbook generation
- `cryptography` — Fernet symmetric encryption
- `requests-credssp` — CredSSP authentication handler
