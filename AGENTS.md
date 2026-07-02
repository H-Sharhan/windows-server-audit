# AGENTS.md — windows-server-audit

## Project

Two Python scripts to remotely audit ~30 Windows Servers and change the local Administrator password. WinRM-based (pywinrm[credssp]), credentials encrypted via Fernet, output is an openpyxl Excel report.

## Source of Truth

- `doc/plan.md` — architecture, dependency rationale, design decisions
- `doc/step-*.md` — implementation details per module, PowerShell commands, data structures

## Directory Layout (planned)

```
src/
├── crypto.py              # Fernet encrypt/decrypt for credentials
├── server_connection.py   # servers.txt loader + WinRM session factory
├── audit_users.py         # local users & privileges
├── audit_groups.py        # local groups & members
├── audit_tasks.py         # scheduled tasks
├── report.py              # Excel workbook builder
├── audit.py               # read-only audit orchestrator (entrypoint)
└── change_password.py     # password change entrypoint
config/
├── servers.txt
└── passwords.enc
output/
└── Audit_Report_*.xlsx
```

## Implementation Order

Steps match `doc/step-{01..10}-*.md`. Build in order: project setup → crypto → connection → users → groups → tasks → report → audit script → password change → examples.

## Key Constraints

- `audit.py` must be **strictly read-only**: no user/group/task modifications ever.
- `change_password.py` changes **only** the local Administrator password — nothing else.
- Credential fallback: try `administrator` then `admin` × all encrypted passwords.
- Never abort entire run for one server failure; collect errors per-server.

## Dependencies

```
pywinrm[credssp]>=0.4.3  # WinRM + CredSSP for double-hop
openpyxl>=3.1.2          # Excel reports
cryptography>=41.0.0     # Fernet encryption
requests-credssp>=2.0.0  # CredSSP transport
```

## Setup Flow

```bash
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
python src/crypto.py --init        # generates config/secret.key
python src/crypto.py --set         # interactive: enter passwords → config/passwords.enc
```

## Code Style

- All public functions **must** have a Google-style docstring (Args / Returns / Raises).
- Use type hints on all function parameters and return values.
- Never add inline comments to explain *what* the code does — use docstrings for interface docs.
- Keep functions single-responsibility; one clear purpose per function.

## Running

```bash
python src/audit.py                # full audit → output/Audit_Report_*.xlsx
python src/change_password.py      # prompts for new pwd → output/Password_Change_*.xlsx
```
