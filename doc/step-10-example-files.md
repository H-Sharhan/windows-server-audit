# Step 10 — Example Files & How to Run

## Sample `config/servers.txt`

```
# Windows Server Audit - Target Servers
# One hostname or IP per line. Lines starting with # are ignored.

192.168.1.10
192.168.1.11
192.168.1.12
srv-dc01.example.com
srv-db01.example.com
srv-web01.example.com
srv-app01.example.com
```

## Initial Setup Steps

```bash
# 1. Create virtual environment and install dependencies
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt

# 2. Generate encryption key
python src/crypto.py --init

# 3. Store passwords (you will be prompted interactively)
python src/crypto.py --set
```

## How to Run Audit Script

```bash
# Run read-only audit
python src/audit.py

# Output: output/Audit_Report_YYYYMMDD_HHMMSS.xlsx
```

## How to Run Password Change Script

```bash
# Run password change
python src/change_password.py

# You will be prompted for the new password (typed twice for confirmation)
# Output: output/Password_Change_YYYYMMDD_HHMMSS.xlsx

# IMPORTANT: After running, update the encrypted passwords file:
python src/crypto.py --set
# Add the old password(s) back if you want to maintain backward compatibility,
# or only add the new password.
```

## Sample Excel Report Layout

### Summary Sheet

| Server | Login Status | Total Users | Admins | Privileged | Disabled | Built-in | Groups | Tasks | User Tasks |
|--------|-------------|-------------|--------|------------|----------|----------|--------|-------|-----------|
| 192.168.1.10 | OK | 5 | 2 | 0 | 1 | 2 | 15 | 42 | 3 |
| 192.168.1.11 | OK | 4 | 2 | 0 | 0 | 2 | 14 | 38 | 1 |

### Server Login Status Sheet

| Server | Success | Username Used | Error |
|--------|---------|---------------|-------|
| 192.168.1.10 | Yes | administrator | |
| 192.168.1.11 | Yes | admin | |
| 192.168.1.12 | No | | All credential combinations failed |

### Local Users and Privileges Sheet

| Server | Name | Full Name | Enabled | Privilege Level | Privileged Groups | Principal Source | Is Built-in |
|--------|------|-----------|---------|----------------|-------------------|-----------------|-------------|
| 192.168.1.10 | Administrator | | True | Local Administrator | Administrators | Local | True |
| 192.168.1.10 | jdoe | John Doe | True | Standard User | | ActiveDirectory | False |

### Local Groups and Members Sheet

| Server | Group Name | Member Name | Account Type | Enabled | Principal Source |
|--------|-----------|-------------|-------------|---------|-----------------|
| 192.168.1.10 | Administrators | CONTOSO\jdoe | Domain User | True | ActiveDirectory |
| 192.168.1.10 | Administrators | BUILTIN\Administrator | Built-in Account | True | Local |

### Scheduled Tasks Sheet

| Server | Task Name | Path | Run As User | Triggers | Actions | State | Last Run | Next Run | Classification |
|--------|-----------|------|-------------|----------|---------|-------|----------|----------|---------------|
| 192.168.1.10 | MyBackup | \MyTasks\ | CONTOSO\svc-backup | Daily at 02:00 | powershell.exe backup.ps1 | Ready | 2024-12-31 02:00 | 2025-01-01 02:00 | User-Created |

### Errors Sheet

| Server | Operation | Error | Timestamp |
|--------|-----------|-------|-----------|
| 192.168.1.12 | Connect | WinRM transport error: timeout | 2025-01-01 12:00:05 |
| 192.168.1.10 | CollectTasks | Access denied enumerating scheduled tasks | 2025-01-01 12:00:10 |
