# Step 8 — Audit Script (Main Orchestrator)

## Goal

Orchestrate the full read-only audit: connect to each server, collect all data, generate Excel report.

## File: `src/audit.py`

### Flow

```
1. Load servers from config/servers.txt
2. Load encrypted passwords (requires config/secret.key)
3. For each server:
   a. Try to connect (credential fallback)
   b. If connected:
      - Collect local users and privileges
      - Collect local groups and members
      - Collect scheduled tasks
   c. If connection failed:
      - Log error
   d. Collect all errors per server
4. Build summary data
5. Generate Excel report with all worksheets
6. Print summary to console
```

### Data Structures

```python
audit_data = {
    "servers": [
        {
            "server": "192.168.1.10",
            "login_success": True,
            "username_used": "administrator",
            "error": None,
            "users": [ ... ],       # from audit_users
            "groups": [ ... ],      # from audit_groups
            "tasks": { ... },       # from audit_tasks
            "errors": [             # per-operation errors
                {"operation": "...", "error": "...", "timestamp": "..."}
            ]
        },
        ...
    ]
}
```

### Command-Line Interface

```bash
python src/audit.py
```

Optional arguments (can be added later):

```
--servers FILE     Path to server list (default: config/servers.txt)
--key FILE         Path to encryption key (default: config/secret.key)
--passwords FILE   Path to encrypted passwords (default: config/passwords.enc)
--output DIR       Output directory (default: output/)
```

### Console Output

During execution, print progress per server:

```
[192.168.1.10] Connecting... OK (administrator)
[192.168.1.10] Collecting users... 5 users found
[192.168.1.10] Collecting groups... 15 groups found
[192.168.1.10] Collecting tasks... 42 tasks (3 user-created)
[192.168.1.11] Connecting... FAILED - All credential combinations failed
---
Audit complete. Report saved to output/Audit_Report_20250101_120000.xlsx
```

### Error Handling

- If connection fails for a server, skip data collection but record the server in Server Login Status and Errors sheets.
- If a data collection step fails (e.g., cannot query users), log the error and continue to the next step.
- Never abort the entire run for a single server failure.
