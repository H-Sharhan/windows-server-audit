# Step 7 — Excel Report Generation

## Goal

Generate a professional, formatted Excel workbook with multiple worksheets, auto-filter, colored headers, and column auto-width.

## File: `src/report.py`

### Worksheets

| Sheet | Content |
|-------|---------|
| Summary | Overview counts per server (total users, admins, disabled, groups, tasks, login status) |
| Server Login Status | Per server: server, login success, username used, error message |
| Local Users and Privileges | Per user: server, name, fullname, enabled, privilege level, privileged groups, is built-in |
| Local Groups and Members | Per group member: server, group name, member name, account type, enabled, description |
| Scheduled Tasks | Per task: server, task name, path, run as user, triggers, actions, state, last run, next run, classification |
| Password Change Status | (used by password change script) server, success/fail, old password hint, error |
| Errors | Per error: server, operation, error message, timestamp |

### Formatting

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
```

### Functions

```
create_workbook() -> Workbook
    - Create new workbook, remove default sheet.

_add_header_style(ws, headers: list[str]) -> None
    - Write headers in row 1 with formatting.
    - Add auto-filter on headers.
    - Freeze top row.

_auto_width(ws) -> None
    - Calculate max width per column (cap at 60).
    - Set column widths.

write_summary(ws, data: list[dict]) -> None
    - Server, Total Users, Admins, Privileged, Disabled, Built-in,
      Groups, Tasks, Login Status

write_login_status(ws, data: list[dict]) -> None
    - Server, Success, Username Used, Error

write_users(ws, data: list[dict]) -> None
    - Server, Name, Full Name, Enabled, Privilege Level,
      Privileged Groups, Principal Source, Is Built-in

write_groups(ws, data: list[dict]) -> None
    - Server, Group Name, Member Name, Account Type,
      Enabled, Principal Source

write_tasks(ws, data: list[dict]) -> None
    - Server, Task Name, Path, Run As User, Triggers,
      Actions, State, Last Run, Next Run, Classification

write_password_status(ws, data: list[dict]) -> None
    - Server, Success, Error

write_errors(ws, data: list[dict]) -> None
    - Server, Operation, Error, Timestamp

save_report(workbook, path: str) -> None
    - Save workbook to path with timestamp in filename.
```

### Report Filename

```
output/Audit_Report_YYYYMMDD_HHMMSS.xlsx
```
