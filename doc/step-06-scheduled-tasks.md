# Step 6 — Scheduled Tasks

## Goal

List all scheduled tasks, classify as system vs user-created, and extract details for user-created tasks.

## File: `src/audit_tasks.py`

### PowerShell Command

```powershell
Get-ScheduledTask | ForEach-Object {
    $task = $_
    $info = Get-ScheduledTaskInfo -TaskPath $task.TaskPath -TaskName $task.TaskName -ErrorAction SilentlyContinue
    $actions = $task.Actions | ForEach-Object { $_.Execute + ' ' + $_.Arguments }
    $triggers = $task.Triggers | ForEach-Object { $_.ToString() }
    $principal = $task.Principal
    [PSCustomObject]@{
        TaskName     = $task.TaskName
        TaskPath     = $task.TaskPath
        RunAsUser    = if ($principal.UserId) { $principal.UserId } else { 'SYSTEM' }
        Triggers     = $triggers -join '; '
        Actions      = $actions -join '; '
        State        = $task.State
        LastRunTime  = if ($info.LastRunTime) { $info.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        NextRunTime  = if ($info.NextRunTime) { $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        TaskPath     = $task.TaskPath
    }
} | ConvertTo-Json
```

### Classification Logic

Tasks stored in `\Microsoft\` or `\Windows\` paths are system tasks.
All others are user-created tasks.

```python
def classify_task(task_path: str) -> str:
    lower = task_path.lower()
    if lower.startswith('\\microsoft\\') or lower.startswith('\\windows\\'):
        return 'System/Default'
    return 'User-Created'
```

### Function

```
collect_tasks(session) -> dict
    - Run PowerShell script.
    - Parse JSON.
    - Classify each task.
    - Return dict:
        {
            "all_tasks": [...],
            "user_tasks": [...],
            "system_tasks": [...]
        }
    - Each task dict: task_name, task_path, run_as_user, triggers,
      actions, state, last_run_time, next_run_time, classification
    - On error, return empty dict and log error.
```
