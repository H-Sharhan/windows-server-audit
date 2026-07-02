"""Scheduled task enumeration module.

Collects all scheduled tasks from a remote Windows server via WinRM,
classifying each as system/default or user-created, and extracting
detailed information for user-created tasks.
"""

import json
import logging

from server_connection import run_ps

logger = logging.getLogger(__name__)

_TASKS_SCRIPT = """
Get-ScheduledTask | ForEach-Object {
    $task = $_
    $info = Get-ScheduledTaskInfo -TaskPath $task.TaskPath -TaskName $task.TaskName -ErrorAction SilentlyContinue
    $actions = if ($task.Actions) { $task.Actions | ForEach-Object { $_.Execute + ' ' + $_.Arguments } } else { @() }
    $triggers = if ($task.Triggers) { $task.Triggers | ForEach-Object { $_.ToString() } } else { @() }
    $principal = $task.Principal
    [PSCustomObject]@{
        TaskName     = $task.TaskName
        TaskPath     = $task.TaskPath
        RunAsUser    = if ($principal -and $principal.UserId) { $principal.UserId } else { 'SYSTEM' }
        Triggers     = ($triggers | Where-Object { $_ }) -join '; '
        Actions      = ($actions | Where-Object { $_ }) -join '; '
        State        = $task.State
        LastRunTime  = if ($info -and $info.LastRunTime) { $info.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        NextRunTime  = if ($info -and $info.NextRunTime) { $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
    }
} | ConvertTo-Json -Depth 3
"""


def classify_task(task_path: str) -> str:
    """Classify a scheduled task as system/default or user-created.

    Tasks stored under ``\\Microsoft\\`` or ``\\Windows\\`` paths are
    considered system/default tasks. All others are user-created.

    Args:
        task_path: The task path
            (e.g. ``\\Microsoft\\Windows\\TaskScheduler\\Task``).

    Returns:
        ``'System/Default'`` or ``'User-Created'``.
    """
    lower = task_path.lower()
    if lower.startswith("\\microsoft\\") or lower.startswith("\\windows\\"):
        return "System/Default"
    return "User-Created"


def collect_tasks(session) -> dict:
    """Enumerate all scheduled tasks on a remote server.

    Runs a PowerShell script via WinRM that collects every scheduled task,
    including its actions, triggers, run-as user, state, and last/next run
    times. Each task is classified as system/default or user-created.

    Args:
        session: A ``winrm.Session`` instance obtained from
            ``server_connection.create_session``.

    Returns:
        A dictionary with three keys:

        - ``all_tasks``: list of all task dicts (sorted: user-created
          first, then system; alphabetical by name within each group).
        - ``user_tasks``: list of user-created task dicts.
        - ``system_tasks``: list of system/default task dicts.

        Each task dict contains: ``task_name``, ``task_path``,
        ``run_as_user``, ``triggers``, ``actions``, ``state``,
        ``last_run_time``, ``next_run_time``, ``classification``.

        Returns an empty dict if the operation fails on the server.
    """
    success, result = run_ps(session, _TASKS_SCRIPT)
    if not success:
        logger.error(
            "Failed to enumerate scheduled tasks: %s",
            result.get("error", "Unknown error"),
        )
        return {}

    if not result:
        return {}

    try:
        raw_tasks = json.loads(result)
    except json.JSONDecodeError as e:
        logger.error("Failed to parse scheduled task JSON: %s", e)
        return {}

    if isinstance(raw_tasks, dict):
        raw_tasks = [raw_tasks]
    elif not isinstance(raw_tasks, list):
        logger.warning(
            "Unexpected JSON structure from task enumeration: %s",
            type(raw_tasks),
        )
        return {}

    all_tasks = []
    for raw in raw_tasks:
        if not isinstance(raw, dict):
            continue

        task_name = raw.get("TaskName", "") or ""
        task_path = raw.get("TaskPath", "") or ""
        run_as_user = raw.get("RunAsUser", "") or ""
        triggers = raw.get("Triggers", "") or ""
        actions = raw.get("Actions", "") or ""
        state = raw.get("State", "") or ""
        last_run_time = raw.get("LastRunTime", "") or ""
        next_run_time = raw.get("NextRunTime", "") or ""
        classification = classify_task(task_path)

        all_tasks.append({
            "task_name": task_name,
            "task_path": task_path,
            "run_as_user": run_as_user,
            "triggers": triggers,
            "actions": actions,
            "state": state,
            "last_run_time": last_run_time,
            "next_run_time": next_run_time,
            "classification": classification,
        })

    all_tasks.sort(key=lambda t: (
        0 if t["classification"] == "User-Created" else 1,
        t["task_name"].lower(),
    ))

    user_tasks = [
        t for t in all_tasks if t["classification"] == "User-Created"
    ]
    system_tasks = [
        t for t in all_tasks if t["classification"] == "System/Default"
    ]

    return {
        "all_tasks": all_tasks,
        "user_tasks": user_tasks,
        "system_tasks": system_tasks,
    }
