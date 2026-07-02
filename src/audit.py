"""Read-only audit orchestrator.

Connects to each server in the server list, collects local users, groups,
and scheduled tasks, and generates a formatted Excel report.
"""

import argparse
import os
import sys
from datetime import datetime

from audit_groups import collect_groups
from audit_tasks import collect_tasks
from audit_users import collect_users
from crypto import read_encrypted_file
from report import (
    create_workbook,
    save_report,
    write_errors,
    write_groups,
    write_login_status,
    write_summary,
    write_tasks,
    write_users,
)
from server_connection import close_session, create_session, load_servers


def build_summary_data(audit_data: dict) -> list[dict]:
    """Build per-server summary statistics.

    Args:
        audit_data: The full audit data structure.

    Returns:
        List of dicts suitable for ``write_summary``.
    """
    rows = []
    for s in audit_data["servers"]:
        users = s.get("users", [])
        groups = s.get("groups", [])
        tasks = s.get("tasks", {})
        rows.append({
            "server": s["server"],
            "total_users": len(users),
            "admins": sum(
                1 for u in users
                if u.get("privilege_level") == "Local Administrator"
            ),
            "privileged": sum(
                1 for u in users
                if u.get("privilege_level") == "Privileged User"
            ),
            "disabled": sum(
                1 for u in users
                if u.get("privilege_level") == "Disabled User"
            ),
            "builtin": sum(1 for u in users if u.get("is_builtin")),
            "groups": len(groups),
            "tasks": len(tasks.get("all_tasks", [])),
            "login_status": "Success" if s.get("login_success") else "Failed",
        })
    return rows


def build_login_status_data(audit_data: dict) -> list[dict]:
    """Build login status rows per server.

    Args:
        audit_data: The full audit data structure.

    Returns:
        List of dicts suitable for ``write_login_status``.
    """
    rows = []
    for s in audit_data["servers"]:
        rows.append({
            "server": s["server"],
            "success": s.get("login_success", False),
            "username_used": s.get("username_used", ""),
            "error": s.get("error") or "",
        })
    return rows


def flatten_users(audit_data: dict) -> list[dict]:
    """Flatten user data with server name for Excel output.

    Args:
        audit_data: The full audit data structure.

    Returns:
        List of dicts, one per user per server.
    """
    rows = []
    for s in audit_data["servers"]:
        for u in s.get("users", []):
            row = {"server": s["server"]}
            row.update(u)
            rows.append(row)
    return rows


def flatten_groups(audit_data: dict) -> list[dict]:
    """Flatten group-member data with server name for Excel output.

    Each group produces one row per member.

    Args:
        audit_data: The full audit data structure.

    Returns:
        List of dicts suitable for ``write_groups``.
    """
    rows = []
    for s in audit_data["servers"]:
        for g in s.get("groups", []):
            group_name = g.get("group_name", "")
            for m in g.get("members", []):
                rows.append({
                    "server": s["server"],
                    "group_name": group_name,
                    "member_name": m.get("name", ""),
                    "account_type": m.get("account_type", ""),
                    "enabled": m.get("enabled"),
                    "principal_source": m.get("principal_source", ""),
                })
    return rows


def flatten_tasks(audit_data: dict) -> list[dict]:
    """Flatten task data with server name for Excel output.

    Args:
        audit_data: The full audit data structure.

    Returns:
        List of dicts, one per task per server.
    """
    rows = []
    for s in audit_data["servers"]:
        for t in s.get("tasks", {}).get("all_tasks", []):
            row = {"server": s["server"]}
            row.update(t)
            rows.append(row)
    return rows


def flatten_errors(audit_data: dict) -> list[dict]:
    """Flatten per-operation error data with server name.

    Args:
        audit_data: The full audit data structure.

    Returns:
        List of dicts suitable for ``write_errors``.
    """
    rows = []
    for s in audit_data["servers"]:
        for e in s.get("errors", []):
            rows.append({
                "server": s["server"],
                "operation": e.get("operation", ""),
                "error": e.get("error", ""),
                "timestamp": e.get("timestamp", ""),
            })
    return rows


def generate_report(audit_data: dict, output_path: str) -> str:
    """Build and save the Excel audit report.

    Creates worksheets for Summary, Server Login Status, Local Users and
    Privileges, Local Groups and Members, Scheduled Tasks, and Errors
    (only if errors exist).  Saves the workbook to disk.

    Args:
        audit_data: The full audit data structure.
        output_path: Destination file path for the Excel file.

    Returns:
        The absolute path to the saved report.
    """
    wb = create_workbook()

    ws_summary = wb.create_sheet("Summary")
    write_summary(ws_summary, build_summary_data(audit_data))

    ws_login = wb.create_sheet("Server Login Status")
    write_login_status(ws_login, build_login_status_data(audit_data))

    ws_users = wb.create_sheet("Local Users and Privileges")
    write_users(ws_users, flatten_users(audit_data))

    ws_groups = wb.create_sheet("Local Groups and Members")
    write_groups(ws_groups, flatten_groups(audit_data))

    ws_tasks = wb.create_sheet("Scheduled Tasks")
    write_tasks(ws_tasks, flatten_tasks(audit_data))

    errors_data = flatten_errors(audit_data)
    if errors_data:
        ws_errors = wb.create_sheet("Errors")
        write_errors(ws_errors, errors_data)

    save_report(wb, output_path)
    return os.path.abspath(output_path)


def main():
    """Entry point for the read-only audit."""
    parser = argparse.ArgumentParser(
        description="Windows Server Audit - Read-Only"
    )
    parser.add_argument(
        "--servers",
        default=os.path.join("config", "servers.txt"),
        help="Path to server list (default: config/servers.txt)",
    )
    parser.add_argument(
        "--key",
        default=os.path.join("config", "secret.key"),
        help="Path to encryption key (default: config/secret.key)",
    )
    parser.add_argument(
        "--passwords",
        default=os.path.join("config", "passwords.enc"),
        help="Path to encrypted passwords (default: config/passwords.enc)",
    )
    parser.add_argument(
        "--output",
        default="output",
        help="Output directory for reports (default: output/)",
    )
    args = parser.parse_args()

    servers = load_servers(args.servers)
    if not servers:
        print("No servers found in server list.", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(args.key):
        print(
            f"Error: Encryption key not found at {args.key}",
            file=sys.stderr,
        )
        sys.exit(1)

    if not os.path.exists(args.passwords):
        print(
            f"Error: Encrypted passwords not found at {args.passwords}",
            file=sys.stderr,
        )
        sys.exit(1)

    passwords = read_encrypted_file(args.key, args.passwords)
    usernames = ["administrator", "admin"]

    audit_data: dict = {"servers": []}

    for server in servers:
        server_entry: dict = {
            "server": server,
            "login_success": False,
            "username_used": "",
            "error": None,
            "users": [],
            "groups": [],
            "tasks": {},
            "errors": [],
        }

        sys.stdout.write(f"[{server}] Connecting... ")
        sys.stdout.flush()
        session, result = create_session(server, usernames, passwords)

        if session is not None:
            server_entry["login_success"] = True
            server_entry["username_used"] = result
            print(f"OK ({result})")

            sys.stdout.write(f"[{server}] Collecting users... ")
            sys.stdout.flush()
            try:
                users = collect_users(session)
                server_entry["users"] = users
                print(f"{len(users)} users found")
            except Exception as e:
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                server_entry["errors"].append({
                    "operation": "collect_users",
                    "error": str(e),
                    "timestamp": ts,
                })
                print(f"ERROR - {e}")

            sys.stdout.write(f"[{server}] Collecting groups... ")
            sys.stdout.flush()
            try:
                groups = collect_groups(session)
                server_entry["groups"] = groups
                print(f"{len(groups)} groups found")
            except Exception as e:
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                server_entry["errors"].append({
                    "operation": "collect_groups",
                    "error": str(e),
                    "timestamp": ts,
                })
                print(f"ERROR - {e}")

            sys.stdout.write(f"[{server}] Collecting tasks... ")
            sys.stdout.flush()
            try:
                tasks = collect_tasks(session)
                server_entry["tasks"] = tasks
                all_count = len(tasks.get("all_tasks", []))
                user_count = len(tasks.get("user_tasks", []))
                print(f"{all_count} tasks ({user_count} user-created)")
            except Exception as e:
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                server_entry["errors"].append({
                    "operation": "collect_tasks",
                    "error": str(e),
                    "timestamp": ts,
                })
                print(f"ERROR - {e}")

            close_session(session)
        else:
            server_entry["error"] = result
            print(f"FAILED - {result}")
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            server_entry["errors"].append({
                "operation": "connect",
                "error": result,
                "timestamp": ts,
            })

        audit_data["servers"].append(server_entry)

    print("---")

    os.makedirs(args.output, exist_ok=True)
    ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Audit_Report_{ts_file}.xlsx"
    output_path = os.path.join(args.output, filename)

    report_path = generate_report(audit_data, output_path)
    print(f"Audit complete. Report saved to {report_path}")


if __name__ == "__main__":
    main()
