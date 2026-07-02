"""Password change script for remote Windows servers.

Changes the local Administrator password on each server in the
server list.  The same new password is applied to every server.
Generates a password change report in Excel format.
"""

import argparse
import os
import sys
from datetime import datetime
from getpass import getpass

from crypto import read_encrypted_file
from report import (
    create_workbook,
    save_report,
    write_errors,
    write_password_status,
)
from server_connection import close_session, create_session, load_servers, run_ps

_CHANGE_PASSWORD_SCRIPT = """
$username = 'Administrator'
try {
    $newPassword = ConvertTo-SecureString -String '{password}' -AsPlainText -Force
    $user = Get-LocalUser -Name $username -ErrorAction Stop
    $user | Set-LocalUser -Password $newPassword -ErrorAction Stop
    "SUCCESS"
} catch {
    "FAILED: $($_.Exception.Message)"
}
"""


def prompt_new_password() -> str:
    """Prompt the user for a new Administrator password with confirmation.

    Returns:
        The confirmed new password string.

    Raises:
        SystemExit: If the passwords do not match.
    """
    pwd = getpass("Enter new Administrator password: ")
    confirm = getpass("Confirm new Administrator password: ")
    if pwd != confirm:
        print("Error: Passwords do not match.", file=sys.stderr)
        sys.exit(1)
    return pwd


def change_password(session, new_password: str) -> tuple:
    """Change the local Administrator password on a remote server.

    Args:
        session: A ``winrm.Session`` instance.
        new_password: The new password to set on the Administrator account.

    Returns:
        ``(True, "")`` on success,
        ``(False, error_message)`` on failure.
    """
    escaped = new_password.replace("'", "''")
    script = _CHANGE_PASSWORD_SCRIPT.replace("{password}", escaped)
    success, result = run_ps(session, script)
    if not success:
        return False, result.get("error", "Unknown error")

    stdout = result.strip() if isinstance(result, str) else ""
    if stdout == "SUCCESS":
        return True, ""
    elif stdout.startswith("FAILED:"):
        return False, stdout[7:].strip()
    else:
        return False, stdout or "Unknown result"


def main():
    """Entry point for the password change script."""
    parser = argparse.ArgumentParser(
        description="Change local Administrator password on remote servers"
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
    new_password = prompt_new_password()
    print()

    password_status_rows: list[dict] = []
    error_rows: list[dict] = []
    success_count = 0
    failed_count = 0

    for server in servers:
        sys.stdout.write(f"[{server}] Connecting... ")
        sys.stdout.flush()
        session, result = create_session(server, usernames, passwords)

        if session is not None:
            username = result
            old_status = f"Connected as {username}"
            sys.stdout.write(f"{old_status}... ")
            sys.stdout.flush()

            pwd_success, pwd_msg = change_password(session, new_password)
            close_session(session)

            if pwd_success:
                print("password changed SUCCESS")
                success_count += 1
            else:
                print(f"password changed FAILED - {pwd_msg}")
                failed_count += 1

            password_status_rows.append({
                "server": server,
                "old_connection_status": old_status,
                "new_password_set": pwd_success,
                "error": pwd_msg if not pwd_success else "",
            })
            if not pwd_success:
                error_rows.append({
                    "server": server,
                    "operation": "change_password",
                    "error": pwd_msg,
                    "timestamp": datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                })
        else:
            print(f"Connection FAILED - {result}")
            failed_count += 1
            password_status_rows.append({
                "server": server,
                "old_connection_status": "Connection FAILED",
                "new_password_set": False,
                "error": result,
            })
            error_rows.append({
                "server": server,
                "operation": "connect",
                "error": result,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })

    print("---")
    print(
        f"Password change complete. "
        f"{success_count} succeeded, {failed_count} failed."
    )

    os.makedirs(args.output, exist_ok=True)
    ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Password_Change_{ts_file}.xlsx"
    output_path = os.path.join(args.output, filename)

    wb = create_workbook()
    ws_status = wb.create_sheet("Password Change Status")
    write_password_status(ws_status, password_status_rows)
    if error_rows:
        ws_errors = wb.create_sheet("Errors")
        write_errors(ws_errors, error_rows)
    save_report(wb, output_path)
    report_path = os.path.abspath(output_path)
    print(f"Report saved to {report_path}")


if __name__ == "__main__":
    main()
