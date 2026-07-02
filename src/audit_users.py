"""Local user enumeration module.

Collects local user accounts and their group memberships from a remote
Windows server via WinRM and classifies each user's privilege level.
"""

import json
import logging

from server_connection import run_ps

logger = logging.getLogger(__name__)

_BUILTIN_ACCOUNTS = {"administrator", "guest", "defaultaccount"}

_USERS_SCRIPT = """
$privilegedGroups = @(
    'Administrators',
    'Backup Operators',
    'Power Users',
    'Remote Desktop Users',
    'Remote Management Users',
    'Hyper-V Administrators'
)
$results = @()
foreach ($user in Get-LocalUser) {
    $memberships = @()
    foreach ($group in $privilegedGroups) {
        $members = net localgroup $group 2>$null
        if ($members -contains $user.Name) {
            $memberships += $group
        }
    }
    $results += [PSCustomObject]@{
        Name             = $user.Name
        FullName         = $user.FullName
        Enabled          = $user.Enabled
        PrivilegedGroups = $memberships -join ', '
        SID              = $user.SID.Value
        PrincipalSource  = $user.PrincipalSource
        Description      = $user.Description
    }
}
$results | ConvertTo-Json
"""


def _classify_privilege(
    name: str,
    enabled: bool,
    privileged_groups: list[str],
    principal_source: str,
) -> str:
    """Classify a user's privilege level.

    Priority order: Administrators → Local Administrator.
    Backup Operators / Power Users / Hyper-V Administrators → Privileged User.
    Disabled account → Disabled User.
    Well-known built-in local account → Built-in/System Account.
    Everything else → Standard User.

    Args:
        name: The user's account name.
        enabled: Whether the account is enabled.
        privileged_groups: List of privileged group names the user belongs to.
        principal_source: The source of the account (e.g. 'Local').

    Returns:
        A string privilege level classification.
    """
    if "Administrators" in privileged_groups:
        return "Local Administrator"
    for group in ("Backup Operators", "Power Users", "Hyper-V Administrators"):
        if group in privileged_groups:
            return "Privileged User"
    if not enabled:
        return "Disabled User"
    if principal_source == "Local" and name.lower() in _BUILTIN_ACCOUNTS:
        return "Built-in/System Account"
    return "Standard User"


def collect_users(session) -> list[dict]:
    """Enumerate local users and their privileges on a remote server.

    Runs a PowerShell script via WinRM that collects all local user
    accounts and checks their membership in well-known privileged groups.
    Each user is then classified into a privilege level.

    Args:
        session: A ``winrm.Session`` instance obtained from
            ``server_connection.create_session``.

    Returns:
        A list of dictionaries, one per user, with keys:
        ``name``, ``fullname``, ``enabled``, ``privilege_level``,
        ``privileged_groups``, ``sid``, ``principal_source``,
        ``description``, ``is_builtin``.
        Returns an empty list if the operation fails on the server.
    """
    success, result = run_ps(session, _USERS_SCRIPT)
    if not success:
        logger.error(
            "Failed to enumerate local users: %s",
            result.get("error", "Unknown error"),
        )
        return []

    if not result:
        return []

    try:
        raw_users = json.loads(result)
    except json.JSONDecodeError as e:
        logger.error("Failed to parse user enumeration JSON: %s", e)
        return []

    if isinstance(raw_users, dict):
        raw_users = [raw_users]
    elif not isinstance(raw_users, list):
        logger.warning(
            "Unexpected JSON structure from user enumeration: %s",
            type(raw_users),
        )
        return []

    users = []
    for raw in raw_users:
        if not isinstance(raw, dict):
            continue

        name = raw.get("Name", "") or ""
        fullname = raw.get("FullName", "") or ""
        enabled = bool(raw.get("Enabled", False))
        sid = raw.get("SID", "") or ""
        principal_source = raw.get("PrincipalSource", "") or ""
        description = raw.get("Description", "") or ""
        groups_str = raw.get("PrivilegedGroups", "") or ""
        privileged_groups = [
            g.strip() for g in groups_str.split(",") if g.strip()
        ]

        is_builtin = (
            principal_source == "Local"
            and name.lower() in _BUILTIN_ACCOUNTS
        )

        privilege_level = _classify_privilege(
            name, enabled, privileged_groups, principal_source
        )

        users.append({
            "name": name,
            "fullname": fullname,
            "enabled": enabled,
            "privilege_level": privilege_level,
            "privileged_groups": groups_str,
            "sid": sid,
            "principal_source": principal_source,
            "description": description,
            "is_builtin": is_builtin,
        })

    return users
