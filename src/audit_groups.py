"""Local group enumeration module.

Collects all local groups and their members from a remote Windows server
via WinRM, classifying each member by account type and checking the
enabled/disabled status for user-type members.
"""

import json
import logging

from server_connection import run_ps

logger = logging.getLogger(__name__)

_PRIORITIZED_GROUPS = frozenset({
    "Administrators",
    "Remote Desktop Users",
    "Backup Operators",
    "Power Users",
    "Remote Management Users",
    "Hyper-V Administrators",
})

_WELL_KNOWN_NAME_PREFIXES = (
    "NT AUTHORITY\\",
    "BUILTIN\\",
    "APPLICATION PACKAGE AUTHORITY\\",
)

_GROUPS_SCRIPT = """
$groups = Get-LocalGroup | Select-Object Name, Description, @{N='SID';E={$_.SID.Value}}
$results = @()
foreach ($group in $groups) {
    $members = @()
    try {
        $groupMembers = Get-LocalGroupMember -Group $group.Name -ErrorAction Stop
        foreach ($m in $groupMembers) {
            $enabled = $null
            if ($m.ObjectClass -eq 'User') {
                try {
                    $adObj = [ADSI]"WinNT://$env:COMPUTERNAME/$($m.Name.Split('\')[-1])"
                    $enabled = -not [bool]$adObj.AccountDisabled
                } catch {
                    $enabled = $null
                }
            }
            $members += [PSCustomObject]@{
                Name        = $m.Name
                SID         = $m.SID.Value
                ObjectClass = $m.ObjectClass
                PrincipalSource = $m.PrincipalSource
                Enabled     = $enabled
            }
        }
    } catch {
        $members = @()
    }
    $results += [PSCustomObject]@{
        GroupName   = $group.Name
        Description = $group.Description
        SID         = $group.SID
        Members     = $members
    }
}
$results | ConvertTo-Json -Depth 3
"""


def _is_well_known_name(name: str) -> bool:
    """Check if a member name indicates a well-known built-in account.

    Args:
        name: The member name (e.g. ``NT AUTHORITY\\NETWORK SERVICE``).

    Returns:
        True if the name starts with a well-known authority prefix.
    """
    return name.upper().startswith(_WELL_KNOWN_NAME_PREFIXES)


def _classify_account_type(
    object_class: str,
    principal_source: str,
    name: str,
    sid: str,
) -> str:
    """Classify a group member's account type.

    Classification follows the rules in the project specification:

    - ``User`` + ``Local`` → ``Local User``
    - ``User`` + ``ActiveDirectory`` → ``Domain User``
    - ``User`` + ``MicrosoftAccount`` → ``Microsoft Account``
    - ``Group`` + ``Local`` → ``Local Group``
    - ``Group`` + ``ActiveDirectory`` → ``Domain Group``
    - Well-known names (NT AUTHORITY\\*, BUILTIN\\*, etc.) → ``Built-in Account``
    - Everything else → ``Built-in Account``

    Args:
        object_class: ``'User'`` or ``'Group'``.
        principal_source: e.g. ``'Local'``, ``'ActiveDirectory'``,
            ``'MicrosoftAccount'``.
        name: The member's full name (e.g. ``'DOMAIN\\User'``).
        sid: The member's security identifier string.

    Returns:
        One of ``'Local User'``, ``'Domain User'``, ``'Microsoft Account'``,
        ``'Local Group'``, ``'Domain Group'``, ``'Built-in Account'``.
    """
    if _is_well_known_name(name):
        return "Built-in Account"

    if object_class == "User":
        if principal_source == "Local":
            return "Local User"
        elif principal_source == "ActiveDirectory":
            return "Domain User"
        elif principal_source == "MicrosoftAccount":
            return "Microsoft Account"
    elif object_class == "Group":
        if principal_source == "Local":
            return "Local Group"
        elif principal_source == "ActiveDirectory":
            return "Domain Group"

    return "Built-in Account"


def _is_prioritized_group(group_name: str) -> bool:
    """Check if the group is in the prioritized reporting list.

    Prioritized groups include Administrators, Remote Desktop Users,
    Backup Operators, Power Users, Remote Management Users, and
    Hyper-V Administrators.

    Args:
        group_name: The local group name.

    Returns:
        True if the group is in the prioritized list.
    """
    return group_name in _PRIORITIZED_GROUPS


def collect_groups(session) -> list[dict]:
    """Enumerate all local groups and their members on a remote server.

    Runs a PowerShell script via WinRM that collects every local group,
    their descriptions, SIDs, and members.  For each member the account
    type is classified and the enabled/disabled status is checked for
    user-type members.

    Args:
        session: A ``winrm.Session`` instance obtained from
            ``server_connection.create_session``.

    Returns:
        A list of dictionaries, one per group, with keys:
        ``group_name``, ``description``, ``sid``, ``is_prioritized``,
        ``members``.

        Each member dict has keys: ``name``, ``sid``, ``object_class``,
        ``principal_source``, ``account_type``, ``enabled``.

        Returns an empty list if the operation fails on the server.
    """
    success, result = run_ps(session, _GROUPS_SCRIPT)
    if not success:
        logger.error(
            "Failed to enumerate local groups: %s",
            result.get("error", "Unknown error"),
        )
        return []

    if not result:
        return []

    try:
        raw_groups = json.loads(result)
    except json.JSONDecodeError as e:
        logger.error("Failed to parse group enumeration JSON: %s", e)
        return []

    if isinstance(raw_groups, dict):
        raw_groups = [raw_groups]
    elif not isinstance(raw_groups, list):
        logger.warning(
            "Unexpected JSON structure from group enumeration: %s",
            type(raw_groups),
        )
        return []

    groups = []
    for raw in raw_groups:
        if not isinstance(raw, dict):
            continue

        group_name = raw.get("GroupName", "") or ""
        description = raw.get("Description", "") or ""
        sid = raw.get("SID", "") or ""
        raw_members = raw.get("Members", [])

        if not isinstance(raw_members, list):
            raw_members = []

        members = []
        for m in raw_members:
            if not isinstance(m, dict):
                continue

            m_name = m.get("Name", "") or ""
            m_sid = m.get("SID", "") or ""
            m_object_class = m.get("ObjectClass", "") or ""
            m_principal_source = m.get("PrincipalSource", "") or ""
            m_enabled = m.get("Enabled")

            account_type = _classify_account_type(
                m_object_class,
                m_principal_source,
                m_name,
                m_sid,
            )

            members.append({
                "name": m_name,
                "sid": m_sid,
                "object_class": m_object_class,
                "principal_source": m_principal_source,
                "account_type": account_type,
                "enabled": m_enabled,
            })

        groups.append({
            "group_name": group_name,
            "description": description,
            "sid": sid,
            "is_prioritized": _is_prioritized_group(group_name),
            "members": members,
        })

    return groups
