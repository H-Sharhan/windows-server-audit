# Step 5 — Local Groups and Membership

## Goal

Enumerate all local groups and their members with account type classification.

## File: `src/audit_groups.py`

### PowerShell Command

```powershell
$groups = Get-LocalGroup | Select-Object Name, Description, SID
$results = @()
foreach ($group in $groups) {
    $members = @()
    try {
        $groupMembers = Get-LocalGroupMember -Group $group.Name -ErrorAction Stop
        foreach ($m in $groupMembers) {
            $members += [PSCustomObject]@{
                Name        = $m.Name
                SID         = $m.SID
                ObjectClass = $m.ObjectClass  # 'User' or 'Group'
                PrincipalSource = $m.PrincipalSource  # 'Local' or 'ActiveDirectory' or 'MicrosoftAccount'
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
```

### Account Type Classification

In Python, classify each member based on `ObjectClass` and `PrincipalSource`:

| ObjectClass | PrincipalSource | Account Type |
|-------------|-----------------|-------------|
| User | Local | Local User |
| User | ActiveDirectory | Domain User |
| User | MicrosoftAccount | Microsoft Account |
| Group | Local | Local Group |
| Group | ActiveDirectory | Domain Group |
| Group | (any) | Built-in Account (if well-known SID) |
| (any) | (any) | Built-in Account (if name is built-in like `NT AUTHORITY\NETWORK SERVICE`) |

### Enabled/Disabled Status for Members

For **User** members, we can check with:

```powershell
$adObj = [ADSI]"WinNT://$env:COMPUTERNAME/$($m.Name.Split('\')[-1])"
$adObj.AccountDisabled  # returns True/False or null for groups
```

### Prioritized Groups (must be highlighted in report)

- Administrators
- Remote Desktop Users
- Backup Operators
- Power Users
- Remote Management Users
- Hyper-V Administrators

### Function

```
collect_groups(session) -> list[dict]
    - Run PowerShell script, parse JSON.
    - Return list of group dicts, each with:
        group_name, description, sid, members[]
    - Each member: name, sid, object_class, principal_source,
      account_type, enabled
    - On error, return [] and log error.
```
