# Step 4 — Local Users and Privileges

## Goal

Enumerate all local users on each server and classify their privilege level.

## File: `src/audit_users.py`

### PowerShell Command

```powershell
Get-LocalUser | Select-Object Name, FullName, Enabled, Description, SID, PrincipalSource, ObjectClass | ConvertTo-Json
```

For each user, determine privilege by checking group membership:

```powershell
$user = $args[0]
$groups = net localgroup Administrators
if ($groups -contains $user) { "Administrator" }
elseif ((Get-LocalUser $user).Enabled -eq $false) { "Disabled" }
else { "Standard" }
```

Better approach — check membership in known privileged groups:

```powershell
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
        PrivilegeLevel   = if ($memberships -contains 'Administrators') { 'Local Administrator' }
                          elseif ($user.Enabled -eq $false) { 'Disabled' }
                          else { 'Standard User' }
        PrivilegedGroups = $memberships -join ', '
        SID              = $user.SID
        PrincipalSource  = $user.PrincipalSource
        Description      = $user.Description
        IsBuiltIn        = $user.PrincipalSource -eq 'Local' -and $user.Name -in @('Administrator', 'Guest', 'DefaultAccount')
    }
}
$results | ConvertTo-Json
```

### Function

```
collect_users(session) -> list[dict]
    - Run the PowerShell script above.
    - Parse JSON output into list of user dicts.
    - Return list with fields: name, fullname, enabled, privilege_level,
      privileged_groups, sid, principal_source, description, is_builtin.
    - On error, return [] and log error.
```

### Classification Rules

| Condition | Privilege Level |
|-----------|----------------|
| Member of Administrators group | Local Administrator |
| Member of Backup Operators / Power Users / Hyper-V Administrators | Privileged User |
| `Enabled == false` | Disabled User |
| `PrincipalSource == 'Local'` and well-known name | Built-in/System Account |
| Everything else | Standard User |
