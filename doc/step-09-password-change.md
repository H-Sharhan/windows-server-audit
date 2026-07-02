# Step 9 — Password Change Script

## Goal

Change the local Administrator password on each server. No other modifications allowed.

## File: `src/change_password.py`

### PowerShell Command

```powershell
$username = 'Administrator'
$newPassword = ConvertTo-SecureString -String '<NEW_PASSWORD>' -AsPlainText -Force
$user = Get-LocalUser -Name $username
$user | Set-LocalUser -Password $newPassword
if ($?) { "SUCCESS" } else { "FAILED: $($error[0].Exception.Message)" }
```

### Flow

```
1. Load servers from config/servers.txt
2. Load encrypted passwords (key + passwords.enc)
3. Prompt user for the NEW password (hidden input with getpass)
4. Prompt for confirmation (type again)
5. If mismatch, exit with error
6. For each server:
   a. Try to connect with existing credentials
   b. If connected:
      - Run the password change PowerShell command
      - Record success/failure
   c. Log result
7. Generate password change Excel report
8. Print summary to console
```

### Command-Line Interface

```bash
python src/change_password.py
```

### Safety Checks

- The script must ONLY change the local Administrator password.
- No users created, deleted, enabled, or disabled.
- No group memberships modified.
- No scheduled tasks touched.
- No services or system configurations changed.

### Password Change Report

Same Excel format, sheets:

- **Password Change Status**: Server, Old Connection Status, New Password Set, Error
- **Errors**: Server, Operation, Error, Timestamp

### Console Output

```
Enter new Administrator password: ********
Confirm new Administrator password: ********

[192.168.1.10] Connected as administrator... Password changed SUCCESS
[192.168.1.11] Connected as administrator... Password changed SUCCESS
[192.168.1.12] Connection FAILED - Skipped
---
Password change complete. 2 succeeded, 1 failed.
Report saved to output/Password_Change_20250101_120000.xlsx
```

### Important Notes

- The same new password is applied to ALL servers.
- After running, update the encrypted password file with `python src/crypto.py --set` to include the new password for future runs.
