# Step 3 — Server List & Connection Module

## Goal

Load server IPs/hostnames from a text file and establish WinRM sessions with automatic credential fallback.

## File: `src/server_connection.py`

### Server List Format (`config/servers.txt`)

```
192.168.1.10
192.168.1.11
srv-dc01.example.com
srv-web02.example.com
```

One hostname or IP per line. Blank lines and `#` comment lines are ignored.

### Functions

```
load_servers(servers_path: str) -> list[str]
    - Read file, strip whitespace, skip blanks and comments.
    - Return list of server addresses.

create_session(server: str, usernames: list[str], passwords: list[str], try_creds: bool = True) -> (Session | None, str)
    - Iterate through username/password combinations.
    - For each combo, attempt to create a winrm.Session.
    - Return (session, username_used) on first success.
    - Return (None, error_message) on total failure.

run_ps(session, script: str) -> (bool, str | dict)
    - Run a PowerShell script block on the session.
    - Return (success, stdout_or_error_dict).
    - Handle WinRM transport errors, timeout, invalid credentials.

close_session(session) -> None
    - Clean up session resources.
```

### WinRM Session Configuration

```python
session = winrm.Session(
    server,
    auth=(username, password),
    transport='credssp',  # fallback to 'ntlm' if credssp unavailable
    server_cert_validation='ignore'
)
```

### Credential Fallback Strategy

```python
usernames = ['administrator', 'admin']
passwords = decrypt_from_file()  # list of possible passwords

for user in usernames:
    for pwd in passwords:
        try:
            session = winrm.Session(server, auth=(user, pwd), ...)
            return session, user
        except Exception:
            continue
return None, "All credential combinations failed"
```

### Error Handling

- Timeouts: catch `winrm.exceptions.WinRMTransportError`.
- Auth failures: catch `requests.exceptions.HTTPError`.
- Network unreachable: catch `requests.exceptions.ConnectionError`.
- Log every failure per server for the Excel report.
