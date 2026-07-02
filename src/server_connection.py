"""Server list loader and WinRM session factory.

Provides functions to load server addresses from a text file, create WinRM
sessions with credential and transport fallback, execute PowerShell scripts,
and clean up session resources.
"""

import logging

import requests.exceptions
import winrm
from winrm.exceptions import (
    AuthenticationError,
    WinRMTransportError,
)

logger = logging.getLogger(__name__)


def load_servers(servers_path: str) -> list[str]:
    """Read server addresses from a text file.

    One hostname or IP per line. Blank lines and lines starting with ``#``
    are ignored.

    Args:
        servers_path: Path to the servers text file.

    Returns:
        List of server address strings.

    Raises:
        FileNotFoundError: If the file does not exist.
    """
    servers = []
    with open(servers_path, "r") as f:
        for line in f:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            servers.append(stripped)
    return servers


def create_session(
    server: str,
    usernames: list[str],
    passwords: list[str],
    try_creds: bool = True,
) -> tuple:
    """Create a WinRM session to a server with credential fallback.

    Tries every combination of usernames and passwords, first with CredSSP
    transport then NTLM as a fallback.  On ``AuthenticationError`` the
    inner transport loop is skipped for that password (wrong creds); all
    other errors continue searching.

    Args:
        server: Hostname or IP of the target server.
        usernames: Ordered list of usernames to try (e.g.
            ``['administrator', 'admin']``).
        passwords: List of candidate passwords.
        try_creds: If ``False`` only the first username and first password
            are attempted (no fallback).  Default ``True``.

    Returns:
        ``(session, username)`` on success, ``(None, error_message)`` on
        failure.
    """
    transports = ["credssp", "ntlm"]

    if not try_creds:
        usernames = usernames[:1]
        passwords = passwords[:1]

    errors: list[str] = []
    for user in usernames:
        for pwd in passwords:
            for transport in transports:
                try:
                    session = winrm.Session(
                        server,
                        auth=(user, pwd),
                        transport=transport,
                        server_cert_validation="ignore",
                    )
                    rs = session.run_ps("Write-Output 'ping'")
                    if rs.status_code == 0:
                        logger.info(
                            "Connected to %s as %s via %s", server, user, transport
                        )
                        return session, user
                except AuthenticationError:
                    errors.append(f"Authentication failed: {user} via {transport}")
                    break
                except WinRMTransportError as e:
                    errors.append(
                        f"WinRMTransportError for {user} via {transport}: {e}"
                    )
                except requests.exceptions.ConnectionError as e:
                    errors.append(f"ConnectionError for {server}: {e}")
                except requests.exceptions.Timeout as e:
                    errors.append(f"Timeout for {server}: {e}")
                except Exception as e:
                    errors.append(
                        f"Unexpected error for {user} via {transport}: {e}"
                    )

    return None, "; ".join(errors) or "All credential combinations failed"


def run_ps(session, script: str) -> tuple:
    """Run a PowerShell script on a WinRM session.

    Args:
        session: A ``winrm.Session`` instance.
        script: PowerShell script text to execute.

    Returns:
        ``(True, stdout_string)`` on success,
        ``(False, {"error": ..., "status_code": ...})`` on failure.
    """
    try:
        rs = session.run_ps(script)
        if rs.status_code == 0:
            return True, rs.std_out.decode("utf-8", errors="replace").strip()
        else:
            err_text = rs.std_err.decode("utf-8", errors="replace").strip()
            return False, {"error": err_text, "status_code": rs.status_code}
    except WinRMTransportError as e:
        return False, {
            "error": f"WinRM transport error: {e}",
            "status_code": e.code if hasattr(e, "code") else 0,
        }
    except AuthenticationError as e:
        return False, {
            "error": f"Authentication error: {e}",
            "status_code": 401,
        }
    except requests.exceptions.Timeout as e:
        return False, {
            "error": f"Operation timed out: {e}",
            "status_code": 0,
        }
    except requests.exceptions.ConnectionError as e:
        return False, {
            "error": f"Connection lost: {e}",
            "status_code": 0,
        }
    except Exception as e:
        return False, {
            "error": f"Unexpected error: {e}",
            "status_code": 0,
        }


def close_session(session) -> None:
    """Release resources held by a WinRM session.

    Safe to call multiple times or with ``None``.

    Args:
        session: A ``winrm.Session`` instance or ``None``.
    """
    if session is not None:
        try:
            session.protocol.transport.close_session()
        except Exception:
            pass
