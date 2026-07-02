import argparse
import json
import os
import sys
from getpass import getpass

from cryptography.fernet import Fernet


def generate_key(key_path: str) -> None:
    """Generate a Fernet encryption key and write it to disk.

    Args:
        key_path: Destination path for the key file.
    """
    key = Fernet.generate_key()
    with open(key_path, "wb") as f:
        f.write(key)
    print(f"Encryption key written to {key_path}")


def load_key(key_path: str) -> bytes:
    """Read a Fernet key from disk.

    Args:
        key_path: Path to the key file.

    Returns:
        The raw key bytes.
    """
    with open(key_path, "rb") as f:
        return f.read()


def encrypt_passwords(passwords: list[str], key: bytes) -> str:
    """Encrypt a list of passwords with Fernet.

    Args:
        passwords: Plain-text passwords to encrypt.
        key: Fernet key bytes.

    Returns:
        JSON string containing the list of encrypted tokens.
    """
    f = Fernet(key)
    encrypted = [f.encrypt(p.encode()).decode() for p in passwords]
    return json.dumps({"passwords": encrypted})


def decrypt_passwords(encrypted_json: str, key: bytes) -> list[str]:
    """Decrypt a JSON string of Fernet-encrypted passwords.

    Args:
        encrypted_json: JSON string as produced by ``encrypt_passwords``.
        key: Fernet key bytes.

    Returns:
        List of decrypted plain-text passwords.
    """
    f = Fernet(key)
    data = json.loads(encrypted_json)
    return [f.decrypt(e.encode()).decode() for e in data["passwords"]]


def write_encrypted_file(
    passwords: list[str], key_path: str, output_path: str
) -> None:
    """Encrypt passwords and write them to a file.

    High-level helper that loads the key, encrypts, and writes the JSON
    output.  Creates parent directories if needed.

    Args:
        passwords: Plain-text passwords to encrypt.
        key_path: Path to the Fernet key file.
        output_path: Destination for the encrypted JSON file.
    """
    key = load_key(key_path)
    encrypted = encrypt_passwords(passwords, key)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w") as f:
        f.write(encrypted)
    print(f"Encrypted passwords written to {output_path}")


def read_encrypted_file(key_path: str, file_path: str) -> list[str]:
    """Read and decrypt the encrypted passwords file.

    Args:
        key_path: Path to the Fernet key file.
        file_path: Path to the encrypted JSON file.

    Returns:
        List of decrypted plain-text passwords.
    """
    key = load_key(key_path)
    with open(file_path, "r") as f:
        encrypted = f.read()
    return decrypt_passwords(encrypted, key)


def main():
    parser = argparse.ArgumentParser(description="Encrypted credential manager")
    parser.add_argument(
        "--init", action="store_true", help="Generate encryption key"
    )
    parser.add_argument(
        "--set",
        action="store_true",
        help="Encrypt and store passwords (interactive)",
    )
    args = parser.parse_args()

    if args.init:
        key_path = os.path.join("config", "secret.key")
        os.makedirs("config", exist_ok=True)
        generate_key(key_path)
    elif args.set:
        key_path = os.path.join("config", "secret.key")
        output_path = os.path.join("config", "passwords.enc")
        if not os.path.exists(key_path):
            print(
                f"Error: Key file not found at {key_path}. "
                f"Run 'python src/crypto.py --init' first.",
                file=sys.stderr,
            )
            sys.exit(1)
        passwords = []
        print(
            "Enter passwords one per line. "
            "Leave blank and press Enter to finish."
        )
        while True:
            pwd = getpass(f"Password {len(passwords) + 1}: ")
            if not pwd:
                break
            passwords.append(pwd)
        if not passwords:
            print("No passwords entered. Exiting.", file=sys.stderr)
            sys.exit(1)
        write_encrypted_file(passwords, key_path, output_path)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
