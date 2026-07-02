# Step 2 — Encrypted Password File Module

## Goal

Provide a secure way to store and retrieve credentials without plain-text passwords on disk.

## Design

- Use **Fernet** symmetric encryption from the `cryptography` library.
- A key file (`config/secret.key`) is generated once and stored securely.
- The password file (`config/passwords.enc`) stores a JSON dict of encrypted passwords.
- Two CLI modes: `--init` (generate key) and `--set` (encrypt and store passwords).
- Password retrieval is done in-memory at runtime.

## File: `src/crypto.py`

### Functions

```
generate_key(key_path: str) -> None
    - Generate a Fernet key and write to key_path.

load_key(key_path: str) -> bytes
    - Read and return the Fernet key.

encrypt_passwords(passwords: list[str], key: bytes) -> str
    - Encrypt each password with Fernet.
    - Return JSON string: {"passwords": [enc1, enc2, ...]}.

decrypt_passwords(encrypted_json: str, key: bytes) -> list[str]
    - Decrypt and return the original password list.

write_encrypted_file(passwords: list[str], key_path: str, output_path: str) -> None
    - High-level: load key, encrypt, write to file.

read_encrypted_file(key_path: str, file_path: str) -> list[str]
    - High-level: load key, read file, decrypt, return list.
```

### CLI Interface

```bash
# Generate encryption key
python src/crypto.py --init

# Encrypt and store passwords (interactive, hidden input)
python src/crypto.py --set
```

When `--set` is used:
- Prompt for passwords one by one (hidden input via getpass).
- Keep prompting until user enters a blank line.
- Encrypt the list and write to `config/passwords.enc`.

### Security Considerations

- `secret.key` must be kept secure and never committed to version control.
- The key file and encrypted password file must be distributed together to the automation runner.
- In a production environment, consider using a secrets manager (Azure Key Vault, HashiCorp Vault) instead.
