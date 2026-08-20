"""Write GitHub Actions secrets from inside a workflow run.

Used by download_delegated.py to persist the rotated Microsoft refresh token
back into the REFRESH_TOKEN secret. Entra ID issues a replacement refresh token
on every refresh and caps each individual token at 90 days from issuance, so a
token that is never written back will always die (AADSTS700082) roughly 90 days
after it was first minted — no matter how often the workflow runs.

Requires a token with 'Secrets: write' permission on the repository, exposed as
GH_PAT. The built-in GITHUB_TOKEN cannot write secrets.
"""

import base64
import os

import requests

API_ROOT = 'https://api.github.com'


def rotation_is_configured():
    """True if we have everything needed to write a secret back."""
    return bool(os.environ.get('GH_PAT') and os.environ.get('GITHUB_REPOSITORY'))


def _api_headers(pat):
    return {
        'Authorization': f'Bearer {pat}',
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
    }


def _encrypt(public_key, secret_value):
    """Seal a secret with the repository's public key (libsodium sealed box)."""
    from nacl import encoding, public

    key = public.PublicKey(public_key.encode('utf-8'), encoding.Base64Encoder())
    sealed = public.SealedBox(key).encrypt(secret_value.encode('utf-8'))
    return base64.b64encode(sealed).decode('utf-8')


def update_secret(name, value):
    """Encrypt and PUT a repository Actions secret.

    Returns True on success. Never raises and never logs the secret value —
    callers treat failure as a loud warning, not a fatal error, so that a
    rotation problem does not discard an otherwise good data download.
    """
    pat = os.environ.get('GH_PAT')
    repo = os.environ.get('GITHUB_REPOSITORY')

    if not pat or not repo:
        print("[!] Secret rotation not configured (GH_PAT / GITHUB_REPOSITORY missing).")
        return False

    headers = _api_headers(pat)

    try:
        key_response = requests.get(
            f"{API_ROOT}/repos/{repo}/actions/secrets/public-key",
            headers=headers,
            timeout=30
        )
        if key_response.status_code != 200:
            print(f"[X] Could not fetch repo public key: {key_response.status_code}")
            print(f"[X] {key_response.text[:300]}")
            return False

        key_data = key_response.json()
        encrypted_value = _encrypt(key_data['key'], value)

        put_response = requests.put(
            f"{API_ROOT}/repos/{repo}/actions/secrets/{name}",
            headers=headers,
            json={'encrypted_value': encrypted_value, 'key_id': key_data['key_id']},
            timeout=30
        )

        # 201 = secret created, 204 = existing secret updated.
        if put_response.status_code in (201, 204):
            print(f"[+] Secret '{name}' updated via GitHub API (repo: {repo})")
            return True

        print(f"[X] Failed to update secret '{name}': {put_response.status_code}")
        print(f"[X] {put_response.text[:300]}")
        if put_response.status_code in (403, 404):
            print("[X] Check that GH_PAT has 'Secrets: write' on this repository.")
        return False

    except ImportError:
        print("[X] PyNaCl is not installed — cannot encrypt the secret.")
        print("[X] Add 'pynacl' to the workflow's pip install step.")
        return False
    except Exception as e:
        print(f"[X] Unexpected error updating secret '{name}': {e}")
        return False
