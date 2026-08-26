# Refresh token runbook

The daily workflow (`.github/workflows/download-excel.yml`) authenticates to
Microsoft Graph with **delegated** permissions: it exchanges a long-lived
refresh token for a short-lived access token on every run.

## Why tokens expire (and how rotation prevents it)

Entra ID issues a **replacement refresh token on every refresh**, and each
individual token expires **90 days after it was issued**. Replaying the same
stored token forever therefore always fails eventually with:

```
AADSTS700082: The refresh token has expired due to inactivity.
The token was issued on <date> and was inactive for 90.00:00:00
```

This happened on 2026-07-29: the stored token had been minted on 2026-04-30 and
hit its 90-day ceiling, even though the workflow ran every day.

`download_delegated.py` now writes each newly issued refresh token back into the
`REFRESH_TOKEN` repository secret (see `github_secrets.py`). Because the workflow
runs daily, the stored token is never more than a day old and the 90-day ceiling
is never reached.

## One-time setup: the rotation token

Rotation needs a credential that can write repository secrets. The built-in
`GITHUB_TOKEN` **cannot** do this, so a separate token is required.

1. Create a **fine-grained personal access token**
   (Settings → Developer settings → Personal access tokens → Fine-grained tokens).
2. Scope it to **this repository only**.
3. Grant **Repository permissions → Secrets: Read and write**.
   No other permission is needed.
4. Set an expiry you will actually renew, and put a calendar reminder on it.
5. Save it as the repository secret **`GH_PAT`**
   (Settings → Secrets and variables → Actions).

If `GH_PAT` is absent the workflow still runs, but it prints a loud
`ACTION REQUIRED` warning and the refresh token will expire within 90 days.

## Recovering from an expired refresh token

Once a refresh token has expired it **cannot** be renewed programmatically. A new
one has to come from an interactive sign-in as the account that owns the files.

There are **two** credentials that expire independently, and the refresh token
cannot be minted while the client secret is dead. Always check the client secret
first — otherwise step 3 below fails with `AADSTS7000222` after the
authorization code has already been burned.

### 0. Check the client secret first

A client secret lasts at most 24 months. If it has expired, every token request
fails with:

```
AADSTS7000222: The provided client secret keys for app '<client-id>' are expired.
```

Recreate it before touching the refresh token:

1. Azure portal → **Microsoft Entra ID** → **App registrations** → this app.
2. **Certificates & secrets** → **Client secrets** → **New client secret**.
3. Add a description and an expiry, then **Add**.
4. Copy the **Value** immediately — it is shown only once. (The *Secret ID* is
   not the secret; the *Value* is.)
5. Update the `CLIENT_SECRET` repository secret with it.
6. Put a calendar reminder on the expiry date. Nothing rotates this
   automatically, and it will silently break the workflow when it lapses.

A new client secret alone does **not** restore access: the refresh token still
has to be re-minted below.

### 1. Build the authorization URL

Replace `<TENANT_ID>` and `<CLIENT_ID>` with the values behind the existing
secrets, then open the URL in a browser and sign in as the file-owning account:

```
https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/authorize
  ?client_id=<CLIENT_ID>
  &response_type=code
  &redirect_uri=http://localhost:8080
  &response_mode=query
  &scope=https://graph.microsoft.com/Files.Read.All%20offline_access
  &prompt=consent
```

`offline_access` is what makes Entra ID return a refresh token — without it you
only get a one-hour access token. Put the URL on a single line (no whitespace).

### 2. Grab the authorization code

After consenting, the browser is redirected to a `http://localhost:8080/?code=...`
URL that will not load. That is expected — copy the `code` parameter out of the
address bar. It is single-use and valid for only a few minutes.

### 3. Exchange the code for a refresh token

```bash
curl -X POST "https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token" \
  -d "client_id=<CLIENT_ID>" \
  -d "client_secret=<CLIENT_SECRET>" \
  -d "code=<CODE_FROM_STEP_2>" \
  -d "redirect_uri=http://localhost:8080" \
  -d "grant_type=authorization_code" \
  -d "scope=https://graph.microsoft.com/Files.Read.All offline_access"
```

Take the `refresh_token` field from the JSON response.

### 4. Store it and verify

1. Update the `REFRESH_TOKEN` repository secret with the new value.
2. Run the workflow manually (Actions → *Download Excel via Delegated Auth* →
   **Run workflow**).
3. Confirm the log shows `[+] Secret 'REFRESH_TOKEN' updated via GitHub API`.
   That line means rotation is working and the token will stay fresh.

## Handling the token safely

The refresh token grants read access to the account's files. Treat it like a
password — and note that **this repository is public**, so its Actions logs,
commits and files are all world-readable:

- **Never paste it into a log, an issue, a PR, or a commit.** Earlier revisions
  of `download_delegated.py` printed the full token into the Actions log on every
  successful run. Because the repository is public, those logs were readable by
  anyone; every token from before this change must be considered compromised.
- **Do not put credentials in shared documents.** Renewal guides circulated as
  Word/PDF files have carried the live `client_secret` inline. Reference the
  secret by name and keep the value only in Azure and GitHub Secrets.
- `.gitignore` blocks `NEW_REFRESH_TOKEN.txt` and similar files from being
  committed.
- Redo the re-consent flow above to invalidate a token you believe has leaked.

## Other failure codes

| Code | Meaning | Fix |
| --- | --- | --- |
| `AADSTS700082` | Refresh token expired | Re-consent (above) |
| `AADSTS50173` | Account password changed | Re-consent (above) |
| `AADSTS7000222` | Client secret **expired** | Recreate it — step 0 above |
| `AADSTS7000215` | Client secret wrong/invalid | Check `CLIENT_SECRET` — step 0 above |
| `AADSTS50011` | Redirect URI mismatch | Must match the app registration exactly (`http://localhost:8080`) |
| `AADSTS65001` | Consent not granted | Re-run authorize URL with `prompt=consent` |
