# Exchange Tester

A standalone Windows tool that replicates Outlook's **"Test E-Mail AutoConfiguration"** dialog (Ctrl + right-click on Outlook's tray icon) — no Outlook, no installation required — and adds a per-endpoint connectivity tester on top.

---

## Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / 11 / Server 2016+ |
| Runtime | Windows PowerShell 5.1 (built in) |
| .NET | .NET Framework 4.x (built in) |
| DNS SRV lookup | Windows 8.1+ (`Resolve-DnsName`) |

No installation, no admin rights needed for basic AutoDiscover tests.

---

## Usage

**Option A — Double-click launcher:**
```
ExchangeTester.bat
```

**Option B — PowerShell directly:**
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ExchangeTester.ps1
```

The e-mail field is pre-filled with the logged-in user's UPN where it can be detected.

---

## What it tests

When **Use SCP** is checked (default on domain-joined machines), the tool runs in the same order as Outlook:

| Step | Method | URL / Source |
|------|--------|--------------|
| 1 | AD SCP lookup | Active Directory Service Connection Point *(skipped if Use SCP is off)* |
| 2 | HTTPS POST | `https://outlook.office365.com/autodiscover/autodiscover.xml` |
| 3 | HTTPS POST | `https://<domain>/autodiscover/autodiscover.xml` |
| 4 | HTTPS POST | `https://autodiscover.<domain>/autodiscover/autodiscover.xml` |
| 4a | Redirect follow | If step 3 or 4 returns 301/302 |
| 5 | HTTP redirect | `http://autodiscover.<domain>/autodiscover/autodiscover.xml` |
| 6 | DNS SRV | `_autodiscover._tcp.<domain>` |

Each step is logged with HTTP status codes identical to Outlook's protocol log (`GetLastError=0; httpStatus=401.` etc.).

The AutoDiscover request sends `X-MapiHttpCapability: 1`, so Exchange returns the **MAPI/HTTP** (`mapiHttp`) protocol block with the MailStore and AddressBook URLs, exactly like Outlook.

---

## Authentication

Pick the top-level mode, then a sub-option:

| Mode | Sub-option | Behaviour |
|------|-----------|-----------|
| **Modern Auth (OAuth2)** | Public Office app + **Device Code Flow** *(default)* | Shows a short code; you sign in at `microsoft.com/devicelogin` in any browser (MFA supported). No app registration required. |
| **Modern Auth (OAuth2)** | Own "Exchange Tester" app + **Auth Code Flow (PKCE)** | Opens your browser automatically, signs in, and redirects to a local listener — no code to copy. Requires a one-time app registration (see below). |
| **Windows Integrated Auth** | Use logged-in user *(default)* | Current Windows credentials via Kerberos / NTLM. |
| **Windows Integrated Auth** | Uncheck + enter password | Sends e-mail address + password as explicit credentials. |

The tenant ID is detected automatically from the e-mail domain when Auth Code Flow is selected.

---

## Modern Auth: registering your own app

Auth Code Flow needs an Azure AD app registration. The **Register App** button (visible when "Own Exchange Tester app" is selected) does the whole thing for you:

1. Enter your e-mail; the **Tenant ID** is detected automatically.
2. Click **Register App**.
3. Your browser opens — sign in as an **Application Administrator** or **Global Admin**.
4. The tool creates the app registration via Microsoft Graph with:
   - Redirect URI `http://localhost` (loopback — any port matches)
   - API permissions: `EWS.AccessAsUser.All`, `User.Read`, `offline_access`
   - **Tenant-wide admin consent** granted automatically
5. The new **Client ID** is filled in and saved to `ExchangeTester.config` for next time.

After that, the **Test** and **Additional Tests** buttons use Auth Code Flow with that app — no portal steps, no manual consent.

If you prefer not to register an app, just use the default **Device Code Flow**, which needs no setup.

---

## Additional Tests (per-endpoint connectivity)

After a successful AutoDiscover run, the **Additional Tests** button opens a dialog that probes every endpoint found in the response and shows the HTTP result per URL — colour-coded (green 200, orange 3xx, amber 401, red 4xx/5xx/error).

Key points:

- **Follows the selected auth mode.** Modern Auth probes send the OAuth2 Bearer token (reusing the token from the main test for ~55 min, otherwise re-authenticating); WIA probes use the current Windows user or explicit credentials.
- **Every protocol section is shown** (EXCH, EXPR, EXHTTP, mapiHttp, WEB), even when several point at the same URL.
- **MAPI/HTTP** URLs are probed at their base path with a browser-like request, so `/mapi/emsmdb` and `/mapi/nspi` return their friendly "Connectivity Endpoint" page (200) instead of 500.
- **Healthcheck basis test.** For each host, the tool probes `/<vdir>/healthcheck.htm` for the standard Exchange virtual directories — `autodiscover`, `ews`, `oab`, `owa`, `ecp`, `mapi`, `rpc`, `Microsoft-Server-ActiveSync` — plus any extra vdir seen in the AutoDiscover URLs. This gives a consistent reachability baseline across all vdirs, including ones AutoDiscover doesn't return (e.g. OAB).
- **Export.** **Copy CSV** / **Save CSV…** (and a right-click *Copy all as CSV*) export the table as semicolon-separated CSV; right-click a row to **Copy URL**.

---

## Tabs

| Tab | Content |
|-----|---------|
| **Results** | Parsed settings grouped by section: User, Account, Protocol (EXCH, EXPR, EXHTTP, mapiHttp, WEB, …) including MailStore/AddressBook URLs |
| **Log** | Step-by-step protocol log, identical format to Outlook |
| **XML** | Raw AutoDiscover XML response, pretty-printed |

**Right-click** on Log or XML for Copy, Select All, and (XML only) **Save XML As…**.

> Note: the `ASUrl` field is labelled **Availability Service URL** — it is the EWS free/busy endpoint, not the ActiveSync URL.

---

## Options

| Option | Description |
|--------|-------------|
| Modern Auth (OAuth2) | Device Code Flow (no setup) or Auth Code Flow + PKCE (own app via **Register App**) |
| Windows Integrated Auth | Current Windows user, or explicit e-mail + password |
| Ignore certificate errors | Bypasses TLS certificate validation — useful for on-premises Exchange with self-signed certificates |
| Use SCP (domain-joined) | Runs AD Service Connection Point lookup as step 1 (priority, like Outlook on domain-joined machines) |

---

## Files

```
ExchangeTester.bat      — Double-click launcher
ExchangeTester.ps1      — Main script (PowerShell 5.1, WinForms)
ExchangeTester.config   — Saved Client ID (created by Register App)
```

`ExchangeTester.bat` and `ExchangeTester.ps1` must be in the same folder.
