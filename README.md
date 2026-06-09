# Exchange Tester

A standalone Windows tool that replicates Outlook's **"Test E-Mail AutoConfiguration"** dialog (Ctrl + right-click on Outlook's tray icon) — no Outlook, no installation required.

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

---

## What it tests

When **Use SCP** is checked (default), the tool runs in the same order as Outlook on domain-joined machines:

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

---

## Authentication

| Option | Behaviour |
|--------|-----------|
| **Use logged-in user (Windows Auth)** *(default)* | Uses current Windows credentials via NTLM / Kerberos |
| Uncheck + enter password | Sends e-mail address + password as Basic / NTLM credentials |
| **Try Modern Auth (OAuth2)** | Opens a browser window for interactive login (incl. MFA). Requires an Azure AD app — see below. |

---

## Modern Auth (OAuth2) setup

Modern Auth uses **Device Code Flow** — no redirect URI, no browser popup inside the tool. When you click **Test**, a small dialog appears showing a short code (e.g. `ABCD-EFGH`). Open any browser, go to the URL shown (usually `https://microsoft.com/devicelogin`), enter the code, sign in (MFA supported), and the tool receives the token automatically.

The tool tries a built-in Microsoft app ID by default. If your tenant's admin has disabled that app, register your own free Azure AD application (takes ~2 minutes, no admin rights required for single-tenant):

1. Go to **portal.azure.com** → **Azure Active Directory** → **App registrations** → **New registration**
2. **Name:** Exchange Tester (or anything)
3. **Supported account types:** *Accounts in this organizational directory only* (single-tenant)
4. **Redirect URI:** leave blank (not needed for device code flow)
5. Click **Register**
6. Go to **Authentication** → enable **"Allow public client flows"** → Save
7. Copy the **Application (client) ID** and paste it into the **OAuth2 Client ID** field in the tool

No client secret, no API permissions, and no redirect URI are needed.

---

## Tabs

| Tab | Content |
|-----|---------|
| **Results** | Parsed settings grouped by section: User, Account, Protocol (EXCH, EXPR, IMAP, POP3, SMTP, WEB, …) |
| **Log** | Step-by-step protocol log, identical format to Outlook |
| **XML** | Raw AutoDiscover XML response, pretty-printed |

**Right-click** on Log or XML for Copy, Select All, and (XML only) **Save XML As…**.

---

## Options

| Option | Description |
|--------|-------------|
| Try Modern Auth (OAuth2) | Device Code Flow: shows a short code; sign in at microsoft.com/devicelogin in any browser. Optionally enter your own Azure AD app Client ID (see above). |
| Use logged-in user | Toggle between Windows Auth and explicit credentials |
| Ignore certificate errors | Bypasses TLS certificate validation — useful for on-premises Exchange with self-signed certificates |
| Use SCP (domain-joined) | Runs AD Service Connection Point lookup as step 1 (priority, like Outlook on domain-joined machines) |

---

## Files

```
ExchangeTester.bat   — Double-click launcher
ExchangeTester.ps1   — Main script (PowerShell 5.1, WinForms)
```

Both files must be in the same folder.

---

## Planned checks (future milestones)

- EWS health check (`/EWS/Exchange.asmx`, `/HealthCheck.htm`)
- OWA availability (`/owa`)
- ActiveSync (`/Microsoft-Server-ActiveSync`)
- OAB virtual directory (`/OAB`)
- MAPI/HTTP (`/mapi/emsmdb`)
