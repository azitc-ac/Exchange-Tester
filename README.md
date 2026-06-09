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

The tool runs the same AutoDiscover discovery sequence as Outlook, in order:

| Step | Method | URL / Source |
|------|--------|--------------|
| 1 | HTTPS POST | `https://outlook.office365.com/autodiscover/autodiscover.xml` |
| 2 | HTTPS POST | `https://<domain>/autodiscover/autodiscover.xml` |
| 3 | HTTPS POST | `https://autodiscover.<domain>/autodiscover/autodiscover.xml` |
| 3a | Redirect follow | If step 2 or 3 returns 301/302 |
| 4 | AD SCP lookup | Active Directory Service Connection Point (domain-joined machines) |
| 5 | HTTP redirect | `http://autodiscover.<domain>/autodiscover/autodiscover.xml` |
| 6 | DNS SRV | `_autodiscover._tcp.<domain>` |

Each step is logged with HTTP status codes identical to Outlook's protocol log (`GetLastError=0; httpStatus=401.` etc.).

---

## Authentication

| Option | Behaviour |
|--------|-----------|
| **Use logged-in user (Windows Auth)** *(default)* | Uses current Windows credentials via NTLM / Kerberos |
| Uncheck + enter password | Sends e-mail address + password as Basic / NTLM credentials |

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
| Use AutoDiscover | Always enabled (placeholder for future Guessmart support) |
| Use logged-in user | Toggle between Windows Auth and explicit credentials |
| Ignore certificate errors | Bypasses TLS certificate validation — useful for on-premises Exchange with self-signed certificates |

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
