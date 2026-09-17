# Exchange Tester

A standalone Windows tool for diagnosing Exchange connectivity — no Outlook, no installation required. It opens on a **start screen** offering four tools:

1. **E-Mail AutoConfiguration** — replicates Outlook's "Test E-Mail AutoConfiguration" (Ctrl + right-click the Outlook tray icon), plus a per-endpoint connectivity tester.
2. **Hybrid Deployment (MRS Proxy)** — a step-by-step diagnostic of the on-premises mailbox-migration endpoint used by hybrid mailbox moves.
3. **Hybrid Connectivity** — a broad reachability sweep of on-prem + Exchange Online coexistence endpoints (AutoDiscover, OAuth, Free/Busy, vdir health).
4. **Free/Busy (Cross-Premises)** — an authenticated availability test in both directions with real mailboxes (On-Prem ↔ Exchange Online).

Closing any tool returns to the start screen; closing the start screen exits.

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

## 1. E-Mail AutoConfiguration

### What it tests

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

Each step is logged with HTTP status codes identical to Outlook's protocol log (`GetLastError=0; httpStatus=401.` etc.). The request sends `X-MapiHttpCapability: 1`, so Exchange returns the **MAPI/HTTP** (`mapiHttp`) protocol block with the MailStore and AddressBook URLs, just like Outlook.

### Authentication

Pick the top-level mode, then a sub-option:

| Mode | Sub-option | Behaviour |
|------|-----------|-----------|
| **Modern Auth (OAuth2)** | Public Office app + **Device Code Flow** *(default)* | Shows a short code; you sign in at `microsoft.com/devicelogin` in any browser (MFA supported). No app registration required. |
| **Modern Auth (OAuth2)** | Own "Exchange Tester" app + **Auth Code Flow (PKCE)** | Opens your browser automatically and redirects to a local listener — no code to copy. Requires a one-time app registration (see below). |
| **Windows Integrated Auth** | Use logged-in user *(default)* | Current Windows credentials via Kerberos / NTLM. |
| **Windows Integrated Auth** | Uncheck + enter password | Sends e-mail address + password as explicit credentials. |

The tenant ID is detected automatically from the e-mail domain. Device Code Flow resolves a concrete tenant (or `/organizations`) and requests the `EWS.AccessAsUser.All` scope, avoiding the `AADSTS50059` "no tenant" error.

### Registering your own app (Auth Code Flow)

The **Register App** button (shown when "Own Exchange Tester app" is selected) does everything via Microsoft Graph:

1. Enter your e-mail; the **Tenant ID** is detected automatically.
2. Click **Register App** and sign in as an **Application Administrator** or **Global Admin**.
3. The app is created with redirect URI `http://localhost`, API permissions `EWS.AccessAsUser.All`, `User.Read`, `offline_access`, and **tenant-wide admin consent** granted automatically.
4. The new **Client ID** is filled in and saved to `ExchangeTester.config`.

No portal steps and no manual consent afterwards. Prefer zero setup? Use the default Device Code Flow.

### Additional Tests (per-endpoint connectivity)

After a successful run, **Additional Tests** probes every endpoint in the response and shows a per-URL result, colour-coded by a reachability **verdict** (not pass/fail — a service answering with an auth challenge or redirect is reachable/healthy):

| Verdict | Meaning |
|---|---|
| OK (200) | Reachable |
| Reachable (redirect) | 3xx |
| Reachable — auth required / needs POST | 401 / 405 |
| Not present | 404 |
| Server error | 5xx |
| Unreachable | connection failed |

- **Follows the selected auth mode** (Modern Auth Bearer token, reused for ~55 min; or Windows / explicit credentials).
- Every protocol section is shown (EXCH, EXPR, EXHTTP, mapiHttp, WEB).
- **MAPI/HTTP** URLs are probed at their base path with a browser-like request so `/mapi/emsmdb` and `/mapi/nspi` return their friendly "Connectivity Endpoint" page instead of 500.
- **Healthcheck basis test** — `/<vdir>/healthcheck.htm` for the standard vdirs (`autodiscover`, `ews`, `oab`, `owa`, `ecp`, `mapi`, `rpc`, `Microsoft-Server-ActiveSync`) plus any vdir seen in the response.
- **Copy CSV** / **Save CSV…** export the table (semicolon-separated); right-click a row to copy the URL.

### Tabs

| Tab | Content |
|-----|---------|
| **Results** | Parsed settings grouped by section: User, Account, Protocol (EXCH, EXPR, EXHTTP, mapiHttp, WEB…), MailStore/AddressBook URLs, **Alternative Mailboxes** (shared/delegate) and **Public Folder Information** |
| **Log** | Step-by-step protocol log, identical format to Outlook |
| **XML** | Raw AutoDiscover XML response, pretty-printed |

Right-click Log or XML for Copy, Select All, and (XML) **Save XML As…**.

> The `ASUrl` field is labelled **Availability Service URL** — it is the EWS free/busy endpoint, not ActiveSync. The real ActiveSync URL is not part of the Outlook AutoDiscover schema.

---

## 2. Hybrid Deployment (MRS Proxy)

Diagnoses the on-premises **migration endpoint** (`/EWS/mrsproxy.svc`) that Exchange Online uses for hybrid mailbox moves. Enter an on-prem mailbox (the endpoint is discovered via AutoDiscover) or the MRS endpoint FQDN directly, plus the on-prem migration credentials, and run a step-by-step check:

**Endpoint discovery → DNS resolution → TCP 443 → TLS handshake → Certificate trust / name match / validity → MRS Proxy endpoint → MRS Proxy authentication → MRS Proxy SOAP.**

Each step reports **OK / WARN / FAIL / INFO** with details. The window is resizable; **hover** a row for a tooltip and **double-click** it to read the full text (also copied to the clipboard). A Log tab shows the raw sequence.

**How the auth check reads results:**

- The endpoint probe returns **401** with a `WWW-Authenticate: NTLM, Negotiate` challenge — that is the healthy "published and reachable" signal.
- The authenticated probe tries **NTLM explicitly** (as Exchange Online does — it can't reach the on-prem KDC for Kerberos) *and* **Negotiate** via WinHTTP, logging each. WinHTTP is used because it can satisfy **Extended Protection** channel binding, which .NET's `HttpWebRequest` cannot.
- **HTTP 400 = authenticated success.** A `GET` to the WCF endpoint returns 400 once the credentials are accepted, because the endpoint expects a SOAP `POST`. This is normal and confirms a working MRS proxy.
- The **SOAP probe** POSTs a real SOAP 1.1 envelope. Because the MRSProxy WCF binding is Microsoft-internal (it rejects standard SOAP 1.1/1.2 with 415/503), a hand-crafted local SOAP call is often inconclusive — the raw server response is written to the Log, and any WCF fault confirms the service is live.

> Use the **on-prem migration credentials** in `DOMAIN\user` or UPN form (leave the field empty to use the logged-in Windows user). Note that a **shared/resource mailbox** usually has a disabled AD account and cannot authenticate.

**Verify from Exchange Online** (checkbox) runs the authoritative `Test-MigrationServerAvailability` inside EXO. It signs in with a device code (as an Exchange admin) and calls the AdminAPI. That REST path cannot serialise the `Credentials` parameter for this cmdlet, so the tool detects the error and shows a ready-to-run PowerShell command instead (double-click the row to copy it):

```powershell
Connect-ExchangeOnline
Test-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer <host> -Credentials (Get-Credential)
```

A successful `New-MigrationEndpoint` in Exchange Online runs the same server-side check and is the definitive end-to-end proof.

---

## 3. Hybrid Connectivity

A broad reachability sweep across on-prem + Exchange Online coexistence endpoints, using the **current Windows user**. Enter an e-mail (its domain drives discovery) and optionally an on-prem host, then **Test**:

| Category | Endpoints |
|---|---|
| AutoDiscover | on-prem (`autodiscover.<domain>` + root), Exchange Online V1 and V2 (`autodiscover.json`) |
| OAuth | on-prem auth-metadata (`/autodiscover/metadata/json/1`), tenant OpenID configuration |
| Free/Busy (EWS) | on-prem and Exchange Online `EWS/Exchange.asmx` |
| Health | `/<vdir>/healthcheck.htm` for the standard on-prem virtual directories |

Results use the same reachability verdict, colouring and CSV export as Additional Tests. On-prem endpoints authenticate with the current Windows user (**200 OK**); **Exchange Online** endpoints return **"Reachable — auth required"** (401 with a Bearer challenge) — this is the expected, healthy result for an unauthenticated probe, not an error. To probe Exchange Online endpoints *authenticated*, use the E-Mail AutoConfiguration test's Modern Auth + Additional Tests.

---

## 4. Free/Busy (Cross-Premises)

Tests hybrid **availability sharing** with real mailboxes, in both directions, using the standard EWS `GetUserAvailability` operation (a documented SOAP call — reliable, unlike the internal MRSProxy binding). Enter the on-prem EWS host, on-prem credentials, and one mailbox on each side, then **Test**:

| Direction | How | Auth |
|---|---|---|
| **On-Prem → EXO** | on-prem EWS queries the free/busy of the **EXO mailbox** | Windows (NTLM/Negotiate via WinHTTP — handles Extended Protection) |
| **EXO → On-Prem** | Exchange Online EWS queries the free/busy of the **on-prem mailbox** | OAuth — device-code sign-in as the EXO mailbox user |

Each row interprets the EWS response:

| Result | Meaning |
|---|---|
| **OK** | `ResponseCode NoError` + a real `FreeBusyViewType` → cross-premises free/busy works |
| **WARN** | `NoError` but `FreeBusyViewType None` → call succeeded but no data (missing cross-org access or calendar permission) |
| **FAIL** | an EWS error code / SOAP fault (shown), or HTTP 401 (auth failed) |

Hover a row for the raw response; **double-click** to view it in full (and copy it to the clipboard). Leaving the On-Prem User empty uses the logged-in Windows user for the On-Prem → EXO direction.

---

## Options (shared)

| Option | Description |
|--------|-------------|
| Ignore certificate errors | Bypasses TLS certificate validation — for on-premises Exchange with self-signed / untrusted certificates. (TLS trust failures also surface a hint to enable this.) |
| Use SCP (domain-joined) | AutoConfiguration only: runs the AD Service Connection Point lookup first, like Outlook on domain-joined machines |

---

## Files

```
ExchangeTester.bat      — Double-click launcher
ExchangeTester.ps1      — Main script (PowerShell 5.1, WinForms)
ExchangeTester.config   — Saved Client ID (created by Register App)
```

`ExchangeTester.bat` and `ExchangeTester.ps1` must be in the same folder.
