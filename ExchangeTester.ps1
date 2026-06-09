#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange AutoDiscover Tester
.DESCRIPTION
    Mirrors Outlook's "Test E-Mail AutoConfiguration" function (Ctrl+right-click tray icon).
    Tests AutoDiscover endpoints, shows HTTP sequence, redirects and the resulting XML.
.NOTES
    No installation required.
    Run: powershell.exe -ExecutionPolicy Bypass -File ExchangeTester.ps1
    Requires: Windows PowerShell 5.1, .NET Framework 4.x (standard on Windows 10/11)
    DNS SRV lookup requires Windows 8.1+ (Resolve-DnsName)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# TLS 1.2 / 1.1 / 1.0 for maximum compatibility with older Exchange servers
[System.Net.ServicePointManager]::SecurityProtocol =
    [System.Net.SecurityProtocolType]::Tls12 -bor
    [System.Net.SecurityProtocolType]::Tls11 -bor
    [System.Net.SecurityProtocolType]::Tls

#region ======================================================================
#  FORM
#==============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Test E-Mail AutoConfiguration"
$form.ClientSize      = New-Object System.Drawing.Size(775, 595)
$form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox     = $false
$form.MinimizeBox     = $true

#region --- Row 1: E-Mail Address ---
$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text      = "E-Mail Address"
$lblEmail.Location  = New-Object System.Drawing.Point(8, 14)
$lblEmail.Size      = New-Object System.Drawing.Size(100, 20)
$lblEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($lblEmail)

$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Location = New-Object System.Drawing.Point(112, 11)
$txtEmail.Size     = New-Object System.Drawing.Size(650, 22)
$txtEmail.TabIndex = 0
$form.Controls.Add($txtEmail)
#endregion

#region --- Row 2: Password ---
$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text      = "Password"
$lblPass.Location  = New-Object System.Drawing.Point(8, 42)
$lblPass.Size      = New-Object System.Drawing.Size(100, 20)
$lblPass.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblPass.Enabled   = $false
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location     = New-Object System.Drawing.Point(112, 39)
$txtPass.Size         = New-Object System.Drawing.Size(280, 22)
$txtPass.PasswordChar = [char]0x25CF   # ●
$txtPass.Enabled      = $false
$txtPass.TabIndex     = 1
$form.Controls.Add($txtPass)
#endregion

#region --- Rows 3-4: Checkboxes ---
# Row 3
$chkModernAuth = New-Object System.Windows.Forms.CheckBox
$chkModernAuth.Text      = "Try Modern Auth (OAuth2)"
$chkModernAuth.Location  = New-Object System.Drawing.Point(8, 70)
$chkModernAuth.Size      = New-Object System.Drawing.Size(180, 20)
$chkModernAuth.Checked   = $false
$chkModernAuth.TabIndex  = 2
$form.Controls.Add($chkModernAuth)

$chkUseCurrentUser = New-Object System.Windows.Forms.CheckBox
$chkUseCurrentUser.Text     = "Use logged-in user (Windows Auth)"
$chkUseCurrentUser.Location = New-Object System.Drawing.Point(195, 70)
$chkUseCurrentUser.Size     = New-Object System.Drawing.Size(225, 20)
$chkUseCurrentUser.Checked  = $true
$chkUseCurrentUser.TabIndex = 3
$form.Controls.Add($chkUseCurrentUser)

# Row 4
$chkIgnoreCert = New-Object System.Windows.Forms.CheckBox
$chkIgnoreCert.Text     = "Ignore certificate errors"
$chkIgnoreCert.Location = New-Object System.Drawing.Point(8, 94)
$chkIgnoreCert.Size     = New-Object System.Drawing.Size(180, 20)
$chkIgnoreCert.Checked  = $false
$chkIgnoreCert.TabIndex = 4
$form.Controls.Add($chkIgnoreCert)

$chkUseSCP = New-Object System.Windows.Forms.CheckBox
$chkUseSCP.Text     = "Use SCP (domain-joined)"
$chkUseSCP.Location = New-Object System.Drawing.Point(195, 94)
$chkUseSCP.Size     = New-Object System.Drawing.Size(175, 20)
$chkUseSCP.Checked  = $true
$chkUseSCP.TabIndex = 5
$form.Controls.Add($chkUseSCP)
#endregion

#region --- Buttons (top-right) ---
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text     = "Test"
$btnTest.Location = New-Object System.Drawing.Point(620, 75)
$btnTest.Size     = New-Object System.Drawing.Size(68, 26)
$btnTest.TabIndex = 6
$form.Controls.Add($btnTest)
$form.AcceptButton = $btnTest

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(696, 75)
$btnCancel.Size     = New-Object System.Drawing.Size(68, 26)
$btnCancel.Enabled  = $false
$btnCancel.TabIndex = 7
$form.Controls.Add($btnCancel)
#endregion

#region --- Row 5: OAuth2 Client ID ---
$lblClientId = New-Object System.Windows.Forms.Label
$lblClientId.Text      = "OAuth2 Client ID:"
$lblClientId.Location  = New-Object System.Drawing.Point(8, 119)
$lblClientId.Size      = New-Object System.Drawing.Size(112, 20)
$lblClientId.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblClientId.Enabled   = $false
$form.Controls.Add($lblClientId)

$txtClientId = New-Object System.Windows.Forms.TextBox
$txtClientId.Location  = New-Object System.Drawing.Point(122, 116)
$txtClientId.Size      = New-Object System.Drawing.Size(490, 22)
$txtClientId.Text      = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
$txtClientId.Enabled   = $false
$txtClientId.TabIndex  = 8
$txtClientId.Font      = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtClientId)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($txtClientId, "Azure AD client ID used for the OAuth2 login popup. Default = Microsoft Office public client. Change only if your tenant has blocked it.")
$toolTip.SetToolTip($lblClientId, "Azure AD client ID used for the OAuth2 login popup. Default = Microsoft Office public client. Change only if your tenant has blocked it.")
#endregion

#region --- Separator + Progress bar ---
$pnlSep = New-Object System.Windows.Forms.Panel
$pnlSep.Location  = New-Object System.Drawing.Point(0, 144)
$pnlSep.Size      = New-Object System.Drawing.Size(775, 2)
$pnlSep.BackColor = [System.Drawing.SystemColors]::ControlDark
$form.Controls.Add($pnlSep)

$prgBar = New-Object System.Windows.Forms.ProgressBar
$prgBar.Location = New-Object System.Drawing.Point(8, 152)
$prgBar.Size     = New-Object System.Drawing.Size(757, 14)
$prgBar.Minimum  = 0
$prgBar.Maximum  = 100
$prgBar.Value    = 0
$form.Controls.Add($prgBar)
#endregion

#region --- TabControl ---
$tabCtrl = New-Object System.Windows.Forms.TabControl
$tabCtrl.Location = New-Object System.Drawing.Point(8, 172)
$tabCtrl.Size     = New-Object System.Drawing.Size(757, 414)
$form.Controls.Add($tabCtrl)

# Tab: Results
$tabResults = New-Object System.Windows.Forms.TabPage
$tabResults.Text = "Results"
$tabCtrl.Controls.Add($tabResults)

$lvwResults = New-Object System.Windows.Forms.ListView
$lvwResults.Dock          = [System.Windows.Forms.DockStyle]::Fill
$lvwResults.View          = [System.Windows.Forms.View]::Details
$lvwResults.FullRowSelect = $true
$lvwResults.GridLines     = $true
$lvwResults.ShowGroups    = $true
$lvwResults.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
[void]$lvwResults.Columns.Add("Setting", 220)
[void]$lvwResults.Columns.Add("Value",   510)
$tabResults.Controls.Add($lvwResults)

# Tab: Log
$tabLog = New-Object System.Windows.Forms.TabPage
$tabLog.Text = "Log"
$tabCtrl.Controls.Add($tabLog)

$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Dock        = [System.Windows.Forms.DockStyle]::Fill
$rtbLog.ReadOnly    = $true
$rtbLog.Font        = New-Object System.Drawing.Font("Consolas", 9)
$rtbLog.BackColor   = [System.Drawing.Color]::White
$rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$rtbLog.WordWrap    = $true
$tabLog.Controls.Add($rtbLog)

# Tab: XML
$tabXml = New-Object System.Windows.Forms.TabPage
$tabXml.Text = "XML"
$tabCtrl.Controls.Add($tabXml)

$rtbXml = New-Object System.Windows.Forms.RichTextBox
$rtbXml.Dock       = [System.Windows.Forms.DockStyle]::Fill
$rtbXml.ReadOnly   = $true
$rtbXml.Font       = New-Object System.Drawing.Font("Consolas", 9)
$rtbXml.BackColor  = [System.Drawing.Color]::White
$rtbXml.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$rtbXml.WordWrap   = $false
$tabXml.Controls.Add($rtbXml)
#endregion

#region --- Context menus for Log and XML ---

# Shared items: Copy selection / Select All
$ctxLog = New-Object System.Windows.Forms.ContextMenuStrip
$miLogCopy = New-Object System.Windows.Forms.ToolStripMenuItem("Copy")
$miLogCopy.ShortcutKeyDisplayString = "Ctrl+C"
$miLogSelAll = New-Object System.Windows.Forms.ToolStripMenuItem("Select All")
$miLogSelAll.ShortcutKeyDisplayString = "Ctrl+A"
$miLogClear = New-Object System.Windows.Forms.ToolStripMenuItem("Clear Log")
[void]$ctxLog.Items.Add($miLogCopy)
[void]$ctxLog.Items.Add($miLogSelAll)
[void]$ctxLog.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ctxLog.Items.Add($miLogClear)
$rtbLog.ContextMenuStrip = $ctxLog

$miLogCopy.Add_Click({
    $t = $rtbLog.SelectedText
    if (-not $t) { $t = $rtbLog.Text }
    if ($t) { [System.Windows.Forms.Clipboard]::SetText($t) }
})
$miLogSelAll.Add_Click({ $rtbLog.SelectAll() })
$miLogClear.Add_Click({ $rtbLog.Clear() })

$ctxXml = New-Object System.Windows.Forms.ContextMenuStrip
$miXmlCopy   = New-Object System.Windows.Forms.ToolStripMenuItem("Copy")
$miXmlCopy.ShortcutKeyDisplayString = "Ctrl+C"
$miXmlSelAll = New-Object System.Windows.Forms.ToolStripMenuItem("Select All")
$miXmlSelAll.ShortcutKeyDisplayString = "Ctrl+A"
$miXmlSave   = New-Object System.Windows.Forms.ToolStripMenuItem("Save XML As...")
[void]$ctxXml.Items.Add($miXmlCopy)
[void]$ctxXml.Items.Add($miXmlSelAll)
[void]$ctxXml.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ctxXml.Items.Add($miXmlSave)
$rtbXml.ContextMenuStrip = $ctxXml

$miXmlCopy.Add_Click({
    $t = $rtbXml.SelectedText
    if (-not $t) { $t = $rtbXml.Text }
    if ($t) { [System.Windows.Forms.Clipboard]::SetText($t) }
})
$miXmlSelAll.Add_Click({ $rtbXml.SelectAll() })
$miXmlSave.Add_Click({
    if (-not $rtbXml.Text) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    $sfd.FileName = "autodiscover.xml"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        [System.IO.File]::WriteAllText($sfd.FileName, $rtbXml.Text, [System.Text.Encoding]::UTF8)
        [System.Windows.Forms.MessageBox]::Show(
            "Saved to:`n$($sfd.FileName)", "Saved",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
})
#endregion

#endregion ===================================================================
#  XML PARSER
#==============================================================================

function ConvertFrom-AutodiscoverXml {
    param([string]$RawXml)

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $xd = [xml]$RawXml
        $ns = New-Object System.Xml.XmlNamespaceManager($xd.NameTable)
        $ns.AddNamespace("ad", "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a")

        # Returns inner text of first matching node, or $null
        $txt = {
            param($node, [string]$xpath)
            $n = $node.SelectSingleNode($xpath, $ns)
            if ($n) { $n.InnerText } else { $null }
        }

        $add = {
            param([string]$group, [string]$setting, $value)
            if ($value -and $value.Trim() -ne '') {
                $rows.Add([PSCustomObject]@{ Group = $group; Setting = $setting; Value = $value.Trim() })
            }
        }

        # --- User ---
        $user = $xd.SelectSingleNode("//ad:User", $ns)
        if ($user) {
            & $add "User" "Display Name"  (& $txt $user "ad:DisplayName")
            & $add "User" "Legacy DN"      (& $txt $user "ad:LegacyDN")
            & $add "User" "Deployment ID"  (& $txt $user "ad:DeploymentId")
        }

        # --- Account ---
        $account = $xd.SelectSingleNode("//ad:Account", $ns)
        if ($account) {
            $action = & $txt $account "ad:Action"
            & $add "Account" "Account Type"     (& $txt $account "ad:AccountType")
            & $add "Account" "Action"             $action
            & $add "Account" "Redirect Address"  (& $txt $account "ad:RedirectAddr")
            & $add "Account" "Redirect URL"      (& $txt $account "ad:RedirectUrl")
        }

        # --- Protocol sections (EXCH, EXPR, IMAP, POP3, SMTP, WEB, ...) ---
        $protocols = $xd.SelectNodes("//ad:Protocol", $ns)
        foreach ($proto in $protocols) {
            $type = & $txt $proto "ad:Type"
            if (-not $type) { $type = "Unknown" }
            $grp = "Protocol: $type"

            & $add $grp "Server"                  (& $txt $proto "ad:Server")
            & $add $grp "Port"                    (& $txt $proto "ad:Port")
            & $add $grp "SSL"                     (& $txt $proto "ad:SSL")
            & $add $grp "Encryption"              (& $txt $proto "ad:Encryption")
            & $add $grp "Login Name"              (& $txt $proto "ad:LoginName")
            & $add $grp "Domain Required"         (& $txt $proto "ad:DomainRequired")
            & $add $grp "Auth Required"           (& $txt $proto "ad:AuthRequired")
            & $add $grp "Auth Package"            (& $txt $proto "ad:AuthPackage")
            & $add $grp "EWS URL"                 (& $txt $proto "ad:EwsUrl")
            & $add $grp "OAB URL"                 (& $txt $proto "ad:OABUrl")
            & $add $grp "OOF URL"                 (& $txt $proto "ad:OOFUrl")
            & $add $grp "ActiveSync URL"          (& $txt $proto "ad:ASUrl")
            & $add $grp "EMWS URL"                (& $txt $proto "ad:EmwsUrl")
            & $add $grp "Public Folder Server"    (& $txt $proto "ad:PublicFolderServer")
            & $add $grp "Autodiscover Internal"   (& $txt $proto "ad:AutodiscoverServiceInternalUri")

            # WEB protocol: OWA URLs are nested under Internal/External
            $inOwa  = $proto.SelectSingleNode("ad:Internal/ad:OWAUrl",  $ns)
            $extOwa = $proto.SelectSingleNode("ad:External/ad:OWAUrl", $ns)
            if ($inOwa)  { & $add $grp "OWA URL (Internal)" $inOwa.InnerText }
            if ($extOwa) { & $add $grp "OWA URL (External)" $extOwa.InnerText }
        }
    } catch {
        $rows.Add([PSCustomObject]@{
            Group   = "Error"
            Setting = "XML parse error"
            Value   = $_.Exception.Message
        })
    }

    return ,$rows   # comma forces List to stay as-is, not unroll
}

#endregion ===================================================================
#  TEST ENGINE  (PowerShell Runspace + WinForms Timer — avoids ThreadPool runspace issue)
#==============================================================================

# Script-level state for the currently running test
$script:CurrentPS      = $null
$script:CurrentRS      = $null
$script:CurrentSync    = $null
$script:PollTimer      = $null

# The actual test logic runs inside a dedicated PS runspace.
# $sync is the only bridge between the runspace and the UI thread.
$script:TestScript = {
    param($sync)

    $email      = $sync.Email
    $password   = $sync.Password
    $useWinAuth = $sync.UseWindowsAuth
    $domain     = ($email -split '@')[1]

    # Credentials
    $netCred = $null
    if (-not $useWinAuth -and $password -ne '') {
        $netCred = New-Object System.Net.NetworkCredential($email, $password)
    }

    # AutoDiscover POST body
    $bodyXml = "<?xml version=""1.0"" encoding=""utf-8""?>" +
        "<Autodiscover xmlns=""http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"">" +
        "<Request><EMailAddress>$email</EMailAddress>" +
        "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>" +
        "</Request></Autodiscover>"
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyXml)

    # Helpers — all run inside the dedicated runspace, no closure issues
    $logLine = { param([string]$msg) $sync.Queue.Enqueue($msg) }
    $setPct  = { param([int]$pct)   $sync.Pct = $pct }

    # HTTP POST → @{Code; Body; Location; WwwAuth; Error}
    # Pass $authHeader to override credentials with an explicit Authorization value (e.g. Bearer token)
    $doPost = {
        param([string]$url, [string]$authHeader = '')
        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method            = "POST"
            $req.ContentType       = "text/xml; charset=utf-8"
            $req.ContentLength     = $bodyBytes.Length
            $req.AllowAutoRedirect = $false
            $req.Timeout           = 30000
            $req.UserAgent         = "Microsoft Office/16.0 (Windows NT 10.0)"
            if ($authHeader -ne '') {
                $req.Headers["Authorization"] = $authHeader
            } elseif ($useWinAuth) {
                $req.UseDefaultCredentials = $true
            } elseif ($netCred) {
                $req.Credentials = $netCred
            }
            $s = $req.GetRequestStream()
            $s.Write($bodyBytes, 0, $bodyBytes.Length)
            $s.Close()
            try {
                $resp = $req.GetResponse()
                $code = [int]$resp.StatusCode
                $body = $null
                if ($code -eq 200) {
                    $rs   = $resp.GetResponseStream()
                    $rdr  = New-Object System.IO.StreamReader($rs, [System.Text.Encoding]::UTF8)
                    $body = $rdr.ReadToEnd()
                    $rdr.Close()
                }
                $resp.Close()
                return @{ Code = $code; Body = $body; Location = $null; WwwAuth = $null; Error = $null }
            } catch [System.Net.WebException] {
                $ex = $_.Exception
                if ($ex.Response) {
                    $code    = [int]$ex.Response.StatusCode
                    $loc          = $ex.Response.Headers["Location"]
                    $wwwAuthVals  = try { $ex.Response.Headers.GetValues("WWW-Authenticate") } catch { $null }
                    $wwwAuth      = if ($wwwAuthVals) { $wwwAuthVals -join ' ' } else { $null }
                    $ex.Response.Close()
                    return @{ Code = $code; Body = $null; Location = $loc; WwwAuth = $wwwAuth; Error = $null }
                }
                return @{ Code = -1; Body = $null; Location = $null; WwwAuth = $null; Error = $ex.Message }
            }
        } catch {
            return @{ Code = -1; Body = $null; Location = $null; WwwAuth = $null; Error = $_.Exception.Message }
        }
    }

    # HTTP GET → @{Code; Location; Error}
    $doGet = {
        param([string]$url)
        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method            = "GET"
            $req.AllowAutoRedirect = $false
            $req.Timeout           = 15000
            $req.UserAgent         = "Microsoft Office/16.0 (Windows NT 10.0)"
            try {
                $resp = $req.GetResponse()
                $code = [int]$resp.StatusCode
                $loc  = $resp.Headers["Location"]
                $resp.Close()
                return @{ Code = $code; Location = $loc; Error = $null }
            } catch [System.Net.WebException] {
                $ex = $_.Exception
                if ($ex.Response) {
                    $code = [int]$ex.Response.StatusCode
                    $loc  = $ex.Response.Headers["Location"]
                    $ex.Response.Close()
                    return @{ Code = $code; Location = $loc; Error = $null }
                }
                return @{ Code = -1; Location = $null; Error = $ex.Message }
            }
        } catch {
            return @{ Code = -1; Location = $null; Error = $_.Exception.Message }
        }
    }

    $tryModernAuth = $sync.ModernAuth
    $useSCP        = $sync.UseSCP

    # Authorization Code Flow with PKCE — embedded IE WebBrowser popup, no external browser.
    # Intercepts the http://localhost redirect before the browser attempts the connection.
    $getToken = {
        param([string]$wwwAuthHeader)

        Add-Type -AssemblyName System.Web

        # Normalize the authorization_uri from the Bearer challenge to v2.0
        $authUri = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
        if ($wwwAuthHeader -match 'authorization_uri\s*=\s*"([^"]+)"') {
            $authUri = $Matches[1] -replace '/oauth2(?:/v2\.0)?/authorize.*', '/oauth2/v2.0/authorize'
        }
        $tokenUrl    = $authUri -replace '/authorize', '/token'
        $clientId    = if ($sync.ClientId) { $sync.ClientId } else { 'd3590ed6-52b3-4102-aeff-aad2292ab01c' }
        $redirectUri = 'http://localhost'
        $scope       = 'https://outlook.office365.com/.default offline_access'

        # PKCE: 48 random bytes → base64url code_verifier, SHA-256 → code_challenge
        $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
        $buf = New-Object byte[] 48
        $rng.GetBytes($buf)
        $codeVerifier  = [Convert]::ToBase64String($buf).TrimEnd('=').Replace('+','-').Replace('/','_')
        $sha256        = [System.Security.Cryptography.SHA256]::Create()
        $codeChallenge = [Convert]::ToBase64String(
            $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($codeVerifier))
        ).TrimEnd('=').Replace('+','-').Replace('/','_')

        $state   = [Guid]::NewGuid().ToString("N")
        $authUrl = $authUri +
            "?client_id=$clientId" +
            "&response_type=code" +
            "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
            "&scope=$([Uri]::EscapeDataString($scope))" +
            "&code_challenge=$codeChallenge" +
            "&code_challenge_method=S256" +
            "&login_hint=$([Uri]::EscapeDataString($email))" +
            "&state=$state" +
            "&prompt=select_account"

        & $logLine "Modern Auth: opening sign-in popup…"

        # Force WebBrowser control to use IE11 rendering engine for this process
        try {
            $procName = [System.IO.Path]::GetFileName(
                [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
            $feKey = "HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
            if (-not (Test-Path $feKey)) { New-Item -Path $feKey -Force | Out-Null }
            Set-ItemProperty -Path $feKey -Name $procName -Value 11001 -Type DWord
        } catch {}

        # Clear any leftover OAuth state from a previous attempt
        $sync.OAuthCode    = $null
        $sync.OAuthErr     = $null
        $sync.OAuthErrDesc = $null

        $loginForm = New-Object System.Windows.Forms.Form
        $loginForm.Text          = "Sign in to Microsoft  —  Exchange Tester"
        $loginForm.Size          = New-Object System.Drawing.Size(520, 660)
        $loginForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $loginForm.MinimizeBox   = $false

        $wb = New-Object System.Windows.Forms.WebBrowser
        $wb.Dock                   = [System.Windows.Forms.DockStyle]::Fill
        $wb.ScriptErrorsSuppressed = $true
        $loginForm.Controls.Add($wb)

        # Intercept the redirect to http://localhost BEFORE the browser tries to connect.
        # Using $sync for data transfer and $s.FindForm() avoids PowerShell closure capture issues.
        $wb.Add_Navigating({
            param($s, $e)
            try {
                $uri = $e.Url
                if (-not $uri) { return }
                if ($uri.Scheme -eq 'http' -and $uri.Host -eq 'localhost') {
                    $qs = [System.Web.HttpUtility]::ParseQueryString($uri.Query)
                    $sync.OAuthCode    = $qs["code"]
                    $sync.OAuthErr     = $qs["error"]
                    $sync.OAuthErrDesc = $qs["error_description"]
                    $e.Cancel = $true
                    $s.FindForm().Close()
                }
            } catch {}
        })

        $loginForm.Add_Shown({ $wb.Navigate($authUrl) })
        [void]$loginForm.ShowDialog()
        $loginForm.Dispose()

        $code      = $sync.OAuthCode;    $sync.OAuthCode    = $null
        $cbErr     = $sync.OAuthErr;     $sync.OAuthErr     = $null
        $cbErrDesc = $sync.OAuthErrDesc; $sync.OAuthErrDesc = $null

        if ($cbErr) {
            & $logLine "Modern Auth: sign-in error — ${cbErr}: $cbErrDesc"
            return $null
        }
        if (-not $code) {
            & $logLine "Modern Auth: sign-in cancelled."
            return $null
        }

        # Exchange authorization code for access token
        $tokBody = "grant_type=authorization_code" +
            "&client_id=$clientId" +
            "&code=$([Uri]::EscapeDataString($code))" +
            "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
            "&code_verifier=$codeVerifier" +
            "&scope=$([Uri]::EscapeDataString($scope))"
        $tokBodyBytes = [System.Text.Encoding]::UTF8.GetBytes($tokBody)
        try {
            $req = [System.Net.HttpWebRequest]::Create($tokenUrl)
            $req.Method        = "POST"
            $req.ContentType   = "application/x-www-form-urlencoded"
            $req.ContentLength = $tokBodyBytes.Length
            $req.Timeout       = 30000
            $s = $req.GetRequestStream(); $s.Write($tokBodyBytes, 0, $tokBodyBytes.Length); $s.Close()
            try {
                $resp = $req.GetResponse()
                $tok  = (New-Object System.IO.StreamReader($resp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                $resp.Close()
                if ($tok.access_token) {
                    & $logLine "Modern Auth: access token acquired."
                    return "Bearer $($tok.access_token)"
                }
                & $logLine "Modern Auth: token response missing access_token."
                return $null
            } catch [System.Net.WebException] {
                $exT = $_.Exception
                if ($exT.Response) {
                    $errJ = (New-Object System.IO.StreamReader($exT.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                    $exT.Response.Close()
                    & $logLine "Modern Auth: token exchange failed — $($errJ.error): $($errJ.error_description)"
                } else {
                    & $logLine "Modern Auth: token exchange failed — $($exT.Message)"
                }
                return $null
            }
        } catch {
            & $logLine "Modern Auth: token exchange error — $($_.Exception.Message)"
            return $null
        }
    }

    # Try one AutoDiscover POST URL; returns XML string on 200, $null otherwise
    $tryUrl = {
        param([string]$url)
        & $logLine "AutoDiscover via $url starting."
        $res = & $doPost $url

        if ($res.Code -eq 200) {
            & $logLine "GetLastError=0; httpStatus=200."
            & $logLine "AutoDiscover via $url succeeded."
            return $res.Body
        }

        if ($res.Code -ge 0) {
            & $logLine "GetLastError=0; httpStatus=$($res.Code)."
        } else {
            & $logLine "AutoDiscover via $url failed: $($res.Error)"
            return $null
        }

        if ($res.Code -eq 301 -or $res.Code -eq 302) {
            $loc = $res.Location
            if ($loc) { & $logLine "URL redirect to $loc at AutoDiscover." }
            if ($loc -and $loc -match '^https://' -and $loc -match 'autodiscover') {
                # Redirect points directly at an autodiscover endpoint — follow it
                & $logLine "Redirect check for $loc starting."
                $res2 = & $doPost $loc
                if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                if ($res2.Code -eq 200) {
                    & $logLine "AutoDiscover via $loc succeeded."
                    return $res2.Body
                }
                & $logLine "Redirect check for $loc failed (0x800C8209)."
            } elseif ($loc) {
                # Non-autodiscover redirect target — GET to verify it's a real endpoint
                & $logLine "Redirect check for $loc starting."
                $res2 = & $doGet $loc
                if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                & $logLine "Redirect check for $loc failed (0x800C8209)."
            }
            & $logLine "AutoDiscover via $url failed (0x800C8204)."
            return $null
        }

        if ($res.Code -eq 401) {
            if ($res.WwwAuth) {
                & $logLine "  WWW-Authenticate: $($res.WwwAuth)"
            } else {
                & $logLine "  WWW-Authenticate: (not present)"
            }
            $hasBearerChallenge = $res.WwwAuth -and $res.WwwAuth -match 'Bearer'
            $isO365Endpoint     = $url -match 'outlook\.office365\.com'
            if ($tryModernAuth -and ($hasBearerChallenge -or $isO365Endpoint)) {
                if ($isO365Endpoint -and -not $hasBearerChallenge) {
                    & $logLine "  O365 endpoint detected — attempting Modern Auth proactively (Bearer not advertised)."
                }
                & $logLine "Modern Auth: initiating OAuth2 Authorization Code Flow…"
                $token = & $getToken $res.WwwAuth
                if ($token) {
                    & $logLine "Retrying AutoDiscover with Bearer token."
                    $res2 = & $doPost $url $token
                    if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
                    if ($res2.Code -eq 200) {
                        & $logLine "AutoDiscover via $url succeeded (Modern Auth)."
                        return $res2.Body
                    }
                    & $logLine "AutoDiscover via $url failed after Modern Auth (httpStatus=$($res2.Code))."
                    return $null
                }
            } elseif ($tryModernAuth) {
                & $logLine "  No Bearer challenge — HMA/OAuth2 not offered by this endpoint."
            }
            & $logLine "AutoDiscover via $url failed (0x800C820E)."
            return $null
        }

        & $logLine "AutoDiscover via $url failed (httpStatus=$($res.Code))."
        return $null
    }

    #endregion

    $foundXml = $null

    # --- Step 1: SCP — runs first when "Use SCP" is checked (domain-joined priority) ---
    & $setPct 5
    if (-not $sync.Cancel -and $useSCP) {
        & $logLine "Local AutoDiscover for $domain starting (SCP)."
        try {
            # Exchange AutoDiscover SCPs are stored in the Configuration partition,
            # not the default domain partition — must set search root explicitly.
            $rootDSE    = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
            $configNC   = $rootDSE.Properties["configurationNamingContext"].Value
            & $logLine "  SCP search root: $configNC"
            $searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$configNC")

            $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
            # Match on the Exchange AutoDiscover GUID keyword (with or without braces)
            # OR on serviceBindingInformation containing "autodiscover" as a fallback.
            $searcher.Filter      = "(&(objectClass=serviceConnectionPoint)(|(serviceBindingInformation=*autodiscover*)(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords={67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68})))"
            $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            [void]$searcher.PropertiesToLoad.Add("serviceBindingInformation")

            $scpHits = $searcher.FindAll()
            & $logLine "  SCP records found: $($scpHits.Count)"
            if ($scpHits.Count -eq 0) {
                & $logLine "Local AutoDiscover for $domain failed (0x8004010F)."
            } else {
                foreach ($hit in $scpHits) {
                    $scpUrl = $hit.Properties["serviceBindingInformation"][0]
                    & $logLine "SCP record: $scpUrl"
                    $res = & $doPost $scpUrl
                    if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
                    if ($res.Code -eq 200) {
                        $foundXml = $res.Body
                        & $logLine "Local AutoDiscover via SCP succeeded."
                        break
                    }
                }
                if (-not $foundXml) { & $logLine "Local AutoDiscover for $domain failed." }
            }
        } catch {
            & $logLine "Local AutoDiscover for $domain failed (0x8004010F): $($_.Exception.Message)"
        }
    }

    # --- Step 2: O365 ---
    & $setPct 20
    if (-not $foundXml -and -not $sync.Cancel) {
        $foundXml = & $tryUrl "https://outlook.office365.com/autodiscover/autodiscover.xml"
    }

    # --- Step 3: https://<domain>/autodiscover/autodiscover.xml ---
    & $setPct 40
    if (-not $foundXml -and -not $sync.Cancel) {
        $foundXml = & $tryUrl "https://$domain/autodiscover/autodiscover.xml"
    }

    # --- Step 4: https://autodiscover.<domain>/autodiscover/autodiscover.xml ---
    & $setPct 58
    if (-not $foundXml -and -not $sync.Cancel) {
        $foundXml = & $tryUrl "https://autodiscover.$domain/autodiscover/autodiscover.xml"
    }

    # --- Step 5: HTTP redirect check (well-known URL) ---
    & $setPct 74
    if (-not $foundXml -and -not $sync.Cancel) {
        $rdUrl = "http://autodiscover.$domain/autodiscover/autodiscover.xml"
        & $logLine "Redirect check for $rdUrl starting."
        $res = & $doGet $rdUrl
        if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
        if ($res.Location -and $res.Location -match '^https://') {
            & $logLine "Redirect to $($res.Location) found."
            $res2 = & $doPost $res.Location
            if ($res2.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res2.Code)." }
            if ($res2.Code -eq 200) {
                $foundXml = $res2.Body
                & $logLine "AutoDiscover via HTTP redirect succeeded."
            } else {
                & $logLine "Redirect check for $rdUrl failed (0x80004005)."
            }
        } else {
            & $logLine "Redirect check for $rdUrl failed (0x80004005)."
        }
    }

    # --- Step 6: DNS SRV _autodiscover._tcp.<domain> ---
    & $setPct 88
    if (-not $foundXml -and -not $sync.Cancel) {
        & $logLine "DNS SRV lookup for $domain starting."
        try {
            $srvRecs = Resolve-DnsName -Name "_autodiscover._tcp.$domain" -Type SRV -ErrorAction Stop
            $srvHit = $false
            foreach ($srv in $srvRecs) {
                if (-not $srv.NameTarget) { continue }
                $srvUrl = "https://$($srv.NameTarget):$($srv.Port)/autodiscover/autodiscover.xml"
                & $logLine "DNS SRV: $($srv.NameTarget):$($srv.Port)"
                $res = & $doPost $srvUrl
                if ($res.Code -ge 0) { & $logLine "GetLastError=0; httpStatus=$($res.Code)." }
                if ($res.Code -eq 200) {
                    $foundXml = $res.Body
                    $srvHit   = $true
                    & $logLine "AutoDiscover via DNS SRV succeeded."
                    break
                }
            }
            if (-not $srvHit) { & $logLine "DNS SRV lookup for $domain failed (0x8004010F)." }
        } catch [System.Management.Automation.CommandNotFoundException] {
            & $logLine "DNS SRV skipped (Resolve-DnsName requires Windows 8.1+)."
        } catch {
            & $logLine "DNS SRV lookup for $domain failed (0x8004010F)."
        }
    }

    if ($sync.Cancel) { $sync.Done = $true; return }

    & $setPct 100
    $sync.Xml  = $foundXml
    $sync.Done = $true
}

# Called from the poll timer when $sync.Done becomes $true
function Complete-Test {
    $script:PollTimer.Stop()

    # Drain any remaining log lines
    $msg = $null
    while ($script:CurrentSync.Queue.TryDequeue([ref]$msg)) {
        $rtbLog.AppendText("$msg`r`n")
    }

    $btnTest.Enabled   = $true
    $btnCancel.Enabled = $false

    if ($script:CurrentSync.Cancel) {
        $form.Text = "Test E-Mail AutoConfiguration"
        $rtbLog.AppendText("`r`nTest cancelled.`r`n")
    } elseif ($script:CurrentSync.Xml) {
        $xml = $script:CurrentSync.Xml

        # XML tab: pretty-print
        try {
            $xd = New-Object System.Xml.XmlDocument
            $xd.LoadXml($xml)
            $sb = New-Object System.Text.StringBuilder
            $sw = New-Object System.IO.StringWriter($sb)
            $xw = New-Object System.Xml.XmlTextWriter($sw)
            $xw.Formatting  = [System.Xml.Formatting]::Indented
            $xw.Indentation = 2
            $xd.WriteTo($xw)
            $xw.Flush()
            $rtbXml.Text = $sb.ToString()
        } catch {
            $rtbXml.Text = $xml
        }

        # Results tab: parsed rows grouped by section
        $rows = ConvertFrom-AutodiscoverXml -RawXml $xml
        $lvwResults.BeginUpdate()
        $lvwResults.Items.Clear()
        $lvwResults.Groups.Clear()
        $groupMap = @{}
        foreach ($row in $rows) {
            if (-not $groupMap.ContainsKey($row.Group)) {
                $lvg = New-Object System.Windows.Forms.ListViewGroup($row.Group, $row.Group)
                [void]$lvwResults.Groups.Add($lvg)
                $groupMap[$row.Group] = $lvg
            }
            $item = New-Object System.Windows.Forms.ListViewItem($row.Setting)
            [void]$item.SubItems.Add($row.Value)
            $item.Group = $groupMap[$row.Group]
            [void]$lvwResults.Items.Add($item)
        }
        $lvwResults.EndUpdate()

        $form.Text = "Test E-Mail AutoConfiguration  —  OK"
        $rtbLog.AppendText("`r`nAutoDiscover completed successfully.`r`n")
        if ($rows.Count -gt 0) { $tabCtrl.SelectedTab = $tabResults }
        else                    { $tabCtrl.SelectedTab = $tabXml }
    } else {
        $form.Text = "Test E-Mail AutoConfiguration  —  No configuration found"
        $rtbLog.AppendText("`r`nAutoDiscover failed for all tested methods.`r`n")
        $tabCtrl.SelectedTab = $tabLog
    }

    # Clean up runspace
    try { $script:CurrentPS.Dispose() }   catch {}
    try { $script:CurrentRS.Close();  $script:CurrentRS.Dispose() } catch {}
    $script:CurrentPS   = $null
    $script:CurrentRS   = $null
    $script:CurrentSync = $null
}

#endregion ===================================================================
#  CONTROL INTERACTIONS
#==============================================================================

$chkModernAuth.Add_CheckedChanged({
    $en = $chkModernAuth.Checked
    $lblClientId.Enabled = $en
    $txtClientId.Enabled = $en
})

$chkUseCurrentUser.Add_CheckedChanged({
    $useExplicit       = -not $chkUseCurrentUser.Checked
    $txtPass.Enabled   = $useExplicit
    $lblPass.Enabled   = $useExplicit
    if ($chkUseCurrentUser.Checked) { $txtPass.Clear() }
})

$btnTest.Add_Click({
    $email = $txtEmail.Text.Trim()
    if ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid e-mail address.",
            "Input Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }


    # Certificate validation callback
    if ($chkIgnoreCert.Checked) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    } else {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }

    # Reset UI
    $form.Text = "Test E-Mail AutoConfiguration  —  $email"
    $rtbLog.Clear()
    $rtbXml.Clear()
    $lvwResults.Groups.Clear()
    $lvwResults.Items.Clear()
    $prgBar.Value        = 0
    $btnTest.Enabled     = $false
    $btnCancel.Enabled   = $true
    $tabCtrl.SelectedTab = $tabLog

    # Build the sync hash that bridges UI thread and test runspace
    $sync = [hashtable]::Synchronized(@{
        Email          = $email
        Password       = $txtPass.Text
        UseWindowsAuth = $chkUseCurrentUser.Checked
        ModernAuth     = $chkModernAuth.Checked
        ClientId       = $txtClientId.Text.Trim()
        UseSCP         = $chkUseSCP.Checked
        Cancel         = $false
        Done           = $false
        Xml            = $null
        Pct            = 0
        Queue          = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    })
    $script:CurrentSync = $sync

    # Create a dedicated runspace so PS script blocks work without issues
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
    $rs.Open()
    $script:CurrentRS = $rs

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:TestScript).AddArgument($sync)
    $script:CurrentPS = $ps
    [void]$ps.BeginInvoke()

    # Poll timer: drain log queue and check for completion (runs on UI thread)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100
    $timer.Add_Tick({
        $msg = $null
        while ($script:CurrentSync -and $script:CurrentSync.Queue.TryDequeue([ref]$msg)) {
            $rtbLog.AppendText("$msg`r`n")
            $rtbLog.ScrollToCaret()
        }
        if ($script:CurrentSync) {
            $v = $script:CurrentSync.Pct
            if ($v -ge 0 -and $v -le 100) { $prgBar.Value = $v }
        }
        if ($script:CurrentSync -and $script:CurrentSync.Done) {
            Complete-Test
        }
    })
    $script:PollTimer = $timer
    $timer.Start()
})

$btnCancel.Add_Click({
    if ($script:CurrentSync) { $script:CurrentSync.Cancel = $true }
    $btnCancel.Enabled = $false
    $rtbLog.AppendText("Cancelling...`r`n")
})

$form.Add_FormClosing({
    if ($script:CurrentSync) { $script:CurrentSync.Cancel = $true }
    if ($script:PollTimer)   { $script:PollTimer.Stop() }
    try { $script:CurrentPS.Dispose() }   catch {}
    try { $script:CurrentRS.Close();  $script:CurrentRS.Dispose() } catch {}
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
})

#==============================================================================
#  START
#==============================================================================

# Pre-fill e-mail with the logged-in user's UPN when available
try {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $upnClaim = $id.Claims | Where-Object { $_.Type -eq 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn' } | Select-Object -First 1
    if ($upnClaim -and $upnClaim.Value) { $txtEmail.Text = $upnClaim.Value }
} catch {}

[System.Windows.Forms.Application]::Run($form)
