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

#region --- Auth mode (top-level radio buttons) ---
$radModernAuth = New-Object System.Windows.Forms.RadioButton
$radModernAuth.Text     = "Modern Auth (OAuth2)"
$radModernAuth.Location = New-Object System.Drawing.Point(8, 65)
$radModernAuth.Size     = New-Object System.Drawing.Size(185, 20)
$radModernAuth.Checked  = $true
$radModernAuth.TabIndex = 2
$form.Controls.Add($radModernAuth)

$radWIA = New-Object System.Windows.Forms.RadioButton
$radWIA.Text     = "Windows Integrated Auth"
$radWIA.Location = New-Object System.Drawing.Point(390, 65)
$radWIA.Size     = New-Object System.Drawing.Size(205, 20)
$radWIA.Checked  = $false
$radWIA.TabIndex = 3
$form.Controls.Add($radWIA)
#endregion

#region --- Buttons (top-right) ---
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text     = "Test"
$btnTest.Location = New-Object System.Drawing.Point(620, 62)
$btnTest.Size     = New-Object System.Drawing.Size(68, 26)
$btnTest.TabIndex = 9
$form.Controls.Add($btnTest)
$form.AcceptButton = $btnTest

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(696, 62)
$btnCancel.Size     = New-Object System.Drawing.Size(68, 26)
$btnCancel.Enabled  = $false
$btnCancel.TabIndex = 10
$form.Controls.Add($btnCancel)
#endregion

#region --- Modern Auth sub-options (Panel keeps radDCF/radACF exclusive with each other only) ---
$pnlModernSub = New-Object System.Windows.Forms.Panel
$pnlModernSub.Location = New-Object System.Drawing.Point(22, 89)
$pnlModernSub.Size     = New-Object System.Drawing.Size(358, 42)
$form.Controls.Add($pnlModernSub)

$radDCF = New-Object System.Windows.Forms.RadioButton
$radDCF.Text     = "Public Office app and Device Code Flow"
$radDCF.Location = New-Object System.Drawing.Point(2, 1)
$radDCF.Size     = New-Object System.Drawing.Size(354, 18)
$radDCF.Checked  = $true
$radDCF.TabIndex = 0
$pnlModernSub.Controls.Add($radDCF)

$radACF = New-Object System.Windows.Forms.RadioButton
$radACF.Text     = "Own 'Exchange Tester' app and Auth Code Flow"
$radACF.Location = New-Object System.Drawing.Point(2, 22)
$radACF.Size     = New-Object System.Drawing.Size(354, 18)
$radACF.Checked  = $false
$radACF.TabIndex = 1
$pnlModernSub.Controls.Add($radACF)
#endregion

#region --- WIA sub-option ---
$chkUseCurrentUser = New-Object System.Windows.Forms.CheckBox
$chkUseCurrentUser.Text     = "Use logged-in user"
$chkUseCurrentUser.Location = New-Object System.Drawing.Point(408, 91)
$chkUseCurrentUser.Size     = New-Object System.Drawing.Size(175, 20)
$chkUseCurrentUser.Checked  = $true
$chkUseCurrentUser.Enabled  = $false
$chkUseCurrentUser.TabIndex = 4
$form.Controls.Add($chkUseCurrentUser)
#endregion

#region --- Standalone options ---
$chkIgnoreCert = New-Object System.Windows.Forms.CheckBox
$chkIgnoreCert.Text     = "Ignore certificate errors"
$chkIgnoreCert.Location = New-Object System.Drawing.Point(8, 133)
$chkIgnoreCert.Size     = New-Object System.Drawing.Size(200, 20)
$chkIgnoreCert.Checked  = $false
$chkIgnoreCert.TabIndex = 5
$form.Controls.Add($chkIgnoreCert)

$chkUseSCP = New-Object System.Windows.Forms.CheckBox
$chkUseSCP.Text     = "Use SCP (domain-joined)"
$chkUseSCP.Location = New-Object System.Drawing.Point(390, 133)
$chkUseSCP.Size     = New-Object System.Drawing.Size(200, 20)
$chkUseSCP.Checked  = $false
$chkUseSCP.TabIndex = 6
$form.Controls.Add($chkUseSCP)
#endregion

#region --- ACF fields: Client ID + Tenant ID (shown only when ACF selected) ---
$lblClientId = New-Object System.Windows.Forms.Label
$lblClientId.Text      = "Client ID:"
$lblClientId.Location  = New-Object System.Drawing.Point(22, 135)
$lblClientId.Size      = New-Object System.Drawing.Size(100, 20)
$lblClientId.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblClientId.Visible   = $false
$form.Controls.Add($lblClientId)

$txtClientId = New-Object System.Windows.Forms.TextBox
$txtClientId.Location  = New-Object System.Drawing.Point(126, 133)
$txtClientId.Size      = New-Object System.Drawing.Size(356, 22)
$txtClientId.Text      = ''
$txtClientId.Visible   = $false
$txtClientId.TabIndex  = 7
$txtClientId.Font      = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtClientId)

$btnCreateApp = New-Object System.Windows.Forms.Button
$btnCreateApp.Text     = "Register App"
$btnCreateApp.Location = New-Object System.Drawing.Point(486, 133)
$btnCreateApp.Size     = New-Object System.Drawing.Size(100, 22)
$btnCreateApp.Visible  = $false
$btnCreateApp.TabIndex = 99
$form.Controls.Add($btnCreateApp)

$lblTenantId = New-Object System.Windows.Forms.Label
$lblTenantId.Text      = "Tenant ID:"
$lblTenantId.Location  = New-Object System.Drawing.Point(22, 159)
$lblTenantId.Size      = New-Object System.Drawing.Size(100, 20)
$lblTenantId.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTenantId.Visible   = $false
$form.Controls.Add($lblTenantId)

$txtTenantId = New-Object System.Windows.Forms.TextBox
$txtTenantId.Location  = New-Object System.Drawing.Point(126, 157)
$txtTenantId.Size      = New-Object System.Drawing.Size(460, 22)
$txtTenantId.Text      = ''
$txtTenantId.Visible   = $false
$txtTenantId.TabIndex  = 8
$txtTenantId.Font      = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtTenantId)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($lblClientId, "Application (client) ID of your 'Exchange Tester' app registration in Azure AD.")
$toolTip.SetToolTip($txtClientId, "Application (client) ID of your 'Exchange Tester' app registration in Azure AD.")
$toolTip.SetToolTip($lblTenantId, "Directory (tenant) ID — GUID or domain, e.g. contoso.onmicrosoft.com.")
$toolTip.SetToolTip($txtTenantId, "Directory (tenant) ID — GUID or domain, e.g. contoso.onmicrosoft.com.")
#endregion

#region --- Separator + Progress bar (y adjusted dynamically by Update-Layout) ---
$pnlSep = New-Object System.Windows.Forms.Panel
$pnlSep.Location  = New-Object System.Drawing.Point(0, 156)
$pnlSep.Size      = New-Object System.Drawing.Size(775, 2)
$pnlSep.BackColor = [System.Drawing.SystemColors]::ControlDark
$form.Controls.Add($pnlSep)

$prgBar = New-Object System.Windows.Forms.ProgressBar
$prgBar.Location = New-Object System.Drawing.Point(8, 164)
$prgBar.Size     = New-Object System.Drawing.Size(757, 14)
$prgBar.Minimum  = 0
$prgBar.Maximum  = 100
$prgBar.Value    = 0
$form.Controls.Add($prgBar)
#endregion

#region --- TabControl ---
$tabCtrl = New-Object System.Windows.Forms.TabControl
$tabCtrl.Location = New-Object System.Drawing.Point(8, 183)
$tabCtrl.Size     = New-Object System.Drawing.Size(757, 403)
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

    # Device Code Flow — no redirect URI registration needed.
    # Shows a short code; user signs in via any browser at microsoft.com/devicelogin.
    $getToken = {
        param([string]$wwwAuthHeader)

        $authUri = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
        if ($wwwAuthHeader -match 'authorization_uri\s*=\s*"([^"]+)"') {
            $authUri = $Matches[1] -replace '/oauth2(?:/v2\.0)?/authorize.*', '/oauth2/v2.0/authorize'
        }
        $deviceCodeUrl = $authUri -replace '/authorize', '/devicecode'
        $tokenUrl      = $authUri -replace '/authorize', '/token'
        $clientId      = if ($sync.ClientId) { $sync.ClientId } else { 'd3590ed6-52b3-4102-aeff-aad2292ab01c' }
        $scope         = 'https://outlook.office365.com/.default offline_access'

        & $logLine "Modern Auth: requesting device code…"

        # Step 1 — get device code
        $dcBytes = [System.Text.Encoding]::UTF8.GetBytes(
            "client_id=$([Uri]::EscapeDataString($clientId))&scope=$([Uri]::EscapeDataString($scope))")
        $dcJson = $null
        try {
            $rq = [System.Net.HttpWebRequest]::Create($deviceCodeUrl)
            $rq.Method        = "POST"
            $rq.ContentType   = "application/x-www-form-urlencoded"
            $rq.ContentLength = $dcBytes.Length
            $rq.Timeout       = 15000
            $ss = $rq.GetRequestStream(); $ss.Write($dcBytes, 0, $dcBytes.Length); $ss.Close()
            $rp = $rq.GetResponse()
            $dcJson = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
            $rp.Close()
        } catch [System.Net.WebException] {
            $exT = $_.Exception
            if ($exT.Response) {
                try {
                    $ej = (New-Object System.IO.StreamReader($exT.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                    $exT.Response.Close()
                    & $logLine "Modern Auth: device code request failed — $($ej.error): $($ej.error_description)"
                } catch { & $logLine "Modern Auth: device code request failed — $($exT.Message)" }
            } else { & $logLine "Modern Auth: device code request failed — $($exT.Message)" }
            return $null
        } catch {
            & $logLine "Modern Auth: device code request error — $($_.Exception.Message)"
            return $null
        }

        $userCode   = $dcJson.user_code
        $deviceCode = $dcJson.device_code
        $verifyUri  = if ($dcJson.verification_uri) { $dcJson.verification_uri } else { $dcJson.verification_url }
        $pollSec    = [int]$dcJson.interval; if ($pollSec -lt 5) { $pollSec = 5 }

        & $logLine "Modern Auth: visit $verifyUri — enter code: $userCode"

        $sync.DeviceToken  = $null
        $sync.DeviceError  = $null
        $sync.DeviceCancel = $false

        # Step 2 — show dialog with code while polling
        $dcForm = New-Object System.Windows.Forms.Form
        $dcForm.Text            = "Sign in to Microsoft"
        $dcForm.Size            = New-Object System.Drawing.Size(440, 210)
        $dcForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $dcForm.MinimizeBox     = $false
        $dcForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

        $lbl1 = New-Object System.Windows.Forms.Label
        $lbl1.Text     = "1.  Open a browser and go to:"
        $lbl1.Location = New-Object System.Drawing.Point(12, 14)
        $lbl1.AutoSize = $true

        $lnk = New-Object System.Windows.Forms.LinkLabel
        $lnk.Text     = $verifyUri
        $lnk.Location = New-Object System.Drawing.Point(28, 34)
        $lnk.AutoSize = $true
        $lnk.Add_LinkClicked({ [System.Diagnostics.Process]::Start($lnk.Text) })

        $lbl2 = New-Object System.Windows.Forms.Label
        $lbl2.Text     = "2.  Enter this code:"
        $lbl2.Location = New-Object System.Drawing.Point(12, 62)
        $lbl2.AutoSize = $true

        $lblCode = New-Object System.Windows.Forms.Label
        $lblCode.Text      = $userCode
        $lblCode.Font      = New-Object System.Drawing.Font("Consolas", 22, [System.Drawing.FontStyle]::Bold)
        $lblCode.Location  = New-Object System.Drawing.Point(28, 80)
        $lblCode.AutoSize  = $true
        $lblCode.ForeColor = [System.Drawing.Color]::DarkBlue

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text     = "Copy"
        $btnCopy.Location = New-Object System.Drawing.Point(340, 82)
        $btnCopy.Size     = New-Object System.Drawing.Size(72, 26)
        $btnCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($userCode) })

        $lblWait = New-Object System.Windows.Forms.Label
        $lblWait.Text      = "Waiting for sign-in…"
        $lblWait.Location  = New-Object System.Drawing.Point(12, 144)
        $lblWait.AutoSize  = $true
        $lblWait.ForeColor = [System.Drawing.Color]::Gray

        $btnCancelDC = New-Object System.Windows.Forms.Button
        $btnCancelDC.Text     = "Cancel"
        $btnCancelDC.Location = New-Object System.Drawing.Point(340, 140)
        $btnCancelDC.Size     = New-Object System.Drawing.Size(72, 26)
        $btnCancelDC.Add_Click({ $sync.DeviceCancel = $true; $dcForm.Close() })

        $dcForm.Controls.AddRange(@($lbl1, $lnk, $lbl2, $lblCode, $btnCopy, $lblWait, $btnCancelDC))

        # Poll token endpoint on a timer; runs on UI thread so keep HTTP timeout short
        $pollTimer = New-Object System.Windows.Forms.Timer
        $pollTimer.Interval = $pollSec * 1000
        $pollTimer.Add_Tick({
            if ($sync.DeviceToken -or $sync.DeviceError -or $sync.DeviceCancel) { return }
            $pb = [System.Text.Encoding]::UTF8.GetBytes(
                "grant_type=urn:ietf:params:oauth:grant-type:device_code" +
                "&client_id=$([Uri]::EscapeDataString($clientId))" +
                "&device_code=$([Uri]::EscapeDataString($deviceCode))")
            try {
                $rq2 = [System.Net.HttpWebRequest]::Create($tokenUrl)
                $rq2.Method        = "POST"
                $rq2.ContentType   = "application/x-www-form-urlencoded"
                $rq2.ContentLength = $pb.Length
                $rq2.Timeout       = 4000
                $ss2 = $rq2.GetRequestStream(); $ss2.Write($pb, 0, $pb.Length); $ss2.Close()
                try {
                    $rp2  = $rq2.GetResponse()
                    $tokJ = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                    $rp2.Close()
                    if ($tokJ.access_token) {
                        $sync.DeviceToken = "Bearer $($tokJ.access_token)"
                        $dcForm.Close()
                    }
                } catch [System.Net.WebException] {
                    $ex2 = $_.Exception
                    if ($ex2.Response) {
                        $ej2 = (New-Object System.IO.StreamReader($ex2.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        $ex2.Response.Close()
                        switch ($ej2.error) {
                            'authorization_pending' {}
                            'slow_down'             { $pollTimer.Interval += 5000 }
                            default {
                                $sync.DeviceError = "$($ej2.error): $($ej2.error_description)"
                                $dcForm.Close()
                            }
                        }
                    }
                }
            } catch {}
        })

        $dcForm.Add_Shown({ $pollTimer.Start() })
        $dcForm.Add_FormClosed({ $pollTimer.Stop() })
        [void]$dcForm.ShowDialog()
        $pollTimer.Dispose()
        $dcForm.Dispose()

        if ($sync.DeviceCancel -or (-not $sync.DeviceToken -and -not $sync.DeviceError)) {
            & $logLine "Modern Auth: sign-in cancelled."
            return $null
        }
        if ($sync.DeviceError) {
            & $logLine "Modern Auth: sign-in error — $($sync.DeviceError)"
            $sync.DeviceError = $null
            return $null
        }
        & $logLine "Modern Auth: access token acquired."
        $tok = $sync.DeviceToken; $sync.DeviceToken = $null
        return $tok
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
        $rtbXml.SelectionStart = 0
        $rtbXml.ScrollToCaret()

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

# Adjusts visible/enabled state of all sub-controls based on current auth selection
function Update-AuthMode {
    $isModern = $radModernAuth.Checked
    $pnlModernSub.Enabled      = $isModern
    $chkUseCurrentUser.Enabled = -not $isModern
    $useExplicit = (-not $isModern) -and (-not $chkUseCurrentUser.Checked)
    $lblPass.Enabled = $useExplicit
    $txtPass.Enabled = $useExplicit
    if (-not $useExplicit) { $txtPass.Clear() }
    Update-Layout
}

# Repositions standalone options, separator, progress bar, and tab control
# based on whether ACF rows are visible
function Update-Layout {
    $showACF = $radACF.Checked -and $radModernAuth.Checked
    $lblClientId.Visible  = $showACF
    $txtClientId.Visible  = $showACF
    $btnCreateApp.Visible = $showACF
    $lblTenantId.Visible  = $showACF
    $txtTenantId.Visible  = $showACF
    $y = if ($showACF) { 179 } else { 133 }
    $chkIgnoreCert.Top = $y
    $chkUseSCP.Top     = $y
    $pnlSep.Top        = $y + 23
    $prgBar.Top        = $y + 31
    $tabCtrl.Top       = $y + 50
    $tabCtrl.Height    = $form.ClientSize.Height - ($y + 50) - 9
}

# Looks up the tenant ID for the domain in $txtEmail via the OIDC discovery endpoint.
# Runs synchronously with DoEvents so "Detecting…" is visible before the HTTP call.
# Triggered when ACF is selected or when the email field loses focus while ACF is active.
function Resolve-TenantId {
    if (-not $radACF.Checked) { return }
    $email = $txtEmail.Text.Trim()
    if ($email -notmatch '@([^@\s]+)$') { return }
    $domain = $Matches[1]
    $txtTenantId.Text    = "Detecting…"
    $txtTenantId.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $tid = $null
    try {
        $req = [System.Net.HttpWebRequest]::Create(
            "https://login.microsoftonline.com/$([Uri]::EscapeDataString($domain))/.well-known/openid-configuration")
        $req.Method  = "GET"
        $req.Timeout = 8000
        $rp = $req.GetResponse()
        $j  = (New-Object System.IO.StreamReader($rp.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp.Close()
        if ($j.token_endpoint -match '/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/') {
            $tid = $Matches[1]
        }
    } catch {}
    $txtTenantId.Text    = if ($tid) { $tid } else { '' }
    $txtTenantId.Enabled = $true
    $form.Cursor         = [System.Windows.Forms.Cursors]::Default
}

$radModernAuth.Add_CheckedChanged({ if ($radModernAuth.Checked) { Update-AuthMode } })
$radWIA.Add_CheckedChanged({        if ($radWIA.Checked)        { Update-AuthMode } })
$radDCF.Add_CheckedChanged({        if ($radDCF.Checked)        { Update-Layout  } })
$radACF.Add_CheckedChanged({        if ($radACF.Checked)        { Update-Layout; Resolve-TenantId } })
$txtEmail.Add_Leave({ Resolve-TenantId })

$chkUseCurrentUser.Add_CheckedChanged({
    $useExplicit     = -not $chkUseCurrentUser.Checked
    $lblPass.Enabled = $useExplicit
    $txtPass.Enabled = $useExplicit
    if (-not $useExplicit) { $txtPass.Clear() }
})

# Config file path (same directory as this script)
$_scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path }
$script:configPath = Join-Path $_scriptDir "ExchangeTester.config"

$btnCreateApp.Add_Click({
    $tenantId = $txtTenantId.Text.Trim()
    if (-not $tenantId) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please detect or enter the Tenant ID first.",
            "Tenant ID Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # ── Auth Code Flow + PKCE via local HTTP listener ────────────────────────
    $graphApp = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools

    # PKCE: 48-byte random verifier → SHA-256 → base64url challenge
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
    $rngBytes = [byte[]]::new(48)
    $rng.GetBytes($rngBytes); $rng.Dispose()
    $verifier  = [Convert]::ToBase64String($rngBytes) -replace '\+','-' -replace '/','_' -replace '=',''
    $sha256    = [System.Security.Cryptography.SHA256Managed]::new()
    $chalBytes = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier)); $sha256.Dispose()
    $challenge = [Convert]::ToBase64String($chalBytes) -replace '\+','-' -replace '/','_' -replace '=',''

    # Find a free ephemeral port
    $tcpT = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $tcpT.Start(); $freePort = $tcpT.LocalEndpoint.Port; $tcpT.Stop()

    $redirectUri = "http://localhost:$freePort/"
    $graphScope  = 'https://graph.microsoft.com/Application.ReadWrite.All ' +
                   'https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All ' +
                   'https://graph.microsoft.com/Directory.ReadWrite.All ' +
                   'https://graph.microsoft.com/Directory.AccessAsUser.All ' +
                   'offline_access'
    $stateVal    = [Guid]::NewGuid().ToString('N')
    $authUrl     = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenantId))/oauth2/v2.0/authorize" +
        "?client_id=$([Uri]::EscapeDataString($graphApp))" +
        "&response_type=code" +
        "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
        "&scope=$([Uri]::EscapeDataString($graphScope))" +
        "&state=$([Uri]::EscapeDataString($stateVal))" +
        "&code_challenge=$([Uri]::EscapeDataString($challenge))" +
        "&code_challenge_method=S256" +
        "&prompt=select_account"

    # Synchronized hashtable for listener ↔ UI thread communication
    $lsync = [hashtable]::Synchronized(@{ Code = $null; Err = $null; Cancel = $false; Done = $false })

    # Background runspace: listens for the OAuth2 redirect callback
    $lRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $lRs.ApartmentState = [System.Threading.ApartmentState]::STA
    $lRs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $lRs.Open()
    $lRs.SessionStateProxy.SetVariable('lsync',    $lsync)
    $lRs.SessionStateProxy.SetVariable('freePort', $freePort)
    $lRs.SessionStateProxy.SetVariable('stateVal', $stateVal)

    $lPs = [System.Management.Automation.PowerShell]::Create()
    $lPs.Runspace = $lRs
    [void]$lPs.AddScript({
        $hl = [System.Net.HttpListener]::new()
        $hl.Prefixes.Add("http://localhost:$freePort/")
        $hl.Start()
        try {
            $ar       = $hl.BeginGetContext($null, $null)
            $deadline = (Get-Date).AddMinutes(5)
            while (-not $ar.IsCompleted -and (Get-Date) -lt $deadline) {
                if ($lsync.Cancel) { $hl.Stop(); return }
                [System.Threading.Thread]::Sleep(200)
            }
            if (-not $ar.IsCompleted) {
                $hl.Stop()
                if (-not $lsync.Cancel) { $lsync.Err = 'Sign-in timed out (5 min)' }
                return
            }
            $ctx  = $hl.EndGetContext($ar)
            $qs   = $ctx.Request.QueryString
            $code = $qs['code']; $retState = $qs['state']
            $oErr = $qs['error']; $oErrD = $qs['error_description']
            $ok   = $code -and $retState -eq $stateVal
            $html = if ($ok) {
                '<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:green">&#10003; Signed in successfully</h2><p>You may close this tab and return to Exchange Tester.</p></body></html>'
            } else {
                '<html><body style="font-family:sans-serif;padding:40px"><h2 style="color:red">&#10007; Sign-in failed or cancelled</h2><p>You may close this tab.</p></body></html>'
            }
            $hb = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType = 'text/html; charset=utf-8'
            $ctx.Response.ContentLength64 = $hb.Length
            $ctx.Response.OutputStream.Write($hb, 0, $hb.Length)
            $ctx.Response.OutputStream.Close()
            $ctx.Response.Close()
            if ($ok)       { $lsync.Code = $code }
            elseif ($oErr) { $lsync.Err  = if ($oErrD) { "${oErr}: $oErrD" } else { $oErr } }
            else           { $lsync.Err  = 'No authorization code received' }
        } catch {
            if (-not $lsync.Cancel) { $lsync.Err = $_.Exception.Message }
        } finally {
            try { $hl.Stop(); $hl.Close() } catch {}
            $lsync.Done = $true
        }
    })
    [void]$lPs.BeginInvoke()

    # Open the system browser at the auth URL
    [System.Diagnostics.Process]::Start($authUrl) | Out-Null

    # ── "Waiting for browser sign-in" dialog ───────────────────────────────────
    $aForm = New-Object System.Windows.Forms.Form
    $aForm.Text            = "Sign in  —  Create App Registration"
    $aForm.Size            = New-Object System.Drawing.Size(520, 155)
    $aForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $aForm.MinimizeBox     = $false
    $aForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    $aL0 = New-Object System.Windows.Forms.Label
    $aL0.Text     = "Sign in as Application Administrator or Global Admin."
    $aL0.Location = New-Object System.Drawing.Point(12, 12)
    $aL0.Size     = New-Object System.Drawing.Size(490, 18)
    $aL0.Font     = New-Object System.Drawing.Font($aForm.Font, [System.Drawing.FontStyle]::Bold)

    $aWait = New-Object System.Windows.Forms.Label
    $aWait.Text      = "The sign-in page has been opened in your browser. Waiting for authentication…"
    $aWait.Location  = New-Object System.Drawing.Point(12, 38)
    $aWait.Size      = New-Object System.Drawing.Size(490, 18)
    $aWait.ForeColor = [System.Drawing.Color]::Gray

    $aDots = New-Object System.Windows.Forms.Label
    $aDots.Text      = ""
    $aDots.Location  = New-Object System.Drawing.Point(12, 62)
    $aDots.AutoSize  = $true
    $aDots.ForeColor = [System.Drawing.Color]::SteelBlue

    $aCancel = New-Object System.Windows.Forms.Button
    $aCancel.Text     = "Cancel"
    $aCancel.Location = New-Object System.Drawing.Point(430, 90)
    $aCancel.Size     = New-Object System.Drawing.Size(72, 26)
    $aCancel.Add_Click({ $lsync.Cancel = $true; $aForm.Close() })

    $aForm.Controls.AddRange(@($aL0, $aWait, $aDots, $aCancel))

    $aPoll = New-Object System.Windows.Forms.Timer; $aPoll.Interval = 300
    $aPoll.Add_Tick({
        if ($lsync.Done -or $lsync.Cancel) { $aForm.Close(); return }
        $lsync.DotCount = (($lsync.DotCount + 1) % 6)
        $aDots.Text = '.' * ($lsync.DotCount + 1)
    })
    $aForm.Add_Shown({ $lsync.DotCount = 0; $aPoll.Start() })
    $aForm.Add_FormClosed({ $aPoll.Stop() })
    [void]$aForm.ShowDialog()
    $aPoll.Dispose(); $aForm.Dispose()

    # Clean up listener runspace
    try { $lPs.Stop() } catch {}
    $lPs.Dispose(); $lRs.Close(); $lRs.Dispose()

    if ($lsync.Cancel -or (-not $lsync.Code -and -not $lsync.Err)) { return }
    if ($lsync.Err) {
        [System.Windows.Forms.MessageBox]::Show("Sign-in error: $($lsync.Err)", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # ── Exchange authorization code for access token ───────────────────────────
    $tokUrl   = "https://login.microsoftonline.com/$([Uri]::EscapeDataString($tenantId))/oauth2/v2.0/token"
    $tokBody  = "grant_type=authorization_code" +
        "&client_id=$([Uri]::EscapeDataString($graphApp))" +
        "&code=$([Uri]::EscapeDataString($lsync.Code))" +
        "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
        "&code_verifier=$([Uri]::EscapeDataString($verifier))"
    $tokBytes = [System.Text.Encoding]::UTF8.GetBytes($tokBody)
    $adminToken = $null
    try {
        $rq2 = [System.Net.HttpWebRequest]::Create($tokUrl)
        $rq2.Method = "POST"; $rq2.ContentType = "application/x-www-form-urlencoded"
        $rq2.ContentLength = $tokBytes.Length; $rq2.Timeout = 15000
        $ss2 = $rq2.GetRequestStream(); $ss2.Write($tokBytes, 0, $tokBytes.Length); $ss2.Close()
        $rp2 = $rq2.GetResponse()
        $tJ  = (New-Object System.IO.StreamReader($rp2.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp2.Close()
        $adminToken = $tJ.access_token
    } catch [System.Net.WebException] {
        $exT = $_.Exception
        $msg = if ($exT.Response) {
            try { $ej=(New-Object System.IO.StreamReader($exT.Response.GetResponseStream())).ReadToEnd()|ConvertFrom-Json; $exT.Response.Close(); "$($ej.error): $($ej.error_description)" } catch { $exT.Message }
        } else { $exT.Message }
        [System.Windows.Forms.MessageBox]::Show("Token exchange failed: $msg", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Token exchange failed: $($_.Exception.Message)", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if (-not $adminToken) {
        [System.Windows.Forms.MessageBox]::Show("Token exchange returned no access token.", "Auth Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # ── Create app registration via Microsoft Graph ───────────────────────────
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    $appBody  = '{"displayName":"Exchange Tester","isFallbackPublicClient":true}'
    $appBytes = [System.Text.Encoding]::UTF8.GetBytes($appBody)
    $newId    = $null
    try {
        $rq3 = [System.Net.HttpWebRequest]::Create("https://graph.microsoft.com/v1.0/applications")
        $rq3.Method = "POST"; $rq3.ContentType = "application/json"
        $rq3.ContentLength = $appBytes.Length; $rq3.Timeout = 20000
        $rq3.Headers["Authorization"] = "Bearer $adminToken"
        $ss3 = $rq3.GetRequestStream(); $ss3.Write($appBytes, 0, $appBytes.Length); $ss3.Close()
        $rp3  = $rq3.GetResponse()
        $aJ   = (New-Object System.IO.StreamReader($rp3.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        $rp3.Close()
        $newId = $aJ.appId
    } catch [System.Net.WebException] {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $exT3 = $_.Exception
        $m3 = if ($exT3.Response) {
            try { $e3=(New-Object System.IO.StreamReader($exT3.Response.GetResponseStream())).ReadToEnd()|ConvertFrom-Json; $exT3.Response.Close(); "$($e3.error.code): $($e3.error.message)" } catch { $exT3.Message }
        } else { $exT3.Message }
        [System.Windows.Forms.MessageBox]::Show($m3, "App Registration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    } catch {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    # ── Save config and update UI ─────────────────────────────────────────────
    try {
        [System.IO.File]::WriteAllText($script:configPath,
            (ConvertTo-Json @{ ClientId = $newId }),
            [System.Text.Encoding]::UTF8)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "App created but config could not be saved:`n$($_.Exception.Message)",
            "Warning", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }

    $txtClientId.Text = $newId
    [System.Windows.Forms.MessageBox]::Show(
        "App registration 'Exchange Tester' created successfully.`n`nClient ID:`n$newId`n`nSaved to ExchangeTester.config.",
        "App Registration Created",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
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

    if ($radModernAuth.Checked -and $radACF.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Authorization Code Flow is not yet implemented.`nPlease use 'Public Office app and Device Code Flow' instead.",
            "Not Implemented",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
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
        UseWindowsAuth = $radWIA.Checked -and $chkUseCurrentUser.Checked
        ModernAuth     = $radModernAuth.Checked
        UseDeviceCode  = $radDCF.Checked
        ClientId       = $txtClientId.Text.Trim()
        TenantId       = $txtTenantId.Text.Trim()
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

# Load saved config (Client ID from a previous "Register App")
if (Test-Path $script:configPath) {
    try {
        $cfg = Get-Content $script:configPath -Raw | ConvertFrom-Json
        if ($cfg.ClientId) { $txtClientId.Text = $cfg.ClientId }
    } catch {}
}

# Detect domain-joined status and set SCP checkbox accordingly
$chkUseSCP.Checked = $false
try {
    [void][System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
    $chkUseSCP.Checked = $true
} catch {}

# Pre-fill e-mail with the logged-in user's UPN (multiple fallback methods)
$upn = ''

# Method 1: WindowsIdentity UPN claim (works on domain-joined with Kerberos)
if (-not $upn) {
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $claim = $id.Claims | Where-Object {
            $_.Type -eq 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn'
        } | Select-Object -First 1
        if ($claim -and $claim.Value) { $upn = $claim.Value }
    } catch {}
}

# Method 2: USERPRINCIPALNAME environment variable (set by some SSO/MDM solutions)
if (-not $upn) {
    try {
        if ($env:USERPRINCIPALNAME -and $env:USERPRINCIPALNAME -match '@') {
            $upn = $env:USERPRINCIPALNAME
        }
    } catch {}
}

# Method 3: ADSI lookup in Active Directory (domain-joined machines)
if (-not $upn) {
    try {
        $searcher = [adsisearcher]"samaccountname=$env:USERNAME"
        [void]$searcher.PropertiesToLoad.Add("userprincipalname")
        $result = $searcher.FindOne()
        if ($result) {
            $val = $result.Properties['userprincipalname']
            if ($val -and $val.Count -gt 0 -and $val[0] -match '@') { $upn = $val[0] }
        }
    } catch {}
}

# Method 4: whoami /upn (works on domain-joined, may be slow on non-domain machines)
if (-not $upn) {
    try {
        $w = & whoami.exe /upn 2>$null
        if ($w -and $w.Trim() -match '@') { $upn = $w.Trim() }
    } catch {}
}

if ($upn) { $txtEmail.Text = $upn }

[System.Windows.Forms.Application]::Run($form)
