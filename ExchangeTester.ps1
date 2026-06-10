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

#region --- Row 3: Checkboxes ---
$chkWinAuth = New-Object System.Windows.Forms.CheckBox
$chkWinAuth.Text      = "Use AutoDiscover"
$chkWinAuth.Location  = New-Object System.Drawing.Point(8, 70)
$chkWinAuth.Size      = New-Object System.Drawing.Size(155, 20)
$chkWinAuth.Checked   = $true
$chkWinAuth.TabIndex  = 2
$form.Controls.Add($chkWinAuth)

$chkUseCurrentUser = New-Object System.Windows.Forms.CheckBox
$chkUseCurrentUser.Text     = "Use logged-in user (Windows Auth)"
$chkUseCurrentUser.Location = New-Object System.Drawing.Point(172, 70)
$chkUseCurrentUser.Size     = New-Object System.Drawing.Size(240, 20)
$chkUseCurrentUser.Checked  = $true
$chkUseCurrentUser.TabIndex = 3
$form.Controls.Add($chkUseCurrentUser)

$chkIgnoreCert = New-Object System.Windows.Forms.CheckBox
$chkIgnoreCert.Text     = "Ignore certificate errors"
$chkIgnoreCert.Location = New-Object System.Drawing.Point(420, 70)
$chkIgnoreCert.Size     = New-Object System.Drawing.Size(195, 20)
$chkIgnoreCert.Checked  = $false
$chkIgnoreCert.TabIndex = 4
$form.Controls.Add($chkIgnoreCert)
#endregion

#region --- Buttons (top-right) ---
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text     = "Test"
$btnTest.Location = New-Object System.Drawing.Point(620, 64)
$btnTest.Size     = New-Object System.Drawing.Size(68, 26)
$btnTest.TabIndex = 5
$form.Controls.Add($btnTest)
$form.AcceptButton = $btnTest

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(696, 64)
$btnCancel.Size     = New-Object System.Drawing.Size(68, 26)
$btnCancel.Enabled  = $false
$btnCancel.TabIndex = 6
$form.Controls.Add($btnCancel)
#endregion

#region --- Separator + Progress bar ---
$pnlSep = New-Object System.Windows.Forms.Panel
$pnlSep.Location  = New-Object System.Drawing.Point(0, 96)
$pnlSep.Size      = New-Object System.Drawing.Size(775, 2)
$pnlSep.BackColor = [System.Drawing.SystemColors]::ControlDark
$form.Controls.Add($pnlSep)

$prgBar = New-Object System.Windows.Forms.ProgressBar
$prgBar.Location = New-Object System.Drawing.Point(8, 104)
$prgBar.Size     = New-Object System.Drawing.Size(757, 14)
$prgBar.Minimum  = 0
$prgBar.Maximum  = 100
$prgBar.Value    = 0
$form.Controls.Add($prgBar)
#endregion

#region --- TabControl ---
$tabCtrl = New-Object System.Windows.Forms.TabControl
$tabCtrl.Location = New-Object System.Drawing.Point(8, 124)
$tabCtrl.Size     = New-Object System.Drawing.Size(757, 462)
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
#  BACKGROUND WORKER
#==============================================================================

# Capture the main runspace so script blocks can be invoked on the ThreadPool thread
$script:MainRunspace = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace

$bgWorker = New-Object System.ComponentModel.BackgroundWorker
$bgWorker.WorkerReportsProgress     = $true
$bgWorker.WorkerSupportsCancellation = $true

$bgWorker.Add_DoWork({
    param($bwSender, $bwArgs)

    # Make PowerShell script block invocations work on this ThreadPool thread
    [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace = $script:MainRunspace

    $a          = $bwArgs.Argument
    $email      = $a.Email
    $password   = $a.Password
    $useWinAuth = $a.UseWindowsAuth
    $domain     = ($email -split '@')[1]

    # Explicit credentials (Basic / NTLM with supplied password)
    $netCred = $null
    if (-not $useWinAuth -and $password -ne '') {
        $netCred = New-Object System.Net.NetworkCredential($email, $password)
    }

    # AutoDiscover POST body
    $bodyXml = @"
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
  <Request>
    <EMailAddress>$email</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
  </Request>
</Autodiscover>
"@
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyXml)

    #region --- helpers (script blocks, capture outer scope via PS scope chain) ---

    $logLine = {
        param([string]$msg)
        $bwSender.ReportProgress(0, [PSCustomObject]@{ Type = 'Log'; Text = $msg })
    }

    $setPct = {
        param([int]$pct)
        $bwSender.ReportProgress($pct, [PSCustomObject]@{ Type = 'Pct' })
    }

    # HTTP POST → @{Code; Body; Location; Error}
    $doPost = {
        param([string]$url)
        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method            = "POST"
            $req.ContentType       = "text/xml; charset=utf-8"
            $req.ContentLength     = $bodyBytes.Length
            $req.AllowAutoRedirect = $false
            $req.Timeout           = 30000
            $req.UserAgent         = "Microsoft Office/16.0 (Windows NT 10.0)"
            if ($useWinAuth) {
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
                return @{ Code = $code; Body = $body; Location = $null; Error = $null }
            } catch [System.Net.WebException] {
                $ex = $_.Exception
                if ($ex.Response) {
                    $code = [int]$ex.Response.StatusCode
                    $loc  = $ex.Response.Headers["Location"]
                    $ex.Response.Close()
                    return @{ Code = $code; Body = $null; Location = $loc; Error = $null }
                }
                return @{ Code = -1; Body = $null; Location = $null; Error = $ex.Message }
            }
        } catch {
            return @{ Code = -1; Body = $null; Location = $null; Error = $_.Exception.Message }
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
            & $logLine "AutoDiscover via $url failed (0x800C820E)."
            return $null
        }

        & $logLine "AutoDiscover via $url failed (httpStatus=$($res.Code))."
        return $null
    }

    #endregion

    $foundXml = $null

    # --- Step 1: O365 ---
    & $setPct 10
    if (-not $bwSender.CancellationPending) {
        $foundXml = & $tryUrl "https://outlook.office365.com/autodiscover/autodiscover.xml"
    }

    # --- Step 2: https://<domain>/autodiscover/autodiscover.xml ---
    & $setPct 30
    if (-not $foundXml -and -not $bwSender.CancellationPending) {
        $foundXml = & $tryUrl "https://$domain/autodiscover/autodiscover.xml"
    }

    # --- Step 3: https://autodiscover.<domain>/autodiscover/autodiscover.xml ---
    & $setPct 50
    if (-not $foundXml -and -not $bwSender.CancellationPending) {
        $foundXml = & $tryUrl "https://autodiscover.$domain/autodiscover/autodiscover.xml"
    }

    # --- Step 4: SCP (Active Directory Service Connection Point) ---
    & $setPct 65
    if (-not $foundXml -and -not $bwSender.CancellationPending) {
        & $logLine "Local AutoDiscover for $domain starting."
        try {
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.Filter = "(&(objectClass=serviceConnectionPoint)(|(serviceBindingInformation=*autodiscover*)(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)))"
            [void]$searcher.PropertiesToLoad.Add("serviceBindingInformation")
            $scpHits = $searcher.FindAll()
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
            & $logLine "Local AutoDiscover for $domain failed (0x8004010F)."
        }
    }

    # --- Step 5: HTTP redirect check (well-known URL) ---
    & $setPct 78
    if (-not $foundXml -and -not $bwSender.CancellationPending) {
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
    & $setPct 90
    if (-not $foundXml -and -not $bwSender.CancellationPending) {
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

    if ($bwSender.CancellationPending) { $bwArgs.Cancel = $true; return }

    & $setPct 100
    $bwArgs.Result = [PSCustomObject]@{
        Xml     = $foundXml
        Success = ($null -ne $foundXml)
    }
})

$bgWorker.Add_ProgressChanged({
    param($sender, $e)
    $state = $e.UserState
    if ($state.Type -eq 'Log') {
        $rtbLog.AppendText("$($state.Text)`r`n")
        $rtbLog.ScrollToCaret()
    } elseif ($state.Type -eq 'Pct') {
        $v = $e.ProgressPercentage
        if ($v -ge 0 -and $v -le 100) { $prgBar.Value = $v }
    }
})

$bgWorker.Add_RunWorkerCompleted({
    param($sender, $e)
    $btnTest.Enabled   = $true
    $btnCancel.Enabled = $false

    if ($e.Cancelled) {
        $form.Text = "Test E-Mail AutoConfiguration"
        $rtbLog.AppendText("`r`nTest cancelled.`r`n")
        return
    }
    if ($e.Error) {
        $form.Text = "Test E-Mail AutoConfiguration  —  Error"
        $rtbLog.AppendText("`r`nUnhandled error: $($e.Error.Message)`r`n")
        return
    }

    $res = $e.Result

    if ($res.Xml) {
        # --- XML tab: pretty-print ---
        try {
            $xd = New-Object System.Xml.XmlDocument
            $xd.LoadXml($res.Xml)
            $sb = New-Object System.Text.StringBuilder
            $sw = New-Object System.IO.StringWriter($sb)
            $xw = New-Object System.Xml.XmlTextWriter($sw)
            $xw.Formatting  = [System.Xml.Formatting]::Indented
            $xw.Indentation = 2
            $xd.WriteTo($xw)
            $xw.Flush()
            $rtbXml.Text = $sb.ToString()
        } catch {
            $rtbXml.Text = $res.Xml
        }

        # --- Results tab: parsed rows grouped by section ---
        $rows = ConvertFrom-AutodiscoverXml -RawXml $res.Xml
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

        # Switch to Results if we got data, otherwise XML
        if ($rows.Count -gt 0) {
            $tabCtrl.SelectedTab = $tabResults
        } else {
            $tabCtrl.SelectedTab = $tabXml
        }
    } else {
        $form.Text = "Test E-Mail AutoConfiguration  —  No configuration found"
        $rtbLog.AppendText("`r`nAutoDiscover failed for all tested methods.`r`n")
        $tabCtrl.SelectedTab = $tabLog
    }
})

#endregion ===================================================================
#  CONTROL INTERACTIONS
#==============================================================================

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

    $bgWorker.RunWorkerAsync([PSCustomObject]@{
        Email          = $email
        Password       = $txtPass.Text
        UseWindowsAuth = $chkUseCurrentUser.Checked
    })
})

$btnCancel.Add_Click({
    $bgWorker.CancelAsync()
    $btnCancel.Enabled = $false
    $rtbLog.AppendText("Cancelling...`r`n")
})

$form.Add_FormClosing({
    if ($bgWorker.IsBusy) { $bgWorker.CancelAsync() }
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
})

#==============================================================================
#  START
#==============================================================================
[System.Windows.Forms.Application]::Run($form)
