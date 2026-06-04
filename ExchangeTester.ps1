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

#endregion ===================================================================
#  CONTROL INTERACTIONS (UI-only, no test logic yet)
#==============================================================================

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

    # Reset UI for new test run
    $rtbLog.Clear()
    $rtbXml.Clear()
    $lvwResults.Items.Clear()
    $prgBar.Value       = 0
    $btnTest.Enabled    = $false
    $btnCancel.Enabled  = $true
    $tabCtrl.SelectedTab = $tabLog

    # ---- logic will be wired here in Milestone 2 ----
    $rtbLog.AppendText("(Test logic not yet implemented - coming in Milestone 2)`r`n")
    $btnTest.Enabled   = $true
    $btnCancel.Enabled = $false
})

$btnCancel.Add_Click({
    # ---- BackgroundWorker cancel will be wired in Milestone 2 ----
    $btnCancel.Enabled = $false
})

$form.Add_FormClosing({
    # ---- worker cleanup will be added in Milestone 2 ----
})

#==============================================================================
#  START
#==============================================================================
[System.Windows.Forms.Application]::Run($form)
