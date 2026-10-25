<# : batch script
@echo off
net session >nul 2>&1 || (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' -Raw)"
exit /b
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------
# Edition list (Product ID | Label) based on ODT Reference
# ------------------------------------------------------------
$editionData = @"
O365ProPlusRetail|Microsoft 365 Apps for enterprise
O365ProPlusEEANoTeamsRetail|Microsoft 365 Apps for enterprise - EEA No Teams
O365BusinessRetail|Microsoft 365 Apps for business
O365BusinessEEANoTeamsRetail|Microsoft 365 Apps for business - EEA No Teams
O365HomePremRetail|Microsoft 365 Family / Personal
VisioProRetail|Visio Professional (Subscription)
ProjectProRetail|Project Professional (Subscription)
AccessRuntimeRetail|Access Runtime only
LanguagePack|Language packs only
Professional2021Retail|Office 2021 Professional (Retail)
Professional2024Retail|Office 2024 Professional (Retail)
ProPlus2021Retail|Office 2021 Professional Plus (Retail)
ProPlus2021Volume|Office 2021 Professional Plus (Volume)
ProPlusSPLA2021Volume|Office 2021 Professional Plus (SPLA)
ProPlus2024Retail|Office 2024 Professional Plus (Retail)
ProPlus2024Volume|Office 2024 Professional Plus (Volume)
Standard2021Volume|Office 2021 Standard (Volume)
StandardSPLA2021Volume|Office 2021 Standard (SPLA)
Standard2024Volume|Office 2024 Standard (Volume)
HomeBusiness2021Retail|Office 2021 Home and Business
HomeBusiness2024Retail|Office 2024 Home and Business
HomeStudent2021Retail|Office 2021 Home and Student
Home2024Retail|Office 2024 Home
Personal2021Retail|Office 2021 Personal
Access2021Retail|Access 2021
Access2024Retail|Access 2024
Excel2021Retail|Excel 2021
Excel2024Retail|Excel 2024
Word2021Retail|Word 2021
Word2024Retail|Word 2024
PowerPoint2021Retail|PowerPoint 2021
PowerPoint2024Retail|PowerPoint 2024
Outlook2021Retail|Outlook 2021
Outlook2024Retail|Outlook 2024
Publisher2021Retail|Publisher 2021
VisioPro2021Retail|Visio Professional 2021
VisioPro2024Retail|Visio Professional 2024
VisioStd2021Retail|Visio Standard 2021
VisioStd2024Retail|Visio Standard 2024
ProjectPro2021Retail|Project Professional 2021
ProjectPro2024Retail|Project Professional 2024
ProjectStd2021Retail|Project Standard 2021
ProjectStd2024Retail|Project Standard 2024
"@

$editions = [ordered]@{}
foreach ($line in ($editionData -split "`r?`n")) {
    if (-not $line.Trim()) { continue }
    $parts = $line.Split('|')
    $display = "{0}   [{1}]" -f $parts[1].Trim(), $parts[0].Trim()
    $editions[$display] = $parts[0].Trim()
}

$channels = [ordered]@{
    "Current Channel (Recommended)"       = "Current"
    "Monthly Enterprise Channel"          = "MonthlyEnterprise"
    "Semi-Annual Enterprise Channel"      = "SemiAnnual"
    "Current Preview"                     = "CurrentPreview"
    "Semi-Annual Preview"                 = "SemiAnnualPreview"
    "Beta Channel"                        = "BetaChannel"
    "Perpetual LTSC 2024 (VL)"            = "PerpetualVL2024"
    "Perpetual LTSC 2021 (VL)"            = "PerpetualVL2021"
    "Perpetual 2019 (VL)"                 = "PerpetualVL2019"
}

$languages = [ordered]@{
    "English (US)"           = "en-us"
    "English (UK)"           = "en-gb"
    "Arabic (Saudi Arabia)"  = "ar-sa"
    "Bulgarian"              = "bg-bg"
    "Czech"                  = "cs-cz"
    "Danish"                 = "da-dk"
    "German"                 = "de-de"
    "Greek"                  = "el-gr"
    "Spanish (Spain)"        = "es-es"
    "Spanish (Mexico)"       = "es-mx"
    "Estonian"               = "et-ee"
    "Finnish"                = "fi-fi"
    "French (France)"        = "fr-fr"
    "French (Canada)"        = "fr-ca"
    "Hebrew"                 = "he-il"
    "Hindi"                  = "hi-in"
    "Italian"                = "it-it"
    "Japanese"               = "ja-jp"
    "Korean"                 = "ko-kr"
    "Dutch"                  = "nl-nl"
    "Polish"                 = "pl-pl"
    "Portuguese (Brazil)"    = "pt-br"
    "Portuguese (Portugal)"  = "pt-pt"
    "Russian"                = "ru-ru"
    "Swedish"                = "sv-se"
    "Thai"                   = "th-th"
    "Turkish"                = "tr-tr"
    "Ukrainian"              = "uk-ua"
    "Vietnamese"             = "vi-vn"
    "Chinese (Simplified)"   = "zh-cn"
    "Chinese (Traditional)"  = "zh-tw"
    "Match Windows language" = "MatchOS"
}

# ------------------------------------------------------------
# GUI Window Configuration
# ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Office Deployment & Activation"
$form.Size = New-Object System.Drawing.Size(560, 830)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Header Title
$title = New-Object System.Windows.Forms.Label
$title.Text = "Microsoft Office Setup Assistant"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$title.Size = New-Object System.Drawing.Size(520, 25)
$title.Location = New-Object System.Drawing.Point(15, 10)
$form.Controls.Add($title)

# Task Selection Box
$grpChoice = New-Object System.Windows.Forms.GroupBox
$grpChoice.Text = "Select Task"
$grpChoice.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpChoice.Location = New-Object System.Drawing.Point(20, 45)
$grpChoice.Size = New-Object System.Drawing.Size(500, 195)

$radNo = New-Object System.Windows.Forms.RadioButton
$radNo.Text = "Install Office"
$radNo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radNo.Location = New-Object System.Drawing.Point(15, 22)
$radNo.Size = New-Object System.Drawing.Size(460, 21)
$radNo.Checked = $true

$radChangeOffice = New-Object System.Windows.Forms.RadioButton
$radChangeOffice.Text = "Change / Add Office Edition"
$radChangeOffice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radChangeOffice.Location = New-Object System.Drawing.Point(15, 48)
$radChangeOffice.Size = New-Object System.Drawing.Size(460, 21)

$radChangeWin = New-Object System.Windows.Forms.RadioButton
$radChangeWin.Text = "Change Windows Edition"
$radChangeWin.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radChangeWin.Location = New-Object System.Drawing.Point(15, 74)
$radChangeWin.Size = New-Object System.Drawing.Size(460, 21)

$radActOffice = New-Object System.Windows.Forms.RadioButton
$radActOffice.Text = "Activate Office"
$radActOffice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radActOffice.Location = New-Object System.Drawing.Point(15, 100)
$radActOffice.Size = New-Object System.Drawing.Size(460, 21)

$radActWin = New-Object System.Windows.Forms.RadioButton
$radActWin.Text = "Activate Windows"
$radActWin.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radActWin.Location = New-Object System.Drawing.Point(15, 126)
$radActWin.Size = New-Object System.Drawing.Size(460, 21)

$radCheckStatus = New-Object System.Windows.Forms.RadioButton
$radCheckStatus.Text = "Check System Activation Status"
$radCheckStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radCheckStatus.Location = New-Object System.Drawing.Point(15, 152)
$radCheckStatus.Size = New-Object System.Drawing.Size(460, 21)

$grpChoice.Controls.AddRange(@($radNo, $radChangeOffice, $radChangeWin, $radActOffice, $radActWin, $radCheckStatus))
$form.Controls.Add($grpChoice)

# Action Button
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "RUN SELECTED PROCESS"
$btnStart.Location = New-Object System.Drawing.Point(20, 255)
$btnStart.Size = New-Object System.Drawing.Size(500, 42)
$btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = "Flat"
$form.Controls.Add($btnStart)

# Log Output Window
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point(20, 310)
$txtLog.Size = New-Object System.Drawing.Size(500, 465)
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.Text = "Status: Ready. Select an option and click 'RUN SELECTED PROCESS'."
$form.Controls.Add($txtLog)

function LogMessage($msg) {
    $txtLog.AppendText("`r`n" + $msg)
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    $form.Refresh()
}

# Helper to check specifically for ACTIVE installer processes (excluding background service)
function Test-ActiveOfficeInstaller {
    $installerNames = @("OfficeC2RClient", "setup", "integratedoffice")
    $procs = Get-Process -Name $installerNames -ErrorAction SilentlyContinue
    return ($procs -ne $null -and $procs.Count -gt 0)
}

# Helper to cleanup run_mas.ps1 and conditionally remove C:\ODT if no other files exist
function Cleanup-MasScript($ps1Path, $odtDir) {
    if (Test-Path $odtDir) {
        $allFiles = Get-ChildItem -Path $odtDir -ErrorAction SilentlyContinue
        $otherFiles = $allFiles | Where-Object { $_.Name -ne "run_mas.ps1" }
        
        if (Test-Path $ps1Path) {
            Remove-Item -Path $ps1Path -Force -ErrorAction SilentlyContinue
        }
        
        if (-not $otherFiles -or $otherFiles.Count -eq 0) {
            Remove-Item -Path $odtDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Custom Windows Restart Prompt Dialog
function Show-CustomRestartDialog {
    $restartDlg = New-Object System.Windows.Forms.Form
    $restartDlg.Text = "Microsoft Windows"
    $restartDlg.Size = New-Object System.Drawing.Size(430, 210)
    $restartDlg.StartPosition = "CenterScreen"
    $restartDlg.FormBorderStyle = "FixedDialog"
    $restartDlg.MaximizeBox = $false
    $restartDlg.MinimizeBox = $false
    $restartDlg.ShowInTaskbar = $false
    $restartDlg.StartPosition = "CenterParent"

    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = "You must restart your computer before the new settings will take effect."
    $lbl1.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $lbl1.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $lbl1.Location = New-Object System.Drawing.Point(20, 20)
    $lbl1.Size = New-Object System.Drawing.Size(375, 45)
    $restartDlg.Controls.Add($lbl1)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = "Do you want to restart your computer now?"
    $lbl2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $lbl2.Location = New-Object System.Drawing.Point(20, 75)
    $lbl2.Size = New-Object System.Drawing.Size(375, 30)
    $restartDlg.Controls.Add($lbl2)

    $btnNow = New-Object System.Windows.Forms.Button
    $btnNow.Text = "Restart Now"
    $btnNow.Location = New-Object System.Drawing.Point(195, 120)
    $btnNow.Size = New-Object System.Drawing.Size(95, 30)
    $btnNow.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $restartDlg.Controls.Add($btnNow)

    $btnLater = New-Object System.Windows.Forms.Button
    $btnLater.Text = "Restart Later"
    $btnLater.Location = New-Object System.Drawing.Point(300, 120)
    $btnLater.Size = New-Object System.Drawing.Size(95, 30)
    $btnLater.DialogResult = [System.Windows.Forms.DialogResult]::No
    $restartDlg.Controls.Add($btnLater)

    $restartDlg.AcceptButton = $btnNow
    $restartDlg.CancelButton = $btnLater

    $result = $restartDlg.ShowDialog()
    $restartDlg.Dispose()
    return $result
}

function Launch-InteractiveMas($stepTitle, $instructionText, $mode = "Office") {
    LogMessage ""
    LogMessage "=================================================="
    LogMessage ">>> $stepTitle"
    LogMessage ">>> INSTRUCTIONS FOR OPENING WINDOW:"
    LogMessage "    $instructionText"
    LogMessage "=================================================="
    
    # 5-second countdown
    for ($i = 5; $i -ge 1; $i--) {
        LogMessage "$i..."
        Start-Sleep -Seconds 1
    }
    
    $odtDir = "C:\ODT"
    if (-not (Test-Path $odtDir)) { 
        New-Item -Path $odtDir -ItemType Directory -Force | Out-Null 
    }
    
    # Exclude C:\ODT path from Defender scanning
    try { Add-MpPreference -ExclusionPath $odtDir -ErrorAction SilentlyContinue } catch {}

    # Create temporary ps1 file using single-quoted heredoc
    $ps1Path = Join-Path $odtDir "run_mas.ps1"
    
    $ps1Content = @'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$mas = Invoke-RestMethod -Uri "https://get.activated.win"
Invoke-Expression $mas
'@
    Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8
    
    LogMessage "Opening interactive menu window..."
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$ps1Path`""
        $psi.UseShellExecute = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
    } catch {
        LogMessage "Error launching interactive process: $_"
        Cleanup-MasScript $ps1Path $odtDir
        return $false
    }
    
    if ($mode -eq "Office") {
        LogMessage "Waiting for active Office installation to begin (monitoring up to 1 hour)..."
        
        $processFound = $false
        $maxWaitStart = 3600 # 1 hour limit to trigger
        $elapsed = 0
        $consecutiveHits = 0
        
        while (-not $processFound -and ($elapsed -lt $maxWaitStart)) {
            if (Test-ActiveOfficeInstaller) {
                $consecutiveHits++
                if ($consecutiveHits -ge 2) {
                    $processFound = $true
                    LogMessage "  Active Office installer process detected!"
                }
            } else {
                $consecutiveHits = 0
            }

            if (-not $processFound) {
                Start-Sleep -Seconds 3
                [System.Windows.Forms.Application]::DoEvents()
                $elapsed += 3
            }
        }
        
        Cleanup-MasScript $ps1Path $odtDir
        
        if (-not $processFound) {
            LogMessage "  [ERROR] No active Office installation started within 1 hour."
            [System.Windows.Forms.MessageBox]::Show(
                "No active Office installation was detected within 1 hour. Please check your internet connection or try again.",
                "Installation Timeout",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }

        LogMessage "  Waiting for Office installation/configuration to complete..."
        $timer = 0
        while ((Test-ActiveOfficeInstaller) -and ($timer -lt 3600)) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Seconds 3
            $timer += 3
        }
        LogMessage "  Office installation/configuration completed!"
    } 
    elseif ($mode -eq "InteractiveWait") {
        LogMessage "Waiting for interactive menu window to close..."
        while (-not $proc.HasExited) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Seconds 1
        }
        Cleanup-MasScript $ps1Path $odtDir
        LogMessage "Interactive menu closed."
    }
    else {
        Start-Sleep -Seconds 5
        Cleanup-MasScript $ps1Path $odtDir
    }
    
    LogMessage "Completed step: $stepTitle"
    return $true
}

# ------------------------------------------------------------
# Smart Edition Picker Dialog
# ------------------------------------------------------------
function Show-EditionDialog($changeMode) {
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $normFont = New-Object System.Drawing.Font("Segoe UI", 9)

    $dlg = New-Object System.Windows.Forms.Form
    if ($changeMode) { $dlg.Text = "Change / Add Office Edition" } else { $dlg.Text = "Install Office - Select Edition" }
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = "CenterParent"

    # Edition
    $grpEdition = New-Object System.Windows.Forms.GroupBox
    $grpEdition.Text = "Office Edition"
    $grpEdition.Font = $boldFont
    $grpEdition.Location = New-Object System.Drawing.Point(15, 10)
    $grpEdition.Size = New-Object System.Drawing.Size(500, 60)

    $cmbEdition = New-Object System.Windows.Forms.ComboBox
    $cmbEdition.DropDownStyle = "DropDownList"
    $cmbEdition.Font = $normFont
    $cmbEdition.Location = New-Object System.Drawing.Point(15, 24)
    $cmbEdition.Size = New-Object System.Drawing.Size(470, 26)
    $cmbEdition.MaxDropDownItems = 20
    foreach ($name in $editions.Keys) { [void]$cmbEdition.Items.Add($name) }
    $cmbEdition.SelectedIndex = 0
    $grpEdition.Controls.Add($cmbEdition)
    $dlg.Controls.Add($grpEdition)

    # Architecture + language
    $grpArch = New-Object System.Windows.Forms.GroupBox
    $grpArch.Text = "Architecture / Language"
    $grpArch.Font = $boldFont
    $grpArch.Location = New-Object System.Drawing.Point(15, 75)
    $grpArch.Size = New-Object System.Drawing.Size(500, 62)

    $rad64 = New-Object System.Windows.Forms.RadioButton
    $rad64.Text = "64-bit"
    $rad64.Font = $normFont
    $rad64.Location = New-Object System.Drawing.Point(15, 26)
    $rad64.Size = New-Object System.Drawing.Size(70, 22)
    $rad64.Checked = $true

    $rad32 = New-Object System.Windows.Forms.RadioButton
    $rad32.Text = "32-bit"
    $rad32.Font = $normFont
    $rad32.Location = New-Object System.Drawing.Point(90, 26)
    $rad32.Size = New-Object System.Drawing.Size(70, 22)

    $lblLang = New-Object System.Windows.Forms.Label
    $lblLang.Text = "Language:"
    $lblLang.Font = $normFont
    $lblLang.Location = New-Object System.Drawing.Point(180, 29)
    $lblLang.Size = New-Object System.Drawing.Size(65, 20)

    $cmbLang = New-Object System.Windows.Forms.ComboBox
    $cmbLang.DropDownStyle = "DropDownList"
    $cmbLang.Font = $normFont
    $cmbLang.Location = New-Object System.Drawing.Point(250, 25)
    $cmbLang.Size = New-Object System.Drawing.Size(235, 26)
    foreach ($name in $languages.Keys) { [void]$cmbLang.Items.Add($name) }
    $cmbLang.SelectedIndex = 0

    $grpArch.Controls.AddRange(@($rad64, $rad32, $lblLang, $cmbLang))
    $dlg.Controls.Add($grpArch)

    # Product key
    $grpKey = New-Object System.Windows.Forms.GroupBox
    $grpKey.Text = "Product key (optional - skips OHook activation if provided)"
    $grpKey.Font = $boldFont
    $grpKey.Location = New-Object System.Drawing.Point(15, 142)
    $grpKey.Size = New-Object System.Drawing.Size(500, 58)

    $txtKey = New-Object System.Windows.Forms.TextBox
    $txtKey.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtKey.Location = New-Object System.Drawing.Point(15, 23)
    $txtKey.Size = New-Object System.Drawing.Size(470, 24)
    $grpKey.Controls.Add($txtKey)
    $dlg.Controls.Add($grpKey)

    $y = 210

    # Buttons
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "NEXT"
    $btnOk.Location = New-Object System.Drawing.Point(255, $y)
    $btnOk.Size = New-Object System.Drawing.Size(125, 36)
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.FlatStyle = "Flat"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(390, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(125, 36)
    $btnCancel.Font = $normFont
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $btnOk.Add_Click({
        $k = $txtKey.Text.Trim().ToUpper()
        if ($k -and ($k -notmatch '^[A-Z0-9]{5}(-[A-Z0-9]{5}){4}$')) {
            [void][System.Windows.Forms.MessageBox]::Show("Product key format should be XXXXX-XXXXX-XXXXX-XXXXX-XXXXX.", "Invalid product key")
            return
        }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }.GetNewClosure())

    $dlg.Controls.AddRange(@($btnOk, $btnCancel))
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel
    $dlg.ClientSize = New-Object System.Drawing.Size(530, ($y + 50))

    $result = $dlg.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        $dlg.Dispose()
        return $null
    }

    $out = @{
        EditionName = $cmbEdition.SelectedItem
        EditionId   = $editions[$cmbEdition.SelectedItem]
        Arch        = $(if ($rad64.Checked) { "64" } else { "32" })
        LangId      = $languages[$cmbLang.SelectedItem]
        Key         = $txtKey.Text.Trim().ToUpper()
    }
    $dlg.Dispose()
    return $out
}

# ------------------------------------------------------------
# Edition-Dependent Dynamic App Config Dialog
# ------------------------------------------------------------
function Show-ChannelAndAppDialog($editionId, $editionDisplayName) {
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $normFont = New-Object System.Drawing.Font("Segoe UI", 9)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Configure Channel & Apps - $editionDisplayName"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = "CenterParent"

    # Update Channel Group
    $grpChan = New-Object System.Windows.Forms.GroupBox
    $grpChan.Text = "Update Channel"
    $grpChan.Font = $boldFont
    $grpChan.Location = New-Object System.Drawing.Point(15, 15)
    $grpChan.Size = New-Object System.Drawing.Size(500, 60)

    $cmbChan = New-Object System.Windows.Forms.ComboBox
    $cmbChan.DropDownStyle = "DropDownList"
    $cmbChan.Font = $normFont
    $cmbChan.Location = New-Object System.Drawing.Point(15, 24)
    $cmbChan.Size = New-Object System.Drawing.Size(470, 26)
    foreach ($name in $channels.Keys) { [void]$cmbChan.Items.Add($name) }
    $cmbChan.SelectedIndex = 0
    $grpChan.Controls.Add($cmbChan)
    $dlg.Controls.Add($grpChan)

    # Suite mapping based on official ODT reference app distribution
    $suiteAppsMap = @{
        "O365ProPlusRetail"           = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "Teams", "OneDrive")
        "O365ProPlusEEANoTeamsRetail" = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "OneDrive")
        "O365BusinessRetail"          = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "Teams", "OneDrive")
        "O365BusinessEEANoTeamsRetail"= @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "OneDrive")
        "O365HomePremRetail"          = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "OneDrive")

        "Professional2021Retail"      = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher")
        "Professional2024Retail"      = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access")

        "ProPlus2021Retail"           = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher")
        "ProPlus2021Volume"           = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher")
        "ProPlusSPLA2021Volume"       = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher")

        "ProPlus2024Retail"           = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access")
        "ProPlus2024Volume"           = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access")

        "Standard2021Volume"          = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher")
        "StandardSPLA2021Volume"      = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher")
        "Standard2024Volume"          = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote")

        "HomeBusiness2021Retail"      = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote")
        "HomeBusiness2024Retail"      = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote")

        "HomeStudent2021Retail"       = @("Word", "Excel", "PowerPoint", "OneNote")
        "Home2024Retail"              = @("Word", "Excel", "PowerPoint", "OneNote")

        "Personal2021Retail"          = @("Word", "Excel", "Outlook")
    }

    $appDisplayNames = @{
        "Word"       = "Word"
        "Excel"      = "Excel"
        "PowerPoint" = "PowerPoint"
        "Access"     = "Access"
        "Publisher"  = "Publisher"
        "Outlook"    = "Outlook"
        "OneNote"    = "OneNote"
        "Teams"      = "Teams"
        "OneDrive"   = "OneDrive / Groove"
    }

    $allowedAppIds = $suiteAppsMap[$editionId]
    if (-not $allowedAppIds) {
        if ($editionId -match '^(Word|Excel|PowerPoint|Access|Outlook|OneNote|Publisher|Visio|Project)(Runtime)?\d*Retail$') {
            $allowedAppIds = @($Matches[1])
        } else {
            $allowedAppIds = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher")
        }
    }

    $appChecks = @()
    $y = 85

    if ($allowedAppIds -and $allowedAppIds.Count -gt 0) {
        $grpApps = New-Object System.Windows.Forms.GroupBox
        $grpApps.Text = "Include / Exclude Apps (Varies by Edition)"
        $grpApps.Font = $boldFont
        $grpApps.Location = New-Object System.Drawing.Point(15, 85)
        $grpApps.Size = New-Object System.Drawing.Size(500, 185)

        $btnCheckAll = New-Object System.Windows.Forms.Button
        $btnCheckAll.Text = "Check All"
        $btnCheckAll.Font = $normFont
        $btnCheckAll.Location = New-Object System.Drawing.Point(15, 24)
        $btnCheckAll.Size = New-Object System.Drawing.Size(85, 24)

        $btnUncheckAll = New-Object System.Windows.Forms.Button
        $btnUncheckAll.Text = "Uncheck All"
        $btnUncheckAll.Font = $normFont
        $btnUncheckAll.Location = New-Object System.Drawing.Point(105, 24)
        $btnUncheckAll.Size = New-Object System.Drawing.Size(85, 24)

        $grpApps.Controls.AddRange(@($btnCheckAll, $btnUncheckAll))

        $i = 0
        foreach ($appId in $allowedAppIds) {
            $displayName = $appDisplayNames[$appId]
            if (-not $displayName) { $displayName = $appId }

            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = $displayName
            $cb.Tag = $appId
            $cb.Font = $normFont
            $cb.Checked = $true
            
            $col = $i % 3
            $row = [math]::Floor($i / 3)
            $cb.Location = New-Object System.Drawing.Point((15 + ($col * 160)), (58 + ($row * 24)))
            $cb.Size = New-Object System.Drawing.Size(155, 22)
            $grpApps.Controls.Add($cb)
            $appChecks += $cb
            $i++
        }
        $dlg.Controls.Add($grpApps)

        $btnCheckAll.Add_Click({
            foreach ($cb in $appChecks) { $cb.Checked = $true }
        })

        $btnUncheckAll.Add_Click({
            foreach ($cb in $appChecks) { $cb.Checked = $false }
        })

        $y = 280
    } else {
        $y = 90
    }

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "START"
    $btnOk.Location = New-Object System.Drawing.Point(255, $y)
    $btnOk.Size = New-Object System.Drawing.Size(125, 36)
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.FlatStyle = "Flat"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(390, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(125, 36)
    $btnCancel.Font = $normFont
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dlg.Controls.AddRange(@($btnOk, $btnCancel))
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel
    $dlg.ClientSize = New-Object System.Drawing.Size(530, ($y + 55))

    $result = $dlg.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        $dlg.Dispose()
        return $null
    }

    $excludeList = @($appChecks | Where-Object { -not $_.Checked } | ForEach-Object { $_.Tag })
    $selectedChannel = $channels[$cmbChan.SelectedItem]
    $dlg.Dispose()

    return @{
        Channel = $selectedChannel
        Exclude = $excludeList
    }
}

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
function Test-MicrosoftSigned($path) {
    $sig = Get-AuthenticodeSignature -FilePath $path
    return ($sig.Status -eq "Valid" -and $sig.SignerCertificate.Subject -match "Microsoft Corporation")
}

function Get-OdtSetup($odtDir) {
    $setupExe = Join-Path $odtDir "setup.exe"

    if ((Test-Path $setupExe) -and (Test-MicrosoftSigned $setupExe)) {
        return $setupExe
    }
    if (Test-Path $setupExe) { Remove-Item $setupExe -Force -ErrorAction SilentlyContinue }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'

    try {
        Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $setupExe -UseBasicParsing -ErrorAction Stop
    } catch {
        Start-Process curl.exe -ArgumentList "-sSL -o `"$setupExe`" `"https://officecdn.microsoft.com/pr/wsus/setup.exe`"" -Wait -NoNewWindow
    }

    if ((Test-Path $setupExe) -and (Test-MicrosoftSigned $setupExe)) { 
        return $setupExe 
    }
    return $null
}

function Build-ConfigXml($editionId, $channel, $arch, $langId, $excludeIds, $key) {
    $x = New-Object System.Collections.Generic.List[string]
    $x.Add('<Configuration>')
    $x.Add("  <Add OfficeClientEdition=`"$arch`" Channel=`"$channel`">")

    $prodLine = "    <Product ID=`"$editionId`""
    if ($key) { $prodLine += " PIDKEY=`"$key`"" }
    $prodLine += ">"
    $x.Add($prodLine)

    $x.Add("      <Language ID=`"$langId`" />")
    foreach ($app in $excludeIds) { $x.Add("      <ExcludeApp ID=`"$app`" />") }

    $x.Add('    </Product>')
    $x.Add('  </Add>')
    $x.Add('  <Updates Enabled="TRUE" />')
    $x.Add('  <RemoveMSI />')
    $x.Add('  <Display Level="Full" AcceptEULA="TRUE" />')
    $x.Add('</Configuration>')
    return ($x -join "`r`n")
}

function Wait-ForOfficeInstaller($statusLabel) {
    LogMessage "  Waiting for Office installer to start ($statusLabel)..."
    
    $processFound = $false
    $maxWaitStart = 3600 # Wait up to 1 hour for installer to appear
    $elapsed = 0
    $consecutiveHits = 0
    
    while (-not $processFound -and ($elapsed -lt $maxWaitStart)) {
        if (Test-ActiveOfficeInstaller) {
            $consecutiveHits++
            if ($consecutiveHits -ge 2) {
                $processFound = $true
                LogMessage "  Active Office installer process detected!"
            }
        } else {
            $consecutiveHits = 0
        }

        if (-not $processFound) {
            Start-Sleep -Seconds 3
            [System.Windows.Forms.Application]::DoEvents()
            $elapsed += 3
        }
    }
    
    if (-not $processFound) {
        LogMessage "  [ERROR] Office installer process did not start within 1 hour."
        [System.Windows.Forms.MessageBox]::Show(
            "The Office installer process did not start within 1 hour. Please check your internet connection or antivirus settings.",
            "Installation Timeout",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }

    LogMessage "  Waiting for Office installation/configuration to complete..."
    $timer = 0
    while ((Test-ActiveOfficeInstaller) -and ($timer -lt 3600)) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 3
        $timer += 3
    }
    LogMessage "  Office installation/configuration completed!"
    return $true
}

# ------------------------------------------------------------
# Main Event Handler
# ------------------------------------------------------------
$btnStart.Add_Click({
    $btnStart.Enabled = $false
    $odtDir = "C:\ODT"
    $configPath = Join-Path $odtDir "configuration.xml"

    if ($radCheckStatus.Checked) {
        LogMessage ""
        LogMessage "=================================================="
        LogMessage ">>> CHECKING SYSTEM & ACTIVATION STATUS..."
        LogMessage "=================================================="
        [System.Windows.Forms.Application]::DoEvents()

        # 1. OS Version & Build Info
        try {
            $os = Get-CimInstance Win32_OperatingSystem -Property Caption, Version, BuildNumber -ErrorAction SilentlyContinue
            LogMessage " [OS Details]"
            LogMessage "   Name    : $($os.Caption)"
            LogMessage "   Version : $($os.Version)"
            LogMessage "   Build   : $($os.BuildNumber)"
        } catch {
            LogMessage " [OS Details] Could not query OS information."
        }
        [System.Windows.Forms.Application]::DoEvents()

        # 2. Fast Windows Activation Check
        LogMessage " [Windows Activation]"
        try {
            $winLic = Get-CimInstance SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus = 1" -Property LicenseStatus -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($winLic) {
                LogMessage "   Status  : Activated (Genuine / License Active)"
            } else {
                LogMessage "   Status  : Not Activated / Unlicensed"
            }
        } catch {
            LogMessage "   Status  : Status check restricted or unavailable."
        }
        [System.Windows.Forms.Application]::DoEvents()

        # 3. Blazing Fast Office Status (Optimized Short-Circuiting)
        LogMessage " [Office Status]"
        $detectedOffice = $false
        $installType = ""
        $editionStr = ""
        $versionStr = ""
        $channelStr = ""

        # Priority A: Click-to-Run (C2R) - Instant Registry Read
        $c2rPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $c2rPath) {
            $c2rConfig = Get-ItemProperty -Path $c2rPath -ErrorAction SilentlyContinue
            if ($c2rConfig -and ($c2rConfig.ClientVersionToReport -or $c2rConfig.ProductReleaseIDs)) {
                $detectedOffice = $true
                $installType = "Click-to-Run (C2R)"
                $editionStr = if ($c2rConfig.ProductReleaseIDs) { $c2rConfig.ProductReleaseIDs } else { "Microsoft Office Click-to-Run" }
                $versionStr = if ($c2rConfig.VersionToReport) { $c2rConfig.VersionToReport } else { $c2rConfig.ClientVersionToReport }
                
                $chan = $c2rConfig.UpdateChannel
                if (-not $chan) { $chan = $c2rConfig.CDNBaseUrl }
                switch -wildcard ($chan) {
                    "*492350f6-3a01-4f97-b9c0-c7c6ddf67d60*" { $chan = "Current Channel" }
                    "*6420b000-8011-4920-b100-80053e1a0a4f*" { $chan = "Monthly Enterprise Channel" }
                    "*7ffbc6bf-bc32-4f92-8982-f9dd17fd3114*" { $chan = "Semi-Annual Enterprise Channel" }
                    "*b8f9b010-1c16-413d-bf80-5d241c05e81d*" { $chan = "Current Preview" }
                    "*20b04d5b-a721-4d4d-8013-a9ea623f7476*" { $chan = "Semi-Annual Preview" }
                    "*ea4a2090-a96b-4776-a737-9892f0e0063f*" { $chan = "Beta Channel" }
                    "*55301822-5e0f-4f76-badf-8be3e464b54e*" { $chan = "LTSC 2021 / 2024 Perpetual" }
                }
                $channelStr = if ($chan) { $chan } else { "Standard / Not configured" }
            }
        }

        # Priority B: Traditional MSI (Legacy suites) - Targeted Subkey Check
        if (-not $detectedOffice) {
            $uninstallPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            foreach ($path in $uninstallPaths) {
                if (Test-Path $path) {
                    $subkeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                    foreach ($sk in $subkeys) {
                        $dName = $sk.GetValue("DisplayName")
                        if ($dName -match "Microsoft Office (2010|2013|2016)|Professional 201|Standard 201" -and $dName -notmatch "Language|Proofing|Access Runtime|Add-in") {
                            $detectedOffice = $true
                            $installType = "Traditional MSI"
                            $editionStr = $dName
                            $versionStr = $sk.GetValue("DisplayVersion")
                            $channelStr = "N/A (Traditional MSI)"
                            break
                        }
                    }
                }
                if ($detectedOffice) { break }
            }
        }

        # Priority C: Microsoft Store / UWP AppX (Fast single package lookup)
        if (-not $detectedOffice) {
            $uwpOffice = Get-AppxPackage Microsoft.Office.Desktop -ErrorAction SilentlyContinue
            if ($uwpOffice) {
                $detectedOffice = $true
                $installType = "Microsoft Store App / UWP"
                $editionStr = "Microsoft Store Office Hub"
                $versionStr = $uwpOffice.Version
                $channelStr = "Managed via Microsoft Store"
            }
        }

        # Priority D: App-V Virtualized Package
        if (-not $detectedOffice) {
            $appvPath = "HKLM:\SOFTWARE\Microsoft\AppV\Client\Packages"
            if (Test-Path $appvPath) {
                $appvPkgs = Get-ChildItem -Path $appvPath -ErrorAction SilentlyContinue | Get-ItemProperty -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "Office" }
                if ($appvPkgs) {
                    $detectedOffice = $true
                    $installType = "App-V Virtualized Package"
                    $editionStr = "Virtualized Office Suite"
                    $versionStr = "Unknown (Virtualized)"
                    $channelStr = "Managed via App-V Server"
                }
            }
        }

        if (-not $detectedOffice) {
            LogMessage "   Installed     : No Office installation found."
        } else {
            LogMessage "   Install Type  : $installType"
            LogMessage "   Edition       : $editionStr"
            LogMessage "   Version/Build : $versionStr"
            LogMessage "   Update Channel: $channelStr"

            [System.Windows.Forms.Application]::DoEvents()

            # Instant OHook File Check (Bypasses slow cscript execution if present)
            $ohookDetected = $false
            $ohookPaths = @(
                "$env:ProgramFiles\Microsoft Office\root\vfs\System\sppc.dll",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\vfs\System\sppc.dll",
                "$env:ProgramFiles\Microsoft Office\root\vfs\System\sppcs.dll",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\vfs\System\sppcs.dll",
                "$env:ProgramFiles\Microsoft Office\Office16\sppc.dll",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office16\sppc.dll"
            )
            foreach ($p in $ohookPaths) {
                if (Test-Path $p) {
                    $ohookDetected = $true
                    break
                }
            }

            if ($ohookDetected) {
                LogMessage "   Activation    : Activated via OHook (Permanent)"
            } else {
                # Fallback to OSPP status check only if OHook is absent
                $officePaths = @(
                    "$env:ProgramFiles\Microsoft Office\Office16\OSPP.vbs",
                    "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.vbs"
                )
                $osppPath = $officePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

                if ($osppPath) {
                    try {
                        $cscriptOut = cscript.exe //nologo "$osppPath" /dstatus 2>&1
                        $activatedCount = 0
                        foreach ($line in $cscriptOut) {
                            if ($line -match "LICENSE STATUS:" -and $line -match "LICENSED") {
                                $activatedCount++
                            }
                        }
                        if ($activatedCount -gt 0) {
                            LogMessage "   Activation    : Activated (Licensed products: $activatedCount)"
                        } else {
                            LogMessage "   Activation    : Not fully activated or evaluation/unlicensed"
                        }
                    } catch {
                        LogMessage "   Activation    : Could not query OSPP details."
                    }
                } else {
                    LogMessage "   Activation    : Managed by Store/ClipSVC or OSPP path not found."
                }
            }
        }

        LogMessage "=================================================="
        LogMessage "=== PROCESS COMPLETE ==="

    } elseif ($radActOffice.Checked) {
        LogMessage "[MODE]: Activate Office"
        LogMessage "Processing Office setup..."
        try {
            $masScript = irm https://get.activated.win
            & ([scriptblock]::Create($masScript)) /OHook
            LogMessage "  Office activation completed."
        } catch {
            LogMessage "  Office activation error: $_"
        }
        LogMessage "=== PROCESS COMPLETE ==="

    } elseif ($radActWin.Checked) {
        LogMessage "[MODE]: Activate Windows"
        LogMessage "Processing Windows setup..."
        try {
            $masScript = irm https://get.activated.win
            & ([scriptblock]::Create($masScript)) /HWID
            LogMessage "  Windows activation completed."
        } catch {
            LogMessage "  Windows activation error: $_"
        }
        LogMessage "=== PROCESS COMPLETE ==="

    } elseif ($radChangeWin.Checked) {
        LogMessage "[MODE]: Change Windows Edition"
        Launch-InteractiveMas -stepTitle "Step 1: Change Windows Edition" -instructionText "Press '6' for Change Windows Edition, then pick your desired Windows Edition." -mode "InteractiveWait"
        
        LogMessage "[2/2] Processing Windows activation (/HWID)..."
        try {
            $masScript = irm https://get.activated.win
            & ([scriptblock]::Create($masScript)) /HWID
            LogMessage "  Windows activation completed."
        } catch {
            LogMessage "  Windows activation error: $_"
        }
        LogMessage "=== PROCESS COMPLETE ==="

        # Prompt for system restart using custom Microsoft Windows style dialog
        $restartResult = Show-CustomRestartDialog

        if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            LogMessage "Initiating immediate system restart..."
            Restart-Computer -Force
        } else {
            LogMessage "Restart postponed by user."
        }

    } elseif ($radChangeOffice.Checked) {
        LogMessage "[MODE]: Change / Add Office Edition"
        $success = Launch-InteractiveMas -stepTitle "Step 1: Change / Add Office Edition" -instructionText "Press '7' for Change / Add Office Edition, then select your edition." -mode "Office"
        
        if (-not $success) {
            LogMessage "=== PROCESS ABORTED / TIMED OUT ==="
            $btnStart.Enabled = $true
            return
        }

        LogMessage "[2/2] Processing Office setup..."
        try {
            $masScript = irm https://get.activated.win
            & ([scriptblock]::Create($masScript)) /OHook
            LogMessage "  Office step completed."
        } catch {
            LogMessage "  Office step error: $_"
        }
        LogMessage "=== PROCESS COMPLETE ==="

    } else {
        LogMessage "[MODE]: Install Office"

        $opts = Show-EditionDialog $false
        if (-not $opts) {
            LogMessage "  Cancelled."
            $btnStart.Enabled = $true
            return
        }

        $editionId = $opts.EditionId
        $arch      = $opts.Arch
        $langId    = $opts.LangId
        $key       = $opts.Key

        $configOpts = Show-ChannelAndAppDialog $editionId $opts.EditionName
        if (-not $configOpts) {
            LogMessage "  Cancelled."
            $btnStart.Enabled = $true
            return
        }

        $channel = $configOpts.Channel
        $exclude = $configOpts.Exclude
        
        if (-not (Test-Path $odtDir)) { New-Item -Path $odtDir -ItemType Directory -Force | Out-Null }
        try { Add-MpPreference -ExclusionPath $odtDir -ErrorAction SilentlyContinue } catch {}

        $setupExe = Join-Path $odtDir "setup.exe"
        
        LogMessage "[1/3] Getting Office Deployment Tool in C:\ODT..."
        $setupExe = Get-OdtSetup $odtDir

        if ($setupExe -and (Test-Path $setupExe)) {
            LogMessage "  setup.exe successfully verified in C:\ODT"
            LogMessage "[2/3] Writing dynamic XML setup configuration file..."
            
            $userXml = Build-ConfigXml $editionId $channel $arch $langId $exclude $key
            Set-Content -Path $configPath -Value $userXml -Encoding UTF8
            LogMessage "  Configuration saved to $configPath"

            LogMessage "[3/3] Starting Office Installation Window..."
            LogMessage "  Please wait for Microsoft's Office installation to complete..."
            
            Start-Process -FilePath $setupExe -ArgumentList "/configure `"$configPath`"" -WorkingDirectory $odtDir
            
            $success = Wait-ForOfficeInstaller -statusLabel "Office Installation"
            if (-not $success) {
                LogMessage "=== PROCESS ABORTED / TIMED OUT ==="
                $btnStart.Enabled = $true
                return
            }

            if ($key) {
                LogMessage "Product key was provided. Skipping automatic OHook activation."
            } else {
                LogMessage "Activating Office with /OHook..."
                try {
                    $masScript = irm https://get.activated.win
                    & ([scriptblock]::Create($masScript)) /OHook
                    LogMessage "  Office activation completed."
                } catch {
                    LogMessage "  Office activation error: $_"
                }
            }

            $askClean = [System.Windows.Forms.MessageBox]::Show(
                "Office installation and activation completed successfully.`n`nDo you want to delete the temporary installation folder (C:\ODT) to free up space?",
                "Cleanup ODT Files?",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($askClean -eq [System.Windows.Forms.DialogResult]::Yes) {
                LogMessage "Cleaning up temporary ODT files in C:\ODT..."
                try {
                    Remove-Item -Path $odtDir -Recurse -Force -ErrorAction Stop
                    LogMessage "  C:\ODT successfully deleted."
                } catch {
                    LogMessage "  Could not delete C:\ODT completely: $_"
                }
            } else {
                LogMessage "Kept C:\ODT folder."
            }

        } else {
            LogMessage "ERROR: Failed to obtain setup.exe in C:\ODT. Check internet or antivirus settings."
        }

        LogMessage "=== PROCESS COMPLETE ==="
    }

    $btnStart.Enabled = $true
})

[void]$form.ShowDialog()