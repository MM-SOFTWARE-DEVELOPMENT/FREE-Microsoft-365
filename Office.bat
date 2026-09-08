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

# GUI Window Configuration
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Office"
$form.Size = New-Object System.Drawing.Size(560, 800)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Header Title
$title = New-Object System.Windows.Forms.Label
$title.Text = "Microsoft Office"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$title.Size = New-Object System.Drawing.Size(520, 25)
$title.Location = New-Object System.Drawing.Point(15, 10)
$form.Controls.Add($title)

# Task Selection Box
$grpChoice = New-Object System.Windows.Forms.GroupBox
$grpChoice.Text = "Select Task"
$grpChoice.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpChoice.Location = New-Object System.Drawing.Point(20, 45)
$grpChoice.Size = New-Object System.Drawing.Size(500, 170)

$radNo = New-Object System.Windows.Forms.RadioButton
$radNo.Text = "Install Office"
$radNo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radNo.Location = New-Object System.Drawing.Point(15, 22)
$radNo.Size = New-Object System.Drawing.Size(460, 22)
$radNo.Checked = $true

$radChangeOffice = New-Object System.Windows.Forms.RadioButton
$radChangeOffice.Text = "Change / Add Office Edition"
$radChangeOffice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radChangeOffice.Location = New-Object System.Drawing.Point(15, 50)
$radChangeOffice.Size = New-Object System.Drawing.Size(460, 22)

$radChangeWin = New-Object System.Windows.Forms.RadioButton
$radChangeWin.Text = "Change Windows Edition"
$radChangeWin.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radChangeWin.Location = New-Object System.Drawing.Point(15, 78)
$radChangeWin.Size = New-Object System.Drawing.Size(460, 22)

$radActOffice = New-Object System.Windows.Forms.RadioButton
$radActOffice.Text = "Activate Office"
$radActOffice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radActOffice.Location = New-Object System.Drawing.Point(15, 106)
$radActOffice.Size = New-Object System.Drawing.Size(460, 22)

$radActWin = New-Object System.Windows.Forms.RadioButton
$radActWin.Text = "Activate Windows"
$radActWin.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radActWin.Location = New-Object System.Drawing.Point(15, 134)
$radActWin.Size = New-Object System.Drawing.Size(460, 22)

$grpChoice.Controls.AddRange(@($radNo, $radChangeOffice, $radChangeWin, $radActOffice, $radActWin))
$form.Controls.Add($grpChoice)

# Action Button
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "RUN SELECTED PROCESS"
$btnStart.Location = New-Object System.Drawing.Point(20, 225)
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
$txtLog.Location = New-Object System.Drawing.Point(20, 280)
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

function Launch-InteractiveMas($stepTitle, $instructionText) {
    LogMessage ""
    LogMessage "=================================================="
    LogMessage ">>> $stepTitle"
    LogMessage ">>> INSTRUCTIONS FOR OPENING WINDOW:"
    LogMessage "    $instructionText"
    LogMessage "=================================================="
    LogMessage "Opening menu window... Please interact directly with that window."
    
    $proc = Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "`"irm https://get.activated.win | iex`"" -PassThru
    $proc | Wait-Process
    LogMessage "Completed step: $stepTitle"
}

function Wait-ForOfficeInstaller($statusLabel) {
    LogMessage "  Waiting for background Office installer to finish ($statusLabel)..."
    
    Start-Sleep -Seconds 10
    
    $timer = 0
    while ((Get-Process -Name "OfficeC2RClient", "setup", "integratedoffice" -ErrorAction SilentlyContinue) -and ($timer -lt 1800)) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 3
        $timer += 3
    }
    LogMessage "  Office installation/configuration completed!"
}

$btnStart.Add_Click({
    $btnStart.Enabled = $false
    $odtDir = "C:\ODT"
    $configPath = Join-Path $odtDir "configuration.xml"

    if ($radActOffice.Checked) {
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
        
        Launch-InteractiveMas -stepTitle "Step 1: Change Windows Edition" -instructionText "Press '6' for Change Windows Edition, then pick your desired Windows Edition."
        
        LogMessage "[2/2] Processing Windows setup..."
        try {
            $masScript = irm https://get.activated.win
            & ([scriptblock]::Create($masScript)) /HWID
            LogMessage "  Windows step completed."
        } catch {
            LogMessage "  Windows step error: $_"
        }

        LogMessage "=== PROCESS COMPLETE ==="

    } elseif ($radChangeOffice.Checked) {
        LogMessage "[MODE]: Change / Add Office Edition"
        
        Launch-InteractiveMas -stepTitle "Step 1: Change / Add Office Edition" -instructionText "Press '7' for Change / Add Office Edition, then select your edition."
        
        Wait-ForOfficeInstaller -statusLabel "Edition Change"

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
        
        if (-not (Test-Path $odtDir)) { New-Item -Path $odtDir -ItemType Directory -Force | Out-Null }
        try { Add-MpPreference -ExclusionPath $odtDir -ErrorAction SilentlyContinue } catch {}

        $setupExe = Join-Path $odtDir "setup.exe"
        $odtInstaller = Join-Path $odtDir "officedeploymenttool.exe"
        
        LogMessage "[1/4] Extracting setup files into C:\ODT..."

        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri "https://download.microsoft.com/download/2/7/A/27AF1BE6-DBA1-460B-8833-87A382A13D03/officedeploymenttool_17328-20162.exe" -OutFile $odtInstaller -UseBasicParsing -ErrorAction Stop
            
            if (Test-Path $odtInstaller) {
                $extProc = Start-Process -FilePath $odtInstaller -ArgumentList "/extract:`"$odtDir`" /quiet /norestart" -PassThru -Wait
                Remove-Item $odtInstaller -Force -ErrorAction SilentlyContinue
            }
        } catch {
            LogMessage "  Package extraction failed, trying direct binary download..."
        }

        if (-not (Test-Path $setupExe)) {
            try {
                Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $setupExe -UseBasicParsing -ErrorAction Stop
            } catch {
                Start-Process curl.exe -ArgumentList "-sSL -o `"$setupExe`" `"https://officecdn.microsoft.com/pr/wsus/setup.exe`"" -Wait -NoNewWindow
            }
        }

        if (Test-Path $setupExe) {
            LogMessage "  setup.exe successfully verified in C:\ODT"
            LogMessage "[2/4] Writing XML setup configuration file..."
            
            $userXml = @"
<Configuration ID="a0dd862e-8678-46f0-8b09-aa271fa10a1c">
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@
            Set-Content -Path $configPath -Value $userXml
            LogMessage "  Configuration saved to $configPath"

            LogMessage "[3/4] Starting Office Installation Window..."
            LogMessage "  Please wait for Microsoft's Office installation to complete..."
            
            Start-Process -FilePath $setupExe -ArgumentList "/configure `"$configPath`"" -WorkingDirectory $odtDir
            
            Wait-ForOfficeInstaller -statusLabel "Office Installation"

            LogMessage "[4/4] Activating Office..."
            try {
                $masScript = irm https://get.activated.win
                & ([scriptblock]::Create($masScript)) /OHook
                LogMessage "  Office activation completed."
            } catch {
                LogMessage "  Office activation error: $_"
            }

            $askEdition = [System.Windows.Forms.MessageBox]::Show(
                "Office installation and activation completed.`n`nDo you want to change or add an Office edition now?",
                "Change Office Edition?",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($askEdition -eq [System.Windows.Forms.DialogResult]::Yes) {
                Launch-InteractiveMas -stepTitle "Optional Step: Change / Add Office Edition" -instructionText "Press '7' for Change / Add Office Edition, then select your desired edition."
                
                Wait-ForOfficeInstaller -statusLabel "Edition Menu Updates"

                LogMessage "Re-activating Office after edition change..."
                try {
                    $masScript = irm https://get.activated.win
                    & ([scriptblock]::Create($masScript)) /OHook
                    LogMessage "  Office re-activation completed."
                } catch {
                    LogMessage "  Office re-activation error: $_"
                }
            } else {
                LogMessage "Skipped edition change step."
            }
        } else {
            LogMessage "ERROR: Failed to obtain setup.exe in C:\ODT. Check internet or antivirus settings."
        }

        LogMessage "=== PROCESS COMPLETE ==="
    }

    $btnStart.Enabled = $true
})

[void]$form.ShowDialog()