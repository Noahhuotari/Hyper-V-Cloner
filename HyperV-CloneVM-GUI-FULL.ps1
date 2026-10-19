<#    
    .NOTES
    ===========================================================================
     Created on:    10/29/2025
     Updated on:	10/19/2026
     Created by:    Noah Huotari
     Organization:  HBS
     Filename:      HyperV-CloneVM-GUI-FULL.ps1
    ===========================================================================
    .DESCRIPTION
        GUI to deploy VMs in Hyper-V using a template
        The script can also customize the hardware, set a static IP, change the hostname, and join the domain
#>

<#
    .ChangeLog
    ===========================================================================
    2025-10-29 - Created script
    2026-06-11 - Added Sysprep pre-flight check and dynamic offline VHDX unattend.xml injection
    2026-10-19 - Added Target Host and CSV selection, updated log directory to C:\hbs, added HBS branding
#>

 # --- Load .NET Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net

# --- Global Variables & HBS Branding Colors ---
$global:logFile = $null
$script:csvMap = @{}
$script:savedLocalAdminCred = $null
$script:savedDomainCred = $null
$script:savedDomainName = $null
$hbsBlue = [System.Drawing.Color]::FromArgb(0, 122, 255)       # HBS Blue (#007aff)
$hbsBlueDark = [System.Drawing.Color]::FromArgb(0, 98, 204)    # Hover / Accent shade
$hbsBgLight = [System.Drawing.Color]::FromArgb(244, 247, 251)  # Subtle cool background
$hbsTextDark = [System.Drawing.Color]::FromArgb(30, 35, 45)    # Primary text color

# --- Create the Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "HBS - Hyper-V VM Cloner"
$form.Size = New-Object System.Drawing.Size(460, 730)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.BackColor = $hbsBgLight
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.ForeColor = $hbsTextDark

# --- HBS Branded Header Banner ---
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
$pnlHeader.Size = New-Object System.Drawing.Size(460, 48)
$pnlHeader.BackColor = $hbsBlue
$form.Controls.Add($pnlHeader)

$lblHeaderTitle = New-Object System.Windows.Forms.Label
$lblHeaderTitle.Text = "HBS  |  Hyper-V VM Cloner"
$lblHeaderTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblHeaderTitle.ForeColor = [System.Drawing.Color]::White
$lblHeaderTitle.Location = New-Object System.Drawing.Point(15, 11)
$lblHeaderTitle.AutoSize = $true
$pnlHeader.Controls.Add($lblHeaderTitle)

# --- Function to Clean Up ---
$form.Add_FormClosing({
    # Clean up form resources
    $form.Dispose()
})

# --- === 1. CREATE MASTER TAB CONTROL === ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 58)
$tabControl.Size = New-Object System.Drawing.Size(420, 390)
$form.Controls.Add($tabControl)

# --- === 2. CREATE "CLONE VM" TAB === ---
$tabClone = New-Object System.Windows.Forms.TabPage
$tabClone.Text = "1. Clone VM"
$tabClone.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabClone)

# --- Add Controls: Host / Cluster ---
$lblHost = New-Object System.Windows.Forms.Label
$lblHost.Text = "Host/Cluster FQDN:"
$lblHost.Location = New-Object System.Drawing.Point(20, 23)
$lblHost.AutoSize = $true
$tabClone.Controls.Add($lblHost)

$txtHost = New-Object System.Windows.Forms.TextBox
$txtHost.Location = New-Object System.Drawing.Point(160, 20)
$txtHost.Size = New-Object System.Drawing.Size(150, 20)
$txtHost.Text = $env:ComputerName
$tabClone.Controls.Add($txtHost)

$chkCluster = New-Object System.Windows.Forms.CheckBox
$chkCluster.Text = "Cluster?"
$chkCluster.Location = New-Object System.Drawing.Point(320, 20)
$chkCluster.AutoSize = $true
$tabClone.Controls.Add($chkCluster)

# --- Add Controls: Get VMs Button ---
$btnGetVMs = New-Object System.Windows.Forms.Button
$btnGetVMs.Text = "Get VMs"
$btnGetVMs.Location = New-Object System.Drawing.Point(20, 60)
$btnGetVMs.Size = New-Object System.Drawing.Size(360, 30)
$btnGetVMs.BackColor = $hbsBlue
$btnGetVMs.ForeColor = [System.Drawing.Color]::White
$btnGetVMs.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGetVMs.FlatAppearance.BorderSize = 0
$btnGetVMs.FlatAppearance.MouseOverBackColor = $hbsBlueDark
$btnGetVMs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabClone.Controls.Add($btnGetVMs)

# --- Add Controls: VM Selection ---
$lblSourceVM = New-Object System.Windows.Forms.Label
$lblSourceVM.Text = "Select VM to Clone:"
$lblSourceVM.Location = New-Object System.Drawing.Point(20, 113)
$lblSourceVM.AutoSize = $true
$tabClone.Controls.Add($lblSourceVM)

$cboVMs = New-Object System.Windows.Forms.ComboBox
$cboVMs.Location = New-Object System.Drawing.Point(160, 110)
$cboVMs.Size = New-Object System.Drawing.Size(220, 20)
$cboVMs.DropDownStyle = "DropDownList"
$cboVMs.Enabled = $false
$tabClone.Controls.Add($cboVMs)

# --- Add Controls: New VM Name ---
$lblNewName = New-Object System.Windows.Forms.Label
$lblNewName.Text = "New VM Name:"
$lblNewName.Location = New-Object System.Drawing.Point(20, 153)
$lblNewName.AutoSize = $true
$tabClone.Controls.Add($lblNewName)

$txtNewName = New-Object System.Windows.Forms.TextBox
$txtNewName.Location = New-Object System.Drawing.Point(160, 150)
$txtNewName.Size = New-Object System.Drawing.Size(220, 20)
$tabClone.Controls.Add($txtNewName)

# --- Add Controls: Target Host Selection ---
$lblTargetHost = New-Object System.Windows.Forms.Label
$lblTargetHost.Text = "Target Host:"
$lblTargetHost.Location = New-Object System.Drawing.Point(20, 193)
$lblTargetHost.AutoSize = $true
$tabClone.Controls.Add($lblTargetHost)

$cboTargetHost = New-Object System.Windows.Forms.ComboBox
$cboTargetHost.Location = New-Object System.Drawing.Point(160, 190)
$cboTargetHost.Size = New-Object System.Drawing.Size(220, 20)
$cboTargetHost.DropDownStyle = "DropDown"
$cboTargetHost.Text = $env:ComputerName
$tabClone.Controls.Add($cboTargetHost)

# --- Add Controls: Target CSV Selection ---
$lblCSV = New-Object System.Windows.Forms.Label
$lblCSV.Text = "Select CSV:"
$lblCSV.Location = New-Object System.Drawing.Point(20, 233)
$lblCSV.AutoSize = $true
$tabClone.Controls.Add($lblCSV)

$cboCSV = New-Object System.Windows.Forms.ComboBox
$cboCSV.Location = New-Object System.Drawing.Point(160, 230)
$cboCSV.Size = New-Object System.Drawing.Size(220, 20)
$cboCSV.DropDownStyle = "DropDownList"
$cboCSV.Enabled = $false
$tabClone.Controls.Add($cboCSV)

# --- Add Controls: Virtual Switch Selection ---
$lblVMSwitch = New-Object System.Windows.Forms.Label
$lblVMSwitch.Text = "Virtual Switch:"
$lblVMSwitch.Location = New-Object System.Drawing.Point(20, 273)
$lblVMSwitch.AutoSize = $true
$tabClone.Controls.Add($lblVMSwitch)

$cboVMSwitch = New-Object System.Windows.Forms.ComboBox
$cboVMSwitch.Location = New-Object System.Drawing.Point(160, 270)
$cboVMSwitch.Size = New-Object System.Drawing.Size(220, 20)
$cboVMSwitch.DropDownStyle = "DropDownList"
$cboVMSwitch.Enabled = $false
$tabClone.Controls.Add($cboVMSwitch)

# --- Add Controls: VLAN ID ---
$lblVlan = New-Object System.Windows.Forms.Label
$lblVlan.Text = "VLAN ID (Optional):"
$lblVlan.Location = New-Object System.Drawing.Point(20, 313)
$lblVlan.AutoSize = $true
$tabClone.Controls.Add($lblVlan)

$txtVlan = New-Object System.Windows.Forms.TextBox
$txtVlan.Location = New-Object System.Drawing.Point(160, 310)
$txtVlan.Size = New-Object System.Drawing.Size(65, 20)
$tabClone.Controls.Add($txtVlan)

$lblVlanHint = New-Object System.Windows.Forms.Label
$lblVlanHint.Text = "(Blank = Template/Untagged)"
$lblVlanHint.Location = New-Object System.Drawing.Point(232, 313)
$lblVlanHint.AutoSize = $true
$lblVlanHint.ForeColor = [System.Drawing.Color]::Gray
$lblVlanHint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$tabClone.Controls.Add($lblVlanHint)

# --- === 3. CREATE "POST-CLONE" TAB === ---
$tabPostClone = New-Object System.Windows.Forms.TabPage
$tabPostClone.Text = "2. Post-Clone Config"
$tabPostClone.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabPostClone)

# --- Hardware Group ---
$grpHardware = New-Object System.Windows.Forms.GroupBox
$grpHardware.Text = "VM Hardware"
$grpHardware.Location = New-Object System.Drawing.Point(15, 15)
$grpHardware.Size = New-Object System.Drawing.Size(380, 55)
$grpHardware.ForeColor = $hbsBlue
$tabPostClone.Controls.Add($grpHardware)

$chkHardware = New-Object System.Windows.Forms.CheckBox
$chkHardware.Text = "Customize:"
$chkHardware.Location = New-Object System.Drawing.Point(15, 23)
$chkHardware.AutoSize = $true
$grpHardware.Controls.Add($chkHardware)

$lblCPU = New-Object System.Windows.Forms.Label
$lblCPU.Text = "vCPUs:"
$lblCPU.Location = New-Object System.Drawing.Point(110, 25)
$lblCPU.AutoSize = $true
$grpHardware.Controls.Add($lblCPU)

$txtCPU = New-Object System.Windows.Forms.TextBox
$txtCPU.Location = New-Object System.Drawing.Point(160, 22)
$txtCPU.Size = New-Object System.Drawing.Size(40, 20)
$txtCPU.Text = "4"
$grpHardware.Controls.Add($txtCPU)

# --- Modified Memory Controls ---
$lblMemory = New-Object System.Windows.Forms.Label
$lblMemory.Text = "Memory (GB):"
$lblMemory.Location = New-Object System.Drawing.Point(215, 25)
$lblMemory.AutoSize = $true
$grpHardware.Controls.Add($lblMemory)

$txtMemory = New-Object System.Windows.Forms.TextBox
$txtMemory.Location = New-Object System.Drawing.Point(300, 22)
$txtMemory.Size = New-Object System.Drawing.Size(40, 20)
$txtMemory.Text = "8"
$grpHardware.Controls.Add($txtMemory)

# --- Guest OS Group ---
$grpGuestOS = New-Object System.Windows.Forms.GroupBox
$grpGuestOS.Text = "Guest OS (Requires VM Start & Guest Services)"
$grpGuestOS.Location = New-Object System.Drawing.Point(15, 80)
$grpGuestOS.Size = New-Object System.Drawing.Size(380, 295) 
$grpGuestOS.ForeColor = $hbsBlue
$tabPostClone.Controls.Add($grpGuestOS)

$chkGuestConfig = New-Object System.Windows.Forms.CheckBox
$chkGuestConfig.Text = "Perform Guest OS Configuration"
$chkGuestConfig.Location = New-Object System.Drawing.Point(15, 25)
$chkGuestConfig.AutoSize = $true
$grpGuestOS.Controls.Add($chkGuestConfig)

$btnClearCreds = New-Object System.Windows.Forms.Button
$btnClearCreds.Text = "Clear Saved Creds"
$btnClearCreds.Location = New-Object System.Drawing.Point(245, 21)
$btnClearCreds.Size = New-Object System.Drawing.Size(120, 24)
$btnClearCreds.BackColor = [System.Drawing.Color]::White
$btnClearCreds.ForeColor = $hbsBlue
$btnClearCreds.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClearCreds.FlatAppearance.BorderColor = $hbsBlue
$btnClearCreds.FlatAppearance.BorderSize = 1
$btnClearCreds.Add_Click({
    $script:savedLocalAdminCred = $null
    $script:savedDomainCred = $null
    $script:savedDomainName = $null
    Add-Log -Message "Cleared saved session credentials (Template & Domain)." -Color Cyan
})
$grpGuestOS.Controls.Add($btnClearCreds)

$chkHostname = New-Object System.Windows.Forms.CheckBox
$chkHostname.Text = "Set Guest Hostname to VM Name"
$chkHostname.Location = New-Object System.Drawing.Point(30, 55)
$chkHostname.AutoSize = $true
$chkHostname.Checked = $true
$grpGuestOS.Controls.Add($chkHostname)

# --- IP Address Sub-Group ---
$grpIP = New-Object System.Windows.Forms.GroupBox
$grpIP.Text = "Network"
$grpIP.Location = New-Object System.Drawing.Point(15, 85)
$grpIP.Size = New-Object System.Drawing.Size(350, 125) 
$grpIP.ForeColor = $hbsBlue
$grpGuestOS.Controls.Add($grpIP)

$chkSetIP = New-Object System.Windows.Forms.CheckBox
$chkSetIP.Text = "Set Static IP Address"
$chkSetIP.Location = New-Object System.Drawing.Point(15, 20)
$chkSetIP.AutoSize = $true
$grpIP.Controls.Add($chkSetIP)

$lblIP = New-Object System.Windows.Forms.Label; $lblIP.Text = "IP:"; $lblIP.Location = New-Object System.Drawing.Point(30, 48); $lblIP.AutoSize = $true; $grpIP.Controls.Add($lblIP)
$txtIP = New-Object System.Windows.Forms.TextBox; $txtIP.Location = New-Object System.Drawing.Point(60, 45); $txtIP.Size = New-Object System.Drawing.Size(100, 20); $grpIP.Controls.Add($txtIP)
$lblSubnet = New-Object System.Windows.Forms.Label; $lblSubnet.Text = "Subnet:"; $lblSubnet.Location = New-Object System.Drawing.Point(170, 48); $lblSubnet.AutoSize = $true; $grpIP.Controls.Add($lblSubnet)
$txtSubnet = New-Object System.Windows.Forms.TextBox; $txtSubnet.Location = New-Object System.Drawing.Point(215, 45); $txtSubnet.Size = New-Object System.Drawing.Size(120, 20); $txtSubnet.Text = "255.255.255.0"; $grpIP.Controls.Add($txtSubnet)
$lblGateway = New-Object System.Windows.Forms.Label; $lblGateway.Text = "GW:"; $lblGateway.Location = New-Object System.Drawing.Point(30, 73); $lblGateway.AutoSize = $true; $grpIP.Controls.Add($lblGateway)
$txtGateway = New-Object System.Windows.Forms.TextBox; $txtGateway.Location = New-Object System.Drawing.Point(60, 70); $txtGateway.Size = New-Object System.Drawing.Point(100, 20); $grpIP.Controls.Add($txtGateway)
$lblDNS = New-Object System.Windows.Forms.Label; $lblDNS.Text = "DNS1:"; $lblDNS.Location = New-Object System.Drawing.Point(170, 73); $lblDNS.AutoSize = $true; $grpIP.Controls.Add($lblDNS)
$txtDNS = New-Object System.Windows.Forms.TextBox; $txtDNS.Location = New-Object System.Drawing.Point(215, 70); $txtDNS.Size = New-Object System.Drawing.Size(120, 20); $grpIP.Controls.Add($txtDNS)

$lblDNS2 = New-Object System.Windows.Forms.Label; $lblDNS2.Text = "DNS2:"; $lblDNS2.Location = New-Object System.Drawing.Point(170, 98); $lblDNS2.AutoSize = $true; $grpIP.Controls.Add($lblDNS2)
$txtDNS2 = New-Object System.Windows.Forms.TextBox; $txtDNS2.Location = New-Object System.Drawing.Point(215, 95); $txtDNS2.Size = New-Object System.Drawing.Size(120, 20); $grpIP.Controls.Add($txtDNS2)

# --- Domain Join Sub-Group ---
$grpDomain = New-Object System.Windows.Forms.GroupBox
$grpDomain.Text = "Active Directory"
$grpDomain.Location = New-Object System.Drawing.Point(15, 220) 
$grpDomain.Size = New-Object System.Drawing.Size(350, 60)
$grpDomain.ForeColor = $hbsBlue
$grpGuestOS.Controls.Add($grpDomain)

$chkDomainJoin = New-Object System.Windows.Forms.CheckBox
$chkDomainJoin.Text = "Join Domain:"
$chkDomainJoin.Location = New-Object System.Drawing.Point(15, 25)
$chkDomainJoin.AutoSize = $true
$grpDomain.Controls.Add($chkDomainJoin)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = New-Object System.Drawing.Point(110, 23)
$txtDomain.Size = New-Object System.Drawing.Size(225, 20)
$grpDomain.Controls.Add($txtDomain)

# Ensure controls inside GroupBoxes retain dark text rather than inheriting GroupBox header blue
foreach ($grp in @($grpHardware, $grpGuestOS, $grpIP, $grpDomain)) {
    foreach ($ctrl in $grp.Controls) {
        if ($ctrl -isnot [System.Windows.Forms.GroupBox]) {
            $ctrl.ForeColor = $hbsTextDark
        }
    }
}

# --- Dynamic GUI Logic ---
$grpHardware.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $false }
$grpGuestOS.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $false }
$grpIP.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $false }
$grpDomain.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $false }

$chkHardware.Add_CheckedChanged({ $grpHardware.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $chkHardware.Checked } })
$chkGuestConfig.Add_CheckedChanged({ $grpGuestOS.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $chkGuestConfig.Checked } })
$chkSetIP.Add_CheckedChanged({ $grpIP.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $chkSetIP.Checked } })
$chkDomainJoin.Add_CheckedChanged({ $grpDomain.Controls | Where-Object { $_ -isnot [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Enabled = $chkDomainJoin.Checked } })
$chkCluster.Add_CheckedChanged({ $cboCSV.Enabled = ($chkCluster.Checked -and $cboCSV.Items.Count -gt 0) })

# --- === 4. ADD CONTROLS *OUTSIDE* TABS === ---
$chkTestMode = New-Object System.Windows.Forms.CheckBox
$chkTestMode.Text = "Test Mode (Dry Run)"
$chkTestMode.Location = New-Object System.Drawing.Point(20, 458)
$chkTestMode.AutoSize = $true
$chkTestMode.Checked = $true
$form.Controls.Add($chkTestMode)

$btnClone = New-Object System.Windows.Forms.Button
$btnClone.Text = "Start Clone"
$btnClone.Location = New-Object System.Drawing.Point(240, 453)
$btnClone.Size = New-Object System.Drawing.Size(90, 28)
$btnClone.BackColor = $hbsBlue
$btnClone.ForeColor = [System.Drawing.Color]::White
$btnClone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClone.FlatAppearance.BorderSize = 0
$btnClone.FlatAppearance.MouseOverBackColor = $hbsBlueDark
$btnClone.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.AcceptButton = $btnClone
$form.Controls.Add($btnClone)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(340, 453)
$btnCancel.Size = New-Object System.Drawing.Size(90, 28)
$btnCancel.BackColor = [System.Drawing.Color]::White
$btnCancel.ForeColor = $hbsBlue
$btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCancel.FlatAppearance.BorderColor = $hbsBlue
$btnCancel.FlatAppearance.BorderSize = 1
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($btnCancel)

# --- Add Log Box ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.Location = New-Object System.Drawing.Point(20, 495)
$lblLog.AutoSize = $true
$lblLog.ForeColor = $hbsBlue
$lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 515)
$txtLog.Size = New-Object System.Drawing.Size(410, 155)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = "White"
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

# --- Helper Function to Write to Log ---
function Add-Log {
    param (
        [string]$Message,
        [string]$Color = "Black" 
    )

    $prefix = ""
    if ($Color -eq "Red") { $prefix = "[ERROR] " } 
    elseif ($Color -eq "Yellow") { $prefix = "[WARN] " } 
    elseif ($Color -eq "Cyan") { $prefix = "[INFO] " } 
    elseif ($Color -eq "Green") { $prefix = "[SUCCESS] " }

    $logEntry = "$(Get-Date -Format 'HH:mm:ss') $prefix$Message"
    $txtLog.AppendText("$logEntry`r`n")

    if ($global:logFile) {
        try {
            $logEntry | Out-File -FilePath $global:logFile -Append -Encoding utf8 -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error writing to log file: $($_.Exception.Message). File logging will be disabled.", "Log Error", "OK", "Warning")
            $global:logFile = $null
        }
    }
}

# --- Event Handlers ---

$btnGetVMs.Add_Click({
    $global:logFile = $null
    $txtLog.Text = ""
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $cboVMs.Items.Clear()
    $cboVMs.Text = ""
    $cboVMs.Enabled = $false
    $cboTargetHost.Items.Clear()
    $cboCSV.Items.Clear()
    $cboCSV.Enabled = $false
    $cboVMSwitch.Items.Clear()
    $cboVMSwitch.Enabled = $false
    $script:csvMap.Clear()
    $form.Update()

    $hostOrClusterName = $txtHost.Text.Trim()
    $isCluster = $chkCluster.Checked

    try {
        $vmHost = $hostOrClusterName
        $switchQueryHost = $vmHost
        if ($isCluster) {
            Add-Log -Message "Cluster selected. Getting all cluster VMs from $hostOrClusterName..." -Color Cyan
            $vms = Get-ClusterGroup -Cluster $hostOrClusterName | Where-Object { $_.GroupType -eq "VirtualMachine" } | Get-VM

            Add-Log -Message "Getting cluster hosts from $hostOrClusterName..." -Color Cyan
            $clusterNodes = @(Get-ClusterNode -Cluster $hostOrClusterName | Where-Object { $_.State -eq "Up" } | Select-Object -ExpandProperty Name | Sort-Object)
            if ($clusterNodes.Count -gt 0) {
                $cboTargetHost.Items.AddRange($clusterNodes)
                $cboTargetHost.SelectedIndex = 0
                $switchQueryHost = $clusterNodes[0]
                Add-Log -Message "Successfully loaded $($clusterNodes.Count) cluster host(s)." -Color Green
            } else {
                Add-Log -Message "No active cluster hosts found on $hostOrClusterName." -Color Yellow
            }

            Add-Log -Message "Getting CSVs from $hostOrClusterName..." -Color Cyan
            $csvs = Get-ClusterSharedVolume -Cluster $hostOrClusterName
            if ($csvs) {
                foreach ($csv in ($csvs | Sort-Object Name)) {
                    foreach ($volInfo in $csv.SharedVolumeInfo) {
                        $csvPath = $volInfo.FriendlyVolumeName
                        if ($csvPath) {
                            $freeGB = [math]::Round($volInfo.Partition.FreeSpace / 1GB, 1)
                            $displayText = "$($csv.Name) ($csvPath - ${freeGB}GB Free)"
                            $script:csvMap[$displayText] = $csvPath
                            [void]$cboCSV.Items.Add($displayText)
                        }
                    }
                }
                if ($cboCSV.Items.Count -gt 0) {
                    $cboCSV.Enabled = $true
                    $cboCSV.SelectedIndex = 0
                    Add-Log -Message "Successfully loaded $($cboCSV.Items.Count) CSV(s)." -Color Green
                }
            } else {
                Add-Log -Message "No CSVs found on $hostOrClusterName." -Color Yellow
            }
        } else {
            Add-Log -Message "Standalone host selected. Getting VMs from $vmHost..." -Color Cyan
            $vms = Get-VM -ComputerName $vmHost
            [void]$cboTargetHost.Items.Add($vmHost)
            $cboTargetHost.Text = $vmHost
        }

        Add-Log -Message "Getting Virtual Switches from $switchQueryHost..." -Color Cyan
        [void]$cboVMSwitch.Items.Add("<Keep Template Switch>")
        $switches = @(Get-VMSwitch -ComputerName $switchQueryHost -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object)
        if ($switches.Count -gt 0) {
            $cboVMSwitch.Items.AddRange($switches)
            Add-Log -Message "Successfully loaded $($switches.Count) Virtual Switch(es)." -Color Green
        }
        $cboVMSwitch.Enabled = $true
        $cboVMSwitch.SelectedIndex = 0

        if ($vms) {
            $vmNames = @($vms | Select-Object -ExpandProperty Name | Sort-Object)
            $cboVMs.Items.AddRange($vmNames)
            $cboVMs.Enabled = $true
            $cboVMs.SelectedIndex = 0
            Add-Log -Message "Successfully loaded $($vmNames.Count) VMs." -Color Green
        } else {
            [System.Windows.Forms.MessageBox]::Show("No VMs found on $hostOrClusterName.", "No VMs", "OK", "Information")
            Add-Log -Message "No VMs found on $hostOrClusterName." -Color Yellow
        }

    } catch {
        $errorMessage = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Failed to get VMs: `n$errorMessage", "Error", "OK", "Error")
        Add-Log -Message "Failed to get VMs: $errorMessage" -Color Red
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# --- === CLONE BUTTON MAIN CODE BLOCK === ---
$btnClone.Add_Click({
    $ConfirmPreference = 'None'

    # --- 1. Validate Form Input ---
    if ([string]::IsNullOrWhiteSpace($cboVMs.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a VM to clone.", "Missing Info", "OK", "Warning")
        $form.DialogResult = [System.Windows.Forms.DialogResult]::None; return
    }
    if ([string]::IsNullOrWhiteSpace($txtNewName.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a name for the new VM.", "Missing Info", "OK", "Warning")
        $form.DialogResult = [System.Windows.Forms.DialogResult]::None; return
    }
    if ([string]::IsNullOrWhiteSpace($cboTargetHost.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select or enter a target host.", "Missing Info", "OK", "Warning")
        $form.DialogResult = [System.Windows.Forms.DialogResult]::None; return
    }
    if ($chkCluster.Checked -and [string]::IsNullOrWhiteSpace($cboCSV.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a target CSV for the cluster deployment.", "Missing Info", "OK", "Warning")
        $form.DialogResult = [System.Windows.Forms.DialogResult]::None; return
    }
    if (-not [string]::IsNullOrWhiteSpace($txtVlan.Text.Trim())) {
        $parsedVlan = 0
        if (-not [int]::TryParse($txtVlan.Text.Trim(), [ref]$parsedVlan) -or $parsedVlan -lt 0 -or $parsedVlan -gt 4094) {
            [System.Windows.Forms.MessageBox]::Show("VLAN ID must be a valid integer between 1 and 4094 (or 0 for untagged).", "Invalid VLAN ID", "OK", "Warning")
            $form.DialogResult = [System.Windows.Forms.DialogResult]::None; return
        }
    }

    # --- Pre-Flight Sysprep Checklist Confirmation ---
    $sysprepMsg = "CRITICAL PRE-FLIGHT CHECK:`n`n" +
                  "Has the source template VM ('$($cboVMs.Text)') been generalized via Sysprep?`n`n" +
                  "If not, boot into the template OS and run:`n" +
                  "C:\Windows\System32\Sysprep\Sysprep.exe /generalize /oobe /shutdown`n`n" +
                  "Do you want to continue with the clone process?"
    
    $sysprepCheck = [System.Windows.Forms.MessageBox]::Show($sysprepMsg, "Verify Sysprep State", "YesNo", "Question", "Button2")
    if ($sysprepCheck -eq "No") {
        $form.DialogResult = [System.Windows.Forms.DialogResult]::None
        return 
    }

    # --- 2. Set up Logging ---
    $logDir = "C:\hbs"
    $global:logFile = $null
    if (-not (Test-Path -Path $logDir)) {
        try { $null = New-Item -ItemType Directory -Path $logDir -ErrorAction Stop }
        catch { [System.Windows.Forms.MessageBox]::Show("Failed to create log directory. Logs will not be saved.", "Log Error", "OK", "Warning") }
    }
    if (Test-Path -Path $logDir) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $global:logFile = Join-Path -Path $logDir -ChildPath "CloneLog_$($txtNewName.Text)_$($timestamp).log"
    }

    # --- 3. Disable controls and set cursor ---
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnClone.Enabled = $false
    $btnGetVMs.Enabled = $false
    $btnCancel.Enabled = $false
    $tabControl.Enabled = $false
    $txtLog.Text = ""

    # --- 4. Get variables from form ---
    $hostOrCluster = $txtHost.Text.Trim()
    $cluster = $chkCluster.Checked
    $cloneSource = $cboVMs.Text
    $newVMName = $txtNewName.Text.Trim()
    $targetHost = $cboTargetHost.Text.Trim()
    $selectedSwitch = $cboVMSwitch.Text
    $vlanId = $txtVlan.Text.Trim()
    $selectedCsvPath = $null
    if ($cluster -and -not [string]::IsNullOrWhiteSpace($cboCSV.Text)) {
        if ($script:csvMap.ContainsKey($cboCSV.Text)) {
            $selectedCsvPath = $script:csvMap[$cboCSV.Text].TrimEnd('\')
        } else {
            $selectedCsvPath = $cboCSV.Text.Trim().TrimEnd('\')
        }
    }
    $isTestMode = $chkTestMode.Checked

    $doHardwareConfig = $chkHardware.Checked
    $newCPU = $txtCPU.Text
    $newMemoryGB = $txtMemory.Text

    $doGuestConfig = $chkGuestConfig.Checked
    $doSetHostname = $chkHostname.Checked
    $doSetIP = $chkSetIP.Checked
    $newIP = $txtIP.Text.Trim()
    $newSubnet = $txtSubnet.Text.Trim()
    $newGateway = $txtGateway.Text.Trim()
    $newDNS1 = $txtDNS.Text.Trim()
    $newDNS2 = $txtDNS2.Text.Trim()

    $doDomainJoin = $chkDomainJoin.Checked
    $newDomain = $txtDomain.Text.Trim()

    if ($isTestMode) { Add-Log -Message "--- RUNNING IN TEST MODE ---" -Color Yellow }
    $sourceHost = $hostOrCluster
    $hvHost = $targetHost

    # --- === SCRIPT PHASES START HERE === ---
    try {
        # --- === PRE-FLIGHT CONFLICT CHECKS === ---
        Add-Log -Message "--- RUNNING PRE-FLIGHT CHECKS ---" -Color Cyan

        # 1. Check for duplicate VM name
        Add-Log -Message "Checking if VM name '$newVMName' already exists..." -Color Cyan
        $existingVM = $null
        if ($cluster) {
            $existingVM = Get-ClusterGroup -Cluster $hostOrCluster -ErrorAction SilentlyContinue | Where-Object { $_.GroupType -eq "VirtualMachine" -and $_.Name -eq $newVMName }
        }
        if (-not $existingVM) {
            $existingVM = Get-VM -Name $newVMName -ComputerName $targetHost -ErrorAction SilentlyContinue
        }
        if ($existingVM) {
            throw "A VM named '$newVMName' already exists on '$hostOrCluster'! Please choose a different name."
        }
        Add-Log -Message "VM name '$newVMName' is available." -Color Green

        # 2. Validate Static IP settings & check if IP is already responding to ping
        if ($doGuestConfig -and $doSetIP) {
            Add-Log -Message "Validating network settings and checking IP availability..." -Color Cyan
            try {
                $ipAddr = [System.Net.IPAddress]::Parse($newIP)
                $subnetMask = [System.Net.IPAddress]::Parse($newSubnet)
                $gatewayAddr = [System.Net.IPAddress]::Parse($newGateway)
                $ipBytes = $ipAddr.GetAddressBytes()
                $maskBytes = $subnetMask.GetAddressBytes()
                $networkBytes = for($i=0; $i -lt $ipBytes.Length; $i++) { $ipBytes[$i] -band $maskBytes[$i] }
                $networkAddr = [System.Net.IPAddress]::new($networkBytes)
                $invertedMaskBytes = $maskBytes | ForEach-Object { $_ -bxor 255 }
                $broadcastBytes = for($i=0; $i -lt $networkBytes.Length; $i++) { $networkBytes[$i] -bor $invertedMaskBytes[$i] }
                $broadcastAddr = [System.Net.IPAddress]::new($broadcastBytes)
                $gatewayBytes = $gatewayAddr.GetAddressBytes()
                $gatewayOnNetwork = $true
                for ($i=0; $i -lt $gatewayBytes.Length; $i++) {
                    if (($gatewayBytes[$i] -band $maskBytes[$i]) -ne $networkBytes[$i]) {
                        $gatewayOnNetwork = $false; break
                    }
                }
                if (-not $gatewayOnNetwork -or $gatewayAddr.Equals($networkAddr) -or $gatewayAddr.Equals($broadcastAddr)) {
                     throw "Gateway address '$($gatewayAddr.IPAddressToString)' is not on the same network segment as IP '$($ipAddr.IPAddressToString)' with subnet '$($subnetMask.IPAddressToString)'."
                }
                Add-Log -Message "Network settings are valid." -Color Green
            } catch { throw "Invalid network configuration provided: $($_.Exception.Message)" }

            Add-Log -Message "Pinging $newIP to check for IP conflicts..." -Color Cyan
            if (Test-Connection -ComputerName $newIP -Count 2 -Quiet -ErrorAction SilentlyContinue) {
                Add-Log -Message "WARNING: $newIP responded to ping!" -Color Yellow
                $ipConflictChoice = [System.Windows.Forms.MessageBox]::Show(
                    "WARNING: The IP address '$newIP' is already responding to ping on the network!`n`nContinuing may cause an IP address conflict.`n`nDo you still want to continue?",
                    "IP Conflict Detected",
                    "YesNo",
                    "Warning",
                    "Button2"
                )
                if ($ipConflictChoice -eq "No") {
                    Add-Log -Message "Clone cancelled by user due to IP conflict on $newIP." -Color Yellow
                    return
                }
                Add-Log -Message "Continuing deployment despite ping response on $newIP." -Color Yellow
            } else {
                Add-Log -Message "IP $newIP did not respond to ping (available)." -Color Green
            }
        }

        # --- === PHASE 1: CLONE VM (VM OFF) === ---
        Add-Log -Message "--- PHASE 1: CLONING VM ---" -Color Cyan

        if($cluster) {
            Add-Log -Message "Cluster detected. Finding owner node for $cloneSource..." -Color Cyan
            $sourceHost = (Get-ClusterGroup -Name $cloneSource -Cluster $hostOrCluster).OwnerNode.Name
            Add-Log -Message "Source VM is hosted on $sourceHost. Selected target host is $targetHost." -Color Cyan

            $localShortName = ($env:ComputerName -split '\.')[0]
            $targetShortName = ($targetHost -split '\.')[0]
            if ($doGuestConfig -and ($localShortName -ne $targetShortName)) {
                $clusterNodeNames = @(Get-ClusterNode -Cluster $hostOrCluster -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
                if ($clusterNodeNames -contains $localShortName) {
                    Add-Log -Message "Guest OS config is enabled: Staging VM locally on $localShortName for PowerShell Direct, then moving to $targetHost." -Color Cyan
                    $hvHost = $localShortName
                }
            }
        } else {
            Add-Log -Message "Source host: $sourceHost | Target host: $hvHost" -Color Cyan
        }

        $vm = Get-VM -Name $cloneSource -ComputerName $sourceHost

        if(!$vm) {
            Add-Log -Message "Source VM with name '$cloneSource' could not be found!" -Color Red
            [System.Windows.Forms.MessageBox]::Show("Source VM with name '$cloneSource' could not be found!", "Error", "OK", "Error")
            return
        }

        if($vm.State -ne "Off") {
            Add-Log -Message "VM Template '$cloneSource' is not powered off, cloning cannot continue!" -Color Red
            [System.Windows.Forms.MessageBox]::Show("VM Template '$cloneSource' is not powered off, cloning cannot continue!", "Error", "OK", "Error")
            return
        } else {
            Add-Log -Message "Source VM $cloneSource is 'Off'. Continuing..." -Color Green
        }

        if ($cluster -and $selectedCsvPath) {
            $baseVmPath = $selectedCsvPath
            $baseDiskPath = $selectedCsvPath
            $exportPath = "$baseVmPath\Exports"
        } else {
            $hostObject = Get-VMHost -ComputerName $hvHost
            $baseVmPath = $hostObject.VirtualMachinePath.TrimEnd('\')
            $baseDiskPath = $hostObject.VirtualHardDiskPath.TrimEnd('\')
            $exportPath = "$baseVmPath\Exports"
        }

        Add-Log -Message "Using VM Path: $baseVmPath"
        Add-Log -Message "Using Disk Path: $baseDiskPath"
        Add-Log -Message "Using Export Path: $exportPath"

        $reuseExisting = $false
        $existingVmcx = Get-ChildItem -Recurse -Path "$exportPath\$cloneSource" -Filter "*.vmcx" -ErrorAction SilentlyContinue
        if($existingVmcx) {
            $dialogResult = [System.Windows.Forms.MessageBox]::Show("An export already lives at '$exportPath'! Do you want to RE-USE this existing export?", "Existing Export Found", "YesNo", "Warning")
            if($dialogResult -eq "Yes") {
                $reuseExisting = $true
            } else {
                Add-Log -Message "Removing existing export!" -Color Yellow
                if (-not $isTestMode) { Remove-Item -Path "$exportPath\$cloneSource" -Recurse -Force }
                else { Add-Log -Message "[TEST MODE] Would remove folder: $exportPath\$cloneSource" -Color Yellow }
            }
        }

        $diskPath = "$baseVmPath\$newVMName\Virtual Hard Disks"
        $vmPath = "$baseVmPath\$newVMName"
        Add-Log -Message "New VM files will be at: $vmPath"
        Add-Log -Message "New VM disks will be at: $diskPath"

        if(!$reuseExisting) {
            Add-Log -Message "--Exporting source VM from $sourceHost..." -Color Cyan
            if (-not $isTestMode) { $null = Export-VM -ComputerName $sourceHost -Name $cloneSource -Path $exportPath }
            else { Add-Log -Message "[TEST MODE] Would export VM $cloneSource from $sourceHost to $exportPath" -Color Yellow }
        }

        $exportVmPath = "$exportPath\$cloneSource\Virtual Machines"
        $exportVmxFile = Get-ChildItem -Path $exportVmPath -Filter "*.vmcx" -ErrorAction SilentlyContinue
        if ($isTestMode -and -not $exportVmxFile) {
            Add-Log -Message "[TEST MODE] Simulating a found VMCX file for import." -Color Yellow
            $exportVmxFile = [PSCustomObject]@{ FullName = "$exportVmPath\SIMULATED-FILE.vmcx" }
        }

        if($reuseExisting) { Add-Log -Message "--Importing VM from existing export to $hvHost..." -Color Cyan }
        else { Add-Log -Message "--Importing VM from new export to $hvHost..." -Color Cyan }

        $importedVM = $null
        if (-not $isTestMode) {
            $importedVM = Import-VM -ComputerName $hvHost -Path $exportVmxFile.FullName -Copy -GenerateNewId -VirtualMachinePath $vmPath -VhdDestinationPath $diskPath
            Sleep -Seconds 3
            Add-Log -Message "--Import complete, renaming VM..." -Color Cyan
            $null = Rename-VM -VM $importedVM -NewName $newVMName
            $importedVM = Get-VM $newVMName -ComputerName $hvHost
        } else {
            Add-Log -Message "[TEST MODE] Would import VM from $($exportVmxFile.FullName) onto host $hvHost" -Color Yellow
            Add-Log -Message "[TEST MODE] Would rename imported VM to $newVMName" -Color Yellow
        }

        $vmDisks = $null
        if ($isTestMode) {
            Add-Log -Message "[TEST MODE] Getting disk list from source VM '$($vm.Name)'..." -Color Yellow
            $vmDisks = $vm | Get-VMHardDiskDrive
        } else {
            $vmDisks = $importedVM | Get-VMHardDiskDrive
        }

        foreach($disk in $vmDisks) {
            $newDiskShortName = "$($newVMName)_$($disk.ControllerNumber)_$($disk.ControllerLocation).vhdx"
            $newNameFullPath = "$diskPath\$newDiskShortName"
            Add-Log -Message "    Renaming disk $($disk.Path) to $newDiskShortName"
            if (-not $isTestMode) {
                $item = Rename-Item -Path $disk.Path -NewName "$newDiskShortName"
                $disk | Set-VMHardDiskDrive -Path $newNameFullPath
            } else {
                Add-Log -Message "[TEST MODE] Would rename disk $($disk.Path) to $newNameFullPath" -Color Yellow
            }
        }

        Add-Log -Message "--Resetting VM UUID and Serial Numbers" -Color Cyan
        $biosGuid = "{$(([System.Guid]::NewGuid()).Guid.ToUpper())}"
        $serial = (1..6 | ForEach { (Get-Random -Minimum 1 -Maximum 9999).ToString().PadLeft(4,'0') }) -join "-"
        $serial += "-$((Get-Random -Minimum 1 -Maximum 99).ToString().PadLeft(2,'0'))"

        if (-not $isTestMode) {
            $MSVM = gwmi -Namespace root\virtualization\v2 -Class msvm_computersystem -Filter "ElementName = '$newVMName'" -ComputerName $hvHost
            $MSVMSystemSettings = $MSVM.GetRelated('msvm_virtualsystemsettingdata') | Select-Object -First 1
            $MSVMSystemSettings['BIOSGUID'] = $biosGuid
            $MSVMSystemSettings['BaseboardSerialNumber'] = $serial
            $MSVMSystemSettings['ChassisAssetTag'] = $serial
            $MSVMSystemSettings['ChassisSerialNumber'] = $serial
            $MSVMSystemSettings['BIOSSerialNumber'] = $serial
            $VMMS = gwmi -Namespace root\virtualization\v2 -Class msvm_virtualsystemmanagementservice -ComputerName $hvHost
            $ModifySystemSettingsParameters = $VMMS.GetMethodParameters('ModifySystemSettings')
            $ModifySystemSettingsParameters['SystemSettings'] = $MSVMSystemSettings.GetText([System.Management.TextFormat]::CimDtd20)
            $wmiresult = $VMMS.InvokeMethod('ModifySystemSettings', $ModifySystemSettingsParameters, $null)
            if($wmiresult.ReturnValue -ne 0) { Add-Log -Message "!! Resetting UUID/Serial failed" -Color Red }
            else { Add-Log -Message "    New BIOS GUID/Serial set." }
        } else {
            $biosGuid = "[TEST BIOS GUID]" 
            Add-Log -Message "[TEST MODE] Would set New BIOS GUID: $biosGuid" -Color Yellow
            Add-Log -Message "[TEST MODE] Would set New Serial: $serial" -Color Yellow
        }

        if($cluster) {
            Add-Log -Message "--Adding VM to Cluster on host $hvHost" -Color Cyan
            if (-not $isTestMode) {
                if (-not $importedVM) { $importedVM = Get-VM $newVMName -ComputerName $hvHost }
                $null = Add-ClusterVirtualMachineRole -VMId $importedVM.VMId -Name $newVMName -Cluster $hostOrCluster
            } else {
                Add-Log -Message "[TEST MODE] Would add VM $newVMName to cluster $hostOrCluster on host $hvHost" -Color Yellow
            }
        }

        # --- === PHASE 2: HARDWARE & NETWORK ADAPTER CONFIG (VM OFF) === ---
        $doSwitchChange = (-not [string]::IsNullOrWhiteSpace($selectedSwitch) -and $selectedSwitch -ne "<Keep Template Switch>")
        $doVlanChange = (-not [string]::IsNullOrWhiteSpace($vlanId))

        if ($doHardwareConfig -or $doSwitchChange -or $doVlanChange) {
            Add-Log -Message "--- PHASE 2: CONFIGURING VM HARDWARE & NETWORK (VM OFF) ---" -Color Cyan

            if (-not $isTestMode -and -not $importedVM) {
                $importedVM = Get-VM $newVMName -ComputerName $hvHost
            }

            if ($doHardwareConfig) {
                if ($newCPU -notmatch '^\d+$') { throw "vCPU count '$newCPU' is not a valid number." }
                if ($newMemoryGB -notmatch '^\d+$') { throw "Memory (GB) '$newMemoryGB' is not a valid number." }

                if (-not $isTestMode) {
                    Add-Log -Message "    Setting vCPU count to $newCPU"
                    Set-VMProcessor -VM $importedVM -Count ([int]$newCPU)

                    $startupMemoryBytes = [int64]$newMemoryGB * 1GB
                    Add-Log -Message "    Setting Startup Memory to ${newMemoryGB}GB"
                    Set-VMMemory -VM $importedVM -StartupBytes $startupMemoryBytes
                } else {
                    Add-Log -Message "[TEST MODE] Would set vCPU count to $newCPU" -Color Yellow
                    Add-Log -Message "[TEST MODE] Would set Startup Memory to ${newMemoryGB}GB" -Color Yellow
                }
            }

            if ($doSwitchChange) {
                if (-not $isTestMode) {
                    Add-Log -Message "    Connecting network adapter to Virtual Switch '$selectedSwitch'..."
                    Get-VMNetworkAdapter -VM $importedVM | Connect-VMNetworkAdapter -SwitchName $selectedSwitch
                } else {
                    Add-Log -Message "[TEST MODE] Would connect network adapter to Virtual Switch '$selectedSwitch'" -Color Yellow
                }
            }

            if ($doVlanChange) {
                $vlanInt = [int]$vlanId
                if (-not $isTestMode) {
                    if ($vlanInt -gt 0) {
                        Add-Log -Message "    Setting VLAN ID to $vlanInt (Access mode)..."
                        Set-VMNetworkAdapterVlan -VM $importedVM -Access -VlanId $vlanInt
                    } else {
                        Add-Log -Message "    Setting network adapter to Untagged..."
                        Set-VMNetworkAdapterVlan -VM $importedVM -Untagged
                    }
                } else {
                    Add-Log -Message "[TEST MODE] Would set VLAN ID to $vlanInt" -Color Yellow
                }
            }
        }

        # --- === PHASE 3: GUEST OS CONFIG (VM ON) === ---
        if ($doGuestConfig) {
            Add-Log -Message "--- PHASE 3: CONFIGURING GUEST OS (VM ON) ---" -Color Cyan

            if ($null -eq $script:savedLocalAdminCred) {
                $script:savedLocalAdminCred = Get-Credential -UserName "Administrator" -Message "Enter LOCAL Admin password for '$cloneSource' template (Saved for session)"
                if (-not $script:savedLocalAdminCred) { throw "Local admin credentials not provided. Aborting guest config." }
                Add-Log -Message "    Saved template local admin credentials for this session." -Color Cyan
            } else {
                Add-Log -Message "    Using saved session credentials for template local admin ($($script:savedLocalAdminCred.UserName))." -Color Cyan
            }
            $localAdminCred = $script:savedLocalAdminCred

            $domainCred = $null
            if ($doDomainJoin) {
                if ($null -eq $script:savedDomainCred -or $script:savedDomainName -ne $newDomain) {
                    $script:savedDomainCred = Get-Credential -UserName "$newDomain\Administrator" -Message "Enter credentials for joining '$newDomain' (Saved for session)"
                    if (-not $script:savedDomainCred) { throw "Domain join credentials not provided. Aborting guest config." }
                    $script:savedDomainName = $newDomain
                    Add-Log -Message "    Saved domain credentials ($($script:savedDomainCred.UserName)) for this session." -Color Cyan
                } else {
                    Add-Log -Message "    Using saved session credentials for domain '$newDomain' ($($script:savedDomainCred.UserName))." -Color Cyan
                }
                $domainCred = $script:savedDomainCred
            }

            # --- DYNAMIC OFFLINE UNATTEND INJECTION ---
            Add-Log -Message "Generating dynamic unattend.xml using provided local credentials..." -Color Cyan
            $plainPassword = $localAdminCred.GetNetworkCredential().Password
            
            $unattendXML = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <OOBE>
                <HideEULAPage>true</HideEULAPage>
                <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
                <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
                <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
                <NetworkLocation>Work</NetworkLocation>
                <ProtectYourPC>3</ProtectYourPC>
            </OOBE>
            <UserAccounts>
                <AdministratorPassword>
                    <Value>$plainPassword</Value>
                    <PlainText>true</PlainText>
                </AdministratorPassword>
            </UserAccounts>
        </component>
        <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <InputLocale>en-US</InputLocale>
            <SystemLocale>en-US</SystemLocale>
            <UILanguage>en-US</UILanguage>
            <UserLocale>en-US</UserLocale>
        </component>
    </settings>
</unattend>
"@

            if (-not $isTestMode) {
                try {
                    Add-Log -Message "Mounting cloned VHDX offline: $newNameFullPath" -Color Cyan
                    $vhdDisk = Mount-VHD -Path $newNameFullPath -Passthru
                    Start-Sleep -Seconds 3
                    
                    $driveLetter = $null
                    $volumes = $vhdDisk | Get-Disk | Get-Partition | Get-Volume
                    foreach ($vol in $volumes) {
                        if ($vol.DriveLetter -and (Test-Path "$($vol.DriveLetter):\Windows")) {
                            $driveLetter = $vol.DriveLetter
                            break
                        }
                    }

                    if (-not $driveLetter) {
                        throw "Could not locate the Windows directory volume on the mounted VHDX."
                    }

                    $pantherPath = "$($driveLetter):\Windows\Panther"
                    if (-not (Test-Path $pantherPath)) {
                        $null = New-Item -ItemType Directory -Path $pantherPath -Force
                    }

                    $unattendXML | Out-File -FilePath "$pantherPath\unattend.xml" -Encoding utf8 -Force
                    Add-Log -Message "Successfully injected unattend.xml into target OS volume ($pantherPath)" -Color Green
                }
                catch {
                    Add-Log -Message "VHDX Injection Failed: $($_.Exception.Message)" -Color Red
                    throw $_
                }
                finally {
                    Add-Log -Message "Dismounting VHDX..."
                    Dismount-VHD -Path $newNameFullPath
                    Start-Sleep -Seconds 2
                }
            } else {
                Add-Log -Message "[TEST MODE] Would mount $newNameFullPath and inject dynamic XML into Windows\Panther" -Color Yellow
            }

            Function Wait-VMReady {
                param($VMName, $Cred)
                Add-Log -Message "    Waiting for Guest OS on '$VMName' (Host: $hvHost) to become responsive... (May take 2-4 minutes)"
                $timeout = (New-TimeSpan -Minutes 6)
                $watch = [System.Diagnostics.Stopwatch]::StartNew()
                $vmReady = $false
                $lastError = ""
                while ($watch.Elapsed -lt $timeout -and -not $vmReady) {
                    $heartbeat = Get-VMIntegrationService -VMName $VMName -ComputerName $hvHost -Name "Heartbeat" -ErrorAction SilentlyContinue
                    if (-not $heartbeat -or ($heartbeat.OperationalStatus -notcontains "OK")) {
                        Add-Log -Message "    Waiting for VM Heartbeat..."
                        Start-Sleep -Seconds 10
                        continue
                    }
                    try {
                        Add-Log -Message "    Heartbeat 'OK'. Probing for PS Direct..."
                        $null = Invoke-Command -VMName $VMName -Credential $Cred -ScriptBlock { $true } -ErrorAction Stop
                        $vmReady = $true
                        Add-Log -Message "    Guest OS is responsive!" -Color Green
                    } catch {
                        $lastError = $_.Exception.Message
                        Add-Log -Message "    PS Direct not ready... Retrying."
                        Start-Sleep -Seconds 10
                    }
                }
                if (-not $vmReady) { throw "Timed out waiting for Guest OS to become responsive to PowerShell Direct. Last error: $lastError" }
            }

            if (-not $isTestMode) {
                Add-Log -Message "    Starting VM $newVMName on $hvHost..."
                Start-VM -Name $newVMName -ComputerName $hvHost
                Wait-VMReady -VMName $newVMName -Cred $localAdminCred
            } else {
                Add-Log -Message "[TEST MODE] Would start VM $newVMName on $hvHost" -Color Yellow
                Add-Log -Message "[TEST MODE] Would wait for Guest Services & PS Direct" -Color Yellow
            }

            $hostnameRestartRequired = $false
            if ($doSetHostname) {
                Add-Log -Message "    Preparing Guest OS script block (Hostname)..."
                $scriptBlock_Hostname = {
                    param($vmName)
                    $InformationPreference = 'Continue'
                    Write-Host "--- GUEST SCRIPT (Hostname) START ---"
                    Write-Host "TASK: Setting Hostname to $vmName..."
                    Rename-Computer -NewName $vmName -Force
                    Write-Host "SUCCESS: Hostname set."
                    Write-Host "--- GUEST SCRIPT (Hostname) END ---"
                    Write-Host "GUEST: Initiating restart..."
                    Restart-Computer -Force
                }

                if (-not $isTestMode) {
                    Add-Log -Message "    Injecting Guest OS script (Hostname & Restart)..."
                    $guestLogStream = New-Object System.Collections.Generic.List[System.Management.Automation.InformationRecord]
                    try {
                        Invoke-Command -VMName $newVMName -Credential $localAdminCred -ScriptBlock $scriptBlock_Hostname -ArgumentList $newVMName -InformationVariable +guestLogStream -ErrorAction Stop
                    } catch {
                        if ($_.Exception.InnerException -is [System.Management.Automation.Remoting.PSRemotingTransportException] -or $_.FullyQualifiedErrorId -like '*PSSessionStateBroken*') {
                             Add-Log -Message "    [GUEST-H] Restart initiated. Session closed as expected." -Color Green
                        } else {
                             Add-Log -Message "    [GUEST-H] ERROR: $($_.Exception.Message)" -Color Red
                             throw "Guest script (Hostname) failed. See log for details."
                        }
                    } finally { $guestLogStream | ForEach-Object { Add-Log -Message "    [GUEST-H] $($_.MessageData)" } }
                    Add-Log -Message "    Hostname configuration sent. Waiting for VM to restart..." -Color Cyan
                    Start-Sleep -Seconds 15
                    Wait-VMReady -VMName $newVMName -Cred $localAdminCred
                    $hostnameRestartRequired = $true
                } else {
                    Add-Log -Message "[TEST MODE] Would set Hostname: $newVMName" -Color Yellow
                    Add-Log -Message "[TEST MODE] Would tell Guest to restart and wait." -Color Yellow
                    $hostnameRestartRequired = $true
                }
            }

            $ipDomainRestartRequired = $false
            if ($doSetIP -or $doDomainJoin) {
                Add-Log -Message "    Preparing Guest OS script block (IP/Domain)..."
                $scriptBlock_IPDomain = {
                    param($doIP, $ip, $subnet, $gw, $dns1, $dns2, $doDomain, $domain, $dCred)
                    $InformationPreference = 'Continue'
                    Write-Host "--- GUEST SCRIPT (IP/Domain) START ---"
                    $needsRestart = $false

                    if ($doIP) {
                        Write-Host "TASK: Setting IP to $ip..."
                        Write-Host "Enabling physical network adapters..."
                        Get-NetAdapter | Where-Object { $_.Virtual -eq $false } | Enable-NetAdapter -Confirm:$false
                        $adapter = Get-NetAdapter | Where-Object { $_.Virtual -eq $false -and $_.Status -ne 'Disconnected' } | Select-Object -First 1

                        if ($adapter) {
                            Write-Host "Found adapter: $($adapter.Name). Configuring..."
                            try {
                                Remove-NetIPAddress -InterfaceAlias $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue
                                $binarySubnet = $subnet.Split('.') | ForEach-Object { [Convert]::ToString([int]$_, 2).PadLeft(8,'0') }
                                $prefix = (-join $binarySubnet).Replace('0','').Length
                                Write-Host "Calculated PrefixLength: $prefix (from $subnet)"
                                New-NetIPAddress -InterfaceAlias $adapter.Name -IPAddress $ip -PrefixLength $prefix -DefaultGateway $gw -ErrorAction Stop
                                Write-Host "IP, Subnet, and Gateway set."
                                $dnsServers = @($dns1)
                                if (-not [string]::IsNullOrWhiteSpace($dns2)) { $dnsServers += $dns2 }
                                Write-Host "Setting DNS Servers: $($dnsServers -join ', ')"
                                Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses $dnsServers -ErrorAction Stop
                                Write-Host "DNS set."
                                Write-Host "SUCCESS: Network configured."
                            } catch { throw "Failed to set IP/DNS on adapter $($adapter.Name). Error: $_" }
                        } else { throw "Could not find a valid, connected network adapter to configure." }
                    }

                    if ($doDomain) {
                        if ($doIP) {
                            Write-Host "Pausing 10 seconds for network settings to apply..."
                            Start-Sleep -Seconds 10
                            Write-Host "Registering DNS..."
                            ipconfig /registerdns
                        }
                        Write-Host "TASK: Joining domain $domain..."
                        Add-Computer -DomainName $domain -Credential $dCred -Force -ErrorAction Stop
                        $needsRestart = $true
                        Write-Host "SUCCESS: Domain joined."
                    }

                    Write-Host "--- GUEST SCRIPT (IP/Domain) END ---"

                    if ($needsRestart) {
                        Write-Host "GUEST: Initiating final restart..."
                        Restart-Computer -Force
                    }
                }

                $sbArgs_IPDomain = @(
                    $doSetIP, $newIP, $newSubnet, $newGateway, $newDNS1, $newDNS2,
                    $doDomainJoin, $newDomain, $domainCred
                )

                if (-not $isTestMode) {
                    Add-Log -Message "    Injecting Guest OS script (IP/Domain)..."
                    $guestLogStream = New-Object System.Collections.Generic.List[System.Management.Automation.InformationRecord]
                    $ipDomainRestartInitiated = $false
                    try {
                        Invoke-Command -VMName $newVMName -Credential $localAdminCred -ScriptBlock $scriptBlock_IPDomain -ArgumentList $sbArgs_IPDomain -InformationVariable +guestLogStream -ErrorAction Stop
                        if ($doDomainJoin) { $ipDomainRestartInitiated = $true }
                    } catch {
                        if ($doDomainJoin -and ($_.Exception.InnerException -is [System.Management.Automation.Remoting.PSRemotingTransportException] -or $_.FullyQualifiedErrorId -like '*PSSessionStateBroken*')) {
                            Add-Log -Message "    [GUEST-IP/D] Restart initiated by domain join. Session closed as expected." -Color Green
                            $ipDomainRestartInitiated = $true
                        } else {
                            Add-Log -Message "    [GUEST-IP/D] ERROR: $($_.Exception.Message)" -Color Red
                            throw "Guest script (IP/Domain) failed. See log for details."
                        }
                    } finally { $guestLogStream | ForEach-Object { Add-Log -Message "    [GUEST-IP/D] $($_.MessageData)" } }
                    Add-Log -Message "    IP/Domain configuration complete." -Color Green

                    Add-Log -Message "    Waiting 5s for final operations..." -Color Cyan
                    Start-Sleep -Seconds 5

                    if ($ipDomainRestartInitiated) {
                         Add-Log -Message "    Guest OS initiated final restart. Process complete." -Color Green
                    } else {
                         Add-Log -Message "    Guest OS settings applied. No final restart needed."
                         Add-Log -Message "    Shutting down VM $newVMName (graceful)."
                         Stop-VM -Name $newVMName -ComputerName $hvHost -Confirm:$false
                    }

                } else {
                    Add-Log -Message "[TEST MODE] Would run IP/Domain script block." -Color Yellow
                    if ($doDomainJoin) { Add-Log -Message "[TEST MODE] Guest would initiate final restart." -Color Yellow }
                }
            }
            elseif ($hostnameRestartRequired) {
                 Add-Log -Message "    VM already restarted for hostname. No further Guest OS actions requested." -Color Green
            }
            else {
                 Add-Log -Message "    No Guest OS configuration requested or needed." -Color Green
                 if ((Get-VM -Name $newVMName -ComputerName $hvHost).State -ne 'Off') {
                    #Add-Log -Message "    Shutting down VM $newVMName (graceful)."
                    #Stop-VM -Name $newVMName -Confirm:$false
                 }
            }

            if ($cluster -and (($hvHost -split '\.')[0] -ne ($targetHost -split '\.')[0])) {
                Add-Log -Message "    Moving VM '$newVMName' from $hvHost to selected target host '$targetHost'..." -Color Cyan
                if (-not $isTestMode) {
                    $null = Move-ClusterVirtualMachineRole -Name $newVMName -Node $targetHost -Cluster $hostOrCluster -ErrorAction Stop
                    $hvHost = $targetHost
                    Add-Log -Message "    Successfully moved '$newVMName' to '$targetHost'." -Color Green
                } else {
                    Add-Log -Message "[TEST MODE] Would move clustered VM '$newVMName' to '$targetHost'" -Color Yellow
                }
            }
        }

        # --- === PHASE 4: CLEANUP === ---
        Add-Log -Message "--- PHASE 4: CLEANUP ---" -Color Cyan
        $cleanupResult = [System.Windows.Forms.MessageBox]::Show("Do you want to delete the export of your template?", "Cleanup Export?", "YesNo", "Question")

        if($cleanupResult -eq "Yes") {
            Add-Log -Message "Removing export!" -Color Yellow
            if (-not $isTestMode) { Remove-Item -Path "$exportPath\$cloneSource" -Recurse -Force }
            else { Add-Log -Message "[TEST MODE] Would remove export folder: $exportPath\$cloneSource" -Color Yellow }
        }

        Add-Log -Message "Clone process is complete!" -Color Green
        [System.Windows.Forms.MessageBox]::Show("Clone process is complete!", "Success", "OK", "Information")

    } catch {
        $errorMessage = $_.Exception.Message
        Add-Log -Message "A critical error occurred: $errorMessage" -Color Red
        [System.Windows.Forms.MessageBox]::Show("A critical error occurred: `n$errorMessage", "Script Error", "OK", "Error")
    } finally {
        if ($global:logFile -and (Test-Path $global:logFile)) {
            Add-Log -Message "Log file saved to: $global:logFile" -Color Green
        }
        # --- Re-enable controls ---
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnClone.Enabled = $true
        $btnGetVMs.Enabled = $true
        $btnCancel.Enabled = $true
        $tabControl.Enabled = $true
    }

    $form.DialogResult = [System.Windows.Forms.DialogResult]::None
})

# --- Show the Form ---
$form.ShowDialog()

# --- Clean up resources ---
$form.Dispose()
