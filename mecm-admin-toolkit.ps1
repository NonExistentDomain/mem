#Requires -Version 3.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    MECM Admin Toolkit - A PowerShell alternative to Recast's MECM Right-Click Tools
.DESCRIPTION
    This script provides a GUI-based tool for managing SCCM/MECM clients, including:
    - Installing/uninstalling applications, programs, software updates, and task sequences
    - Downloading content from the MECM server
    - Performing remote client actions
.NOTES
    File Name: MECM-AdminToolkit.ps1
    Author: NonExistent
    Requires: PowerShell 3.0 or later, .NET Framework 4.0, Run as Administrator
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import Configuration Manager PowerShell module if site server is local
$CMModulePath = Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH -Parent) "ConfigurationManager.psd1"
if (Test-Path $CMModulePath) {
    Import-Module $CMModulePath
}

#region Helper Functions

function Get-MecmSite {
    # Get the MECM site code from registry or site server
    try {
        $SiteCode = Get-WmiObject -Namespace "root\ccm" -Class "SMS_Client" -ErrorAction Stop | Select-Object -ExpandProperty "ClientSiteCode"
        if (-not $SiteCode) {
            throw "Unable to determine site code from client WMI."
        }
        return $SiteCode
    }
    catch {
        # Prompt user for site code
        [System.Windows.Forms.MessageBox]::Show("Unable to detect site code automatically. Please enter it manually in the next dialog.", "Site Code", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "MECM Site Code"
        $form.Size = New-Object System.Drawing.Size(300,150)
        $form.StartPosition = "CenterScreen"
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = "Please enter your MECM Site Code:"
        $form.Controls.Add($label)
        
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = New-Object System.Drawing.Point(10,50)
        $textbox.Size = New-Object System.Drawing.Size(260,20)
        $form.Controls.Add($textbox)
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(75,80)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($okButton)
        $form.AcceptButton = $okButton
        
        $form.Topmost = $true
        $result = $form.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $textbox.Text.Trim()
        }
        else {
            return $null
        }
    }
}

function Get-MecmServer {
    # Get the MECM server from registry or prompt
    try {
        $MecmServer = Get-WmiObject -Namespace "root\ccm" -Class "SMS_Authority" -ErrorAction Stop | Select-Object -ExpandProperty "CurrentManagementPoint"
        if (-not $MecmServer) {
            throw "Unable to determine MECM server from client WMI."
        }
        return $MecmServer
    }
    catch {
        # Prompt user for server name
        [System.Windows.Forms.MessageBox]::Show("Unable to detect MECM server automatically. Please enter it manually in the next dialog.", "MECM Server", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "MECM Server"
        $form.Size = New-Object System.Drawing.Size(300,150)
        $form.StartPosition = "CenterScreen"
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = "Please enter your MECM Server hostname:"
        $form.Controls.Add($label)
        
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = New-Object System.Drawing.Point(10,50)
        $textbox.Size = New-Object System.Drawing.Size(260,20)
        $form.Controls.Add($textbox)
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(75,80)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($okButton)
        $form.AcceptButton = $okButton
        
        $form.Topmost = $true
        $result = $form.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $textbox.Text.Trim()
        }
        else {
            return $null
        }
    }
}

function Install-MecmApplication {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$ApplicationID
    )
    
    try {
        $Application = Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -Filter "Id='$ApplicationID'" -ErrorAction Stop
        
        $Args = @{
            EnforcePreference = [uint32]0
            Id = "$($Application.Id)"
            IsMachineTarget = $true
            IsRebootIfNeeded = $false
            Priority = 'High'
            Revision = "$($Application.Revision)"
        }
        
        Invoke-CimMethod -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -MethodName "Install" -Arguments $Args
        return $true
    }
    catch {
        Write-Error "Failed to install application: $_"
        return $false
    }
}

function Uninstall-MecmApplication {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$ApplicationID
    )
    
    try {
        $Application = Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -Filter "Id='$ApplicationID'" -ErrorAction Stop
        
        $Args = @{
            Id = "$($Application.Id)"
            IsMachineTarget = $true
            IsRebootIfNeeded = $false
            Priority = 'High'
            Revision = "$($Application.Revision)"
        }
        
        Invoke-CimMethod -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -MethodName "Uninstall" -Arguments $Args
        return $true
    }
    catch {
        Write-Error "Failed to uninstall application: $_"
        return $false
    }
}

function Install-MecmPackage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$PackageID,
        
        [Parameter(Mandatory=$true)]
        [string]$ProgramName
    )
    
    try {
        $Args = @{
            PackageID = $PackageID
            ProgramID = $ProgramName
        }
        
        Invoke-CimMethod -ComputerName $ComputerName -Namespace "root\ccm" -ClassName SMS_Client -MethodName "TriggerRunAdvertisement" -Arguments $Args
        return $true
    }
    catch {
        Write-Error "Failed to install package: $_"
        return $false
    }
}

function Install-MecmUpdate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string[]]$UpdateIDs
    )
    
    try {
        # Get CCM_SoftwareUpdate instances
        $Updates = @()
        foreach ($UpdateID in $UpdateIDs) {
            $Update = Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_SoftwareUpdate -Filter "UpdateID='$UpdateID'" -ErrorAction Stop
            $Updates += $Update
        }
        
        # Create update collection
        $UpdateCollection = @($Updates)
        
        # Install updates
        $Args = @{
            CCMUpdates = $UpdateCollection
            EnforcePreference = [uint32]0
            IsRebootIfNeeded = $false
            Priority = 'High'
        }
        
        Invoke-CimMethod -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_SoftwareUpdatesManager -MethodName "InstallUpdates" -Arguments $Args
        return $true
    }
    catch {
        Write-Error "Failed to install updates: $_"
        return $false
    }
}

function Install-MecmTaskSequence {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$PackageID
    )
    
    try {
        $Args = @{
            AdvertID = $PackageID
            Priority = [uint32]"High"
        }
        
        Invoke-CimMethod -ComputerName $ComputerName -Namespace "root\ccm" -ClassName SMS_Client -MethodName "TriggerSchedule" -Arguments $Args
        return $true
    }
    catch {
        Write-Error "Failed to install task sequence: $_"
        return $false
    }
}

function Get-MecmContent {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ContentID,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Application", "Package", "Update", "TaskSequence")]
        [string]$ContentType,
        
        [Parameter(Mandatory=$true)]
        [string]$SiteCode,
        
        [Parameter(Mandatory=$true)]
        [string]$SiteServer
    )
    
    try {
        # Connect to the site server WMI
        $NameSpace = "root\SMS\site_$SiteCode"
        
        # Get content location based on content type
        switch ($ContentType) {
            "Application" {
                $ContentInfo = Get-WmiObject -ComputerName $SiteServer -Namespace $NameSpace -Class SMS_Application -Filter "CI_ID='$ContentID'"
                $ContentLocation = Get-WmiObject -ComputerName $SiteServer -Namespace $NameSpace -Class SMS_ApplicationLatest -Filter "CI_ID='$ContentID'"
                $ContentPath = $ContentLocation.ContentSourcePath
            }
            "Package" {
                $ContentInfo = Get-WmiObject -ComputerName $SiteServer -Namespace $NameSpace -Class SMS_Package -Filter "PackageID='$ContentID'"
                $ContentPath = $ContentInfo.PkgSourcePath
            }
            "Update" {
                $ContentInfo = Get-WmiObject -ComputerName $SiteServer -Namespace $NameSpace -Class SMS_SoftwareUpdate -Filter "CI_ID='$ContentID'"
                # For updates, need to get content location from the content library
                $ContentPath = Get-WmiObject -ComputerName $SiteServer -Namespace $NameSpace -Class SMS_CIContentFiles -Filter "ContentID='$ContentID'" | Select-Object -ExpandProperty SourceURL
            }
            "TaskSequence" {
                $ContentInfo = Get-WmiObject -ComputerName $SiteServer -Namespace $NameSpace -Class SMS_TaskSequencePackage -Filter "PackageID='$ContentID'"
                $ContentPath = $ContentInfo.PkgSourcePath
            }
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Copy content to destination
        if ($ContentPath -match "^\\\\") {
            # Network path
            Copy-Item -Path $ContentPath\* -Destination $DestinationPath -Recurse -Force
        }
        else {
            # Local path on server, use admin share
            $ServerName = $SiteServer.Split(".")[0]
            $DriveLetter = $ContentPath.Substring(0, 1)
            $RemainderPath = $ContentPath.Substring(3)
            $NetworkPath = "\\$ServerName\$DriveLetter$\$RemainderPath"
            
            Copy-Item -Path $NetworkPath\* -Destination $DestinationPath -Recurse -Force
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to get content: $_"
        return $false
    }
}

function Trigger-MecmClientAction {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("MachinePolicy", "DiscoveryData", "SoftwareInventory", "HardwareInventory", "UpdateScan", "UpdateDeployment")]
        [string]$ActionType
    )
    
    try {
        $ScheduleIDs = @{
            "MachinePolicy" = "{00000000-0000-0000-0000-000000000021}"
            "DiscoveryData" = "{00000000-0000-0000-0000-000000000003}"
            "SoftwareInventory" = "{00000000-0000-0000-0000-000000000001}"
            "HardwareInventory" = "{00000000-0000-0000-0000-000000000002}"
            "UpdateScan" = "{00000000-0000-0000-0000-000000000113}"
            "UpdateDeployment" = "{00000000-0000-0000-0000-000000000108}"
        }
        
        $ScheduleID = $ScheduleIDs[$ActionType]
        
        $Args = @{
            ScheduleID = $ScheduleID
        }
        
        Invoke-CimMethod -ComputerName $ComputerName -Namespace "root\ccm" -ClassName SMS_Client -MethodName "TriggerSchedule" -Arguments $Args
        return $true
    }
    catch {
        Write-Error "Failed to trigger client action: $_"
        return $false
    }
}

function Get-AvailableMecmApplications {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    try {
        Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -ErrorAction Stop | 
            Select-Object Id, Name, Publisher, SoftwareVersion, IsInstalled
    }
    catch {
        Write-Error "Failed to get available applications: $_"
        return @()
    }
}

function Get-AvailableMecmPackages {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    try {
        Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_Program -ErrorAction Stop |
            Select-Object PackageID, ProgramID, Name, Description
    }
    catch {
        Write-Error "Failed to get available packages: $_"
        return @()
    }
}

function Get-AvailableMecmUpdates {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    try {
        Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_SoftwareUpdate -Filter "ComplianceState=0" -ErrorAction Stop |
            Select-Object UpdateID, Name, Description, ArticleID
    }
    catch {
        Write-Error "Failed to get available updates: $_"
        return @()
    }
}

function Get-AvailableMecmTaskSequences {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    try {
        Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" -ClassName CCM_TaskSequence -ErrorAction Stop |
            Select-Object PackageID, Name, Description
    }
    catch {
        Write-Error "Failed to get available task sequences: $_"
        return @()
    }
}

#endregion

#region UI Functions

function Show-MainForm {
    [void][System.Windows.Forms.Application]::EnableVisualStyles()
    
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.Text = "MECM Admin Toolkit"
    $MainForm.Size = New-Object System.Drawing.Size(800, 600)
    $MainForm.StartPosition = "CenterScreen"
    $MainForm.Icon = [System.Drawing.SystemIcons]::Application
    
    # Site information
    $SiteCode = Get-MecmSite
    $SiteServer = Get-MecmServer
    
    if (-not $SiteCode -or -not $SiteServer) {
        [System.Windows.Forms.MessageBox]::Show("Unable to determine MECM site information. The application will now exit.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Controls
    $TabControl = New-Object System.Windows.Forms.TabControl
    $TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    
    # Computer selection
    $TopPanel = New-Object System.Windows.Forms.Panel
    $TopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $TopPanel.Height = 60
    
    $ComputerLabel = New-Object System.Windows.Forms.Label
    $ComputerLabel.Text = "Computer(s):"
    $ComputerLabel.Location = New-Object System.Drawing.Point(10, 20)
    $ComputerLabel.AutoSize = $true
    
    $ComputerTextBox = New-Object System.Windows.Forms.TextBox
    $ComputerTextBox.Location = New-Object System.Drawing.Point(110, 20)
    $ComputerTextBox.Size = New-Object System.Drawing.Size(450, 20)
    $ComputerTextBox.Text = $env:COMPUTERNAME
    
    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Location = New-Object System.Drawing.Point(570, 20)
    $BrowseButton.Size = New-Object System.Drawing.Size(90, 23)
    $BrowseButton.Text = "Browse..."
    $BrowseButton.Add_Click({
        $ADPicker = New-Object System.DirectoryServices.UI.DirectoryObjectPickerDialog
        $ADPicker.AllowedObjectTypes = [System.DirectoryServices.UI.DirectoryObjectTypes]::Computers
        $ADPicker.DefaultObjectTypes = [System.DirectoryServices.UI.DirectoryObjectTypes]::Computers
        $ADPicker.AllowedLocations = [System.DirectoryServices.UI.DirectoryLocations]::All
        $ADPicker.DefaultLocations = [System.DirectoryServices.UI.DirectoryLocations]::JoinedDomain
        $ADPicker.MultiSelect = $true
        
        if ($ADPicker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $SelectedComputers = $ADPicker.SelectedObjects | ForEach-Object { $_.Name }
            $ComputerTextBox.Text = $SelectedComputers -join ","
        }
    })
    
    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = New-Object System.Drawing.Point(670, 20)
    $RefreshButton.Size = New-Object System.Drawing.Size(90, 23)
    $RefreshButton.Text = "Refresh"
    $RefreshButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        if ($TabControl.SelectedTab.Name -eq "ApplicationsTab") {
            $ApplicationsView.Items.Clear()
            foreach ($Computer in $Computers) {
                $Applications = Get-AvailableMecmApplications -ComputerName $Computer
                foreach ($App in $Applications) {
                    $Item = New-Object System.Windows.Forms.ListViewItem
                    $Item.Text = $App.Name
                    $Item.SubItems.Add($App.Publisher)
                    $Item.SubItems.Add($App.SoftwareVersion)
                    if ($App.IsInstalled) { $Items.SubItems.Add("Yes") } else { $Items.SubItems.Add("No") }
                    $Item.SubItems.Add($App.Id)
                    $Item.SubItems.Add($Computer)
                    $ApplicationsView.Items.Add($Item)
                }
            }
        }
        elseif ($TabControl.SelectedTab.Name -eq "PackagesTab") {
            $PackagesView.Items.Clear()
            foreach ($Computer in $Computers) {
                $Packages = Get-AvailableMecmPackages -ComputerName $Computer
                foreach ($Pkg in $Packages) {
                    $Item = New-Object System.Windows.Forms.ListViewItem
                    $Item.Text = $Pkg.Name
                    $Item.SubItems.Add($Pkg.Description)
                    $Item.SubItems.Add($Pkg.PackageID)
                    $Item.SubItems.Add($Pkg.ProgramID)
                    $Item.SubItems.Add($Computer)
                    $PackagesView.Items.Add($Item)
                }
            }
        }
        elseif ($TabControl.SelectedTab.Name -eq "UpdatesTab") {
            $UpdatesView.Items.Clear()
            foreach ($Computer in $Computers) {
                $Updates = Get-AvailableMecmUpdates -ComputerName $Computer
                foreach ($Update in $Updates) {
                    $Item = New-Object System.Windows.Forms.ListViewItem
                    $Item.Text = $Update.Name
                    $Item.SubItems.Add($Update.ArticleID)
                    $Item.SubItems.Add($Update.Description)
                    $Item.SubItems.Add($Update.UpdateID)
                    $Item.SubItems.Add($Computer)
                    $UpdatesView.Items.Add($Item)
                }
            }
        }
        elseif ($TabControl.SelectedTab.Name -eq "TaskSequencesTab") {
            $TaskSequencesView.Items.Clear()
            foreach ($Computer in $Computers) {
                $TaskSequences = Get-AvailableMecmTaskSequences -ComputerName $Computer
                foreach ($TS in $TaskSequences) {
                    $Item = New-Object System.Windows.Forms.ListViewItem
                    $Item.Text = $TS.Name
                    $Item.SubItems.Add($TS.Description)
                    $Item.SubItems.Add($TS.PackageID)
                    $Item.SubItems.Add($Computer)
                    $TaskSequencesView.Items.Add($Item)
                }
            }
        }
        elseif ($TabControl.SelectedTab.Name -eq "ClientActionsTab") {
            # No refresh needed for client actions tab
        }
    })
    
    $StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $StatusLabel.Text = "Ready"
    $SiteInfoLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $SiteInfoLabel.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $SiteInfoLabel.Text = "Site: $SiteCode | Server: $SiteServer"
    
    $StatusStrip.Items.Add($StatusLabel)
    $StatusStrip.Items.Add($SiteInfoLabel)
    
    # Add controls to the top panel
    $TopPanel.Controls.Add($ComputerLabel)
    $TopPanel.Controls.Add($ComputerTextBox)
    $TopPanel.Controls.Add($BrowseButton)
    $TopPanel.Controls.Add($RefreshButton)
    
    # Applications Tab
    $ApplicationsTab = New-Object System.Windows.Forms.TabPage
    $ApplicationsTab.Text = "Applications"
    $ApplicationsTab.Name = "ApplicationsTab"
    
    $ApplicationsView = New-Object System.Windows.Forms.ListView
    $ApplicationsView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ApplicationsView.View = [System.Windows.Forms.View]::Details
    $ApplicationsView.FullRowSelect = $true
    $ApplicationsView.MultiSelect = $true
    
    $ApplicationsView.Columns.Add("Name", 200)
    $ApplicationsView.Columns.Add("Publisher", 100)
    $ApplicationsView.Columns.Add("Version", 100)
    $ApplicationsView.Columns.Add("Installed", 70)
    $ApplicationsView.Columns.Add("ID", 150) | Out-Null
    $ApplicationsView.Columns.Add("Computer", 100) | Out-Null
    
    $ApplicationsContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $InstallApplicationMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $InstallApplicationMenuItem.Text = "Install"
    $InstallApplicationMenuItem.Add_Click({
        $SelectedItems = $ApplicationsView.SelectedItems
        foreach ($Item in $SelectedItems) {
            $ComputerName = $Item.SubItems[5].Text
            $ApplicationID = $Item.SubItems[4].Text
            
            $StatusLabel.Text = "Installing application $($Item.Text) on $ComputerName..."
            
            $Result = Install-MecmApplication -ComputerName $ComputerName -ApplicationID $ApplicationID
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered installation of $($Item.Text) on $ComputerName"
            }
            else {
                $StatusLabel.Text = "Failed to install $($Item.Text) on $ComputerName"
            }
        }
    })
    
    $UninstallApplicationMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $UninstallApplicationMenuItem.Text = "Uninstall"
    $UninstallApplicationMenuItem.Add_Click({
        $SelectedItems = $ApplicationsView.SelectedItems
        foreach ($Item in $SelectedItems) {
            $ComputerName = $Item.SubItems[5].Text
            $ApplicationID = $Item.SubItems[4].Text
            
            $StatusLabel.Text = "Uninstalling application $($Item.Text) on $ComputerName..."
            
            $Result = Uninstall-MecmApplication -ComputerName $ComputerName -ApplicationID $ApplicationID
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered uninstallation of $($Item.Text) on $ComputerName"
            }
            else {
                $StatusLabel.Text = "Failed to uninstall $($Item.Text) on $ComputerName"
            }
        }
    })
    
    $DownloadApplicationMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $DownloadApplicationMenuItem.Text = "Download Content"
    $DownloadApplicationMenuItem.Add_Click({
        $SelectedItems = $ApplicationsView.SelectedItems
        
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select destination folder for content"
        $FolderBrowser.ShowNewFolderButton = $true
        
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $DestinationPath = $FolderBrowser.SelectedPath
            
            foreach ($Item in $SelectedItems) {
                $ApplicationID = $Item.SubItems[4].Text
                
                $StatusLabel.Text = "Downloading content for $($Item.Text)..."
                
                $Result = Get-MecmContent -ContentID $ApplicationID -DestinationPath $DestinationPath -ContentType "Application" -SiteCode $SiteCode -SiteServer $SiteServer
                
                if ($Result) {
                    $StatusLabel.Text = "Successfully downloaded content for $($Item.Text) to $DestinationPath"
                }
                else {
                    $StatusLabel.Text = "Failed to download content for $($Item.Text)"
                }
            }
        }
    })
    
    $ApplicationsContextMenu.Items.Add($InstallApplicationMenuItem)
    $ApplicationsContextMenu.Items.Add($UninstallApplicationMenuItem)
    $ApplicationsContextMenu.Items.Add($DownloadApplicationMenuItem)
    
    $ApplicationsView.ContextMenuStrip = $ApplicationsContextMenu
    $ApplicationsTab.Controls.Add($ApplicationsView)
    
    # Packages Tab
    $PackagesTab = New-Object System.Windows.Forms.TabPage
    $PackagesTab.Text = "Packages"
    $PackagesTab.Name = "PackagesTab"
    
    $PackagesView = New-Object System.Windows.Forms.ListView
    $PackagesView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $PackagesView.View = [System.Windows.Forms.View]::Details
    $PackagesView.FullRowSelect = $true
    $PackagesView.MultiSelect = $true
    
    $PackagesView.Columns.Add("Name", 200)
    $PackagesView.Columns.Add("Description", 200)
    $PackagesView.Columns.Add("PackageID", 100)
    $PackagesView.Columns.Add("ProgramID", 100)
    $PackagesView.Columns.Add("Computer", 100) | Out-Null
    
    $PackagesContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $InstallPackageMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $InstallPackageMenuItem.Text = "Install"
    $InstallPackageMenuItem.Add_Click({
        $SelectedItems = $PackagesView.SelectedItems
        foreach ($Item in $SelectedItems) {
            $ComputerName = $Item.SubItems[4].Text
            $PackageID = $Item.SubItems[2].Text
            $ProgramID = $Item.SubItems[3].Text
            
            $StatusLabel.Text = "Installing package $($Item.Text) on $ComputerName..."
            
            $Result = Install-MecmPackage -ComputerName $ComputerName -PackageID $PackageID -ProgramName $ProgramID
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered installation of $($Item.Text) on $ComputerName"
            }
            else {
                $StatusLabel.Text = "Failed to install $($Item.Text) on $ComputerName"
            }
        }
    })
    
    $DownloadPackageMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $DownloadPackageMenuItem.Text = "Download Content"
    $DownloadPackageMenuItem.Add_Click({
        $SelectedItems = $PackagesView.SelectedItems
        
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select destination folder for content"
        $FolderBrowser.ShowNewFolderButton = $true
        
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $DestinationPath = $FolderBrowser.SelectedPath
            
            foreach ($Item in $SelectedItems) {
                $PackageID = $Item.SubItems[2].Text
                
                $StatusLabel.Text = "Downloading content for $($Item.Text)..."
                
                $Result = Get-MecmContent -ContentID $PackageID -DestinationPath $DestinationPath -ContentType "Package" -SiteCode $SiteCode -SiteServer $SiteServer
                
                if ($Result) {
                    $StatusLabel.Text = "Successfully downloaded content for $($Item.Text) to $DestinationPath"
                }
                else {
                    $StatusLabel.Text = "Failed to download content for $($Item.Text)"
                }
            }
        }
    })
    
    $PackagesContextMenu.Items.Add($InstallPackageMenuItem)
    $PackagesContextMenu.Items.Add($DownloadPackageMenuItem)
    
    $PackagesView.ContextMenuStrip = $PackagesContextMenu
    $PackagesTab.Controls.Add($PackagesView)
    
    # Updates Tab
    $UpdatesTab = New-Object System.Windows.Forms.TabPage
    $UpdatesTab.Text = "Updates"
    $UpdatesTab.Name = "UpdatesTab"
    
    $UpdatesView = New-Object System.Windows.Forms.ListView
    $UpdatesView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $UpdatesView.View = [System.Windows.Forms.View]::Details
    $UpdatesView.FullRowSelect = $true
    $UpdatesView.MultiSelect = $true
    
    $UpdatesView.Columns.Add("Name", 250)
    $UpdatesView.Columns.Add("ArticleID", 100)
    $UpdatesView.Columns.Add("Description", 200)
    $UpdatesView.Columns.Add("UpdateID", 150) | Out-Null
    $UpdatesView.Columns.Add("Computer", 100) | Out-Null
    
    $UpdatesContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $InstallUpdateMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $InstallUpdateMenuItem.Text = "Install"
    $InstallUpdateMenuItem.Add_Click({
        $SelectedItems = $UpdatesView.SelectedItems
        
        # Group by computer
        $ComputerUpdates = @{}
        foreach ($Item in $SelectedItems) {
            $ComputerName = $Item.SubItems[4].Text
            $UpdateID = $Item.SubItems[3].Text
            
            if (-not $ComputerUpdates.ContainsKey($ComputerName)) {
                $ComputerUpdates[$ComputerName] = @()
            }
            
            $ComputerUpdates[$ComputerName] += $UpdateID
        }
        
        # Install updates by computer
        foreach ($Computer in $ComputerUpdates.Keys) {
            $UpdateIDs = $ComputerUpdates[$Computer]
            
            $StatusLabel.Text = "Installing updates on $Computer..."
            
            $Result = Install-MecmUpdate -ComputerName $Computer -UpdateIDs $UpdateIDs
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered installation of updates on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to install updates on $Computer"
            }
        }
    })
    
    $DownloadUpdateMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $DownloadUpdateMenuItem.Text = "Download Content"
    $DownloadUpdateMenuItem.Add_Click({
        $SelectedItems = $UpdatesView.SelectedItems
        
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select destination folder for content"
        $FolderBrowser.ShowNewFolderButton = $true
        
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $DestinationPath = $FolderBrowser.SelectedPath
            
            foreach ($Item in $SelectedItems) {
                $UpdateID = $Item.SubItems[3].Text
                
                $StatusLabel.Text = "Downloading content for $($Item.Text)..."
                
                $Result = Get-MecmContent -ContentID $UpdateID -DestinationPath $DestinationPath -ContentType "Update" -SiteCode $SiteCode -SiteServer $SiteServer
                
                if ($Result) {
                    $StatusLabel.Text = "Successfully downloaded content for $($Item.Text) to $DestinationPath"
                }
                else {
                    $StatusLabel.Text = "Failed to download content for $($Item.Text)"
                }
            }
        }
    })
    
    $UpdatesContextMenu.Items.Add($InstallUpdateMenuItem)
    $UpdatesContextMenu.Items.Add($DownloadUpdateMenuItem)
    
    $UpdatesView.ContextMenuStrip = $UpdatesContextMenu
    $UpdatesTab.Controls.Add($UpdatesView)
    
    # Task Sequences Tab
    $TaskSequencesTab = New-Object System.Windows.Forms.TabPage
    $TaskSequencesTab.Text = "Task Sequences"
    $TaskSequencesTab.Name = "TaskSequencesTab"
    
    $TaskSequencesView = New-Object System.Windows.Forms.ListView
    $TaskSequencesView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $TaskSequencesView.View = [System.Windows.Forms.View]::Details
    $TaskSequencesView.FullRowSelect = $true
    $TaskSequencesView.MultiSelect = $true
    
    $TaskSequencesView.Columns.Add("Name", 200)
    $TaskSequencesView.Columns.Add("Description", 300)
    $TaskSequencesView.Columns.Add("PackageID", 100) | Out-Null
    $TaskSequencesView.Columns.Add("Computer", 100) | Out-Null
    
    $TaskSequencesContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $InstallTaskSequenceMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $InstallTaskSequenceMenuItem.Text = "Run"
    $InstallTaskSequenceMenuItem.Add_Click({
        $SelectedItems = $TaskSequencesView.SelectedItems
        foreach ($Item in $SelectedItems) {
            $ComputerName = $Item.SubItems[3].Text
            $PackageID = $Item.SubItems[2].Text
            
            $StatusLabel.Text = "Running task sequence $($Item.Text) on $ComputerName..."
            
            $Result = Install-MecmTaskSequence -ComputerName $ComputerName -PackageID $PackageID
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered task sequence $($Item.Text) on $ComputerName"
            }
            else {
                $StatusLabel.Text = "Failed to run task sequence $($Item.Text) on $ComputerName"
            }
        }
    })
    
    $DownloadTaskSequenceMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $DownloadTaskSequenceMenuItem.Text = "Download Content"
    $DownloadTaskSequenceMenuItem.Add_Click({
        $SelectedItems = $TaskSequencesView.SelectedItems
        
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select destination folder for content"
        $FolderBrowser.ShowNewFolderButton = $true
        
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $DestinationPath = $FolderBrowser.SelectedPath
            
            foreach ($Item in $SelectedItems) {
                $PackageID = $Item.SubItems[2].Text
                
                $StatusLabel.Text = "Downloading content for $($Item.Text)..."
                
                $Result = Get-MecmContent -ContentID $PackageID -DestinationPath $DestinationPath -ContentType "TaskSequence" -SiteCode $SiteCode -SiteServer $SiteServer
                
                if ($Result) {
                    $StatusLabel.Text = "Successfully downloaded content for $($Item.Text) to $DestinationPath"
                }
                else {
                    $StatusLabel.Text = "Failed to download content for $($Item.Text)"
                }
            }
        }
    })
    
    $TaskSequencesContextMenu.Items.Add($InstallTaskSequenceMenuItem)
    $TaskSequencesContextMenu.Items.Add($DownloadTaskSequenceMenuItem)
    
    $TaskSequencesView.ContextMenuStrip = $TaskSequencesContextMenu
    $TaskSequencesTab.Controls.Add($TaskSequencesView)
    
    # Client Actions Tab
    $ClientActionsTab = New-Object System.Windows.Forms.TabPage
    $ClientActionsTab.Text = "Client Actions"
    $ClientActionsTab.Name = "ClientActionsTab"
    
    $ClientActionsPanel = New-Object System.Windows.Forms.Panel
    $ClientActionsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ClientActionsPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $MachinePolLabel = New-Object System.Windows.Forms.Label
    $MachinePolLabel.Text = "Machine Policy Retrieval and Evaluation Cycle"
    $MachinePolLabel.Location = New-Object System.Drawing.Point(20, 20)
    $MachinePolLabel.Size = New-Object System.Drawing.Size(300, 20)
    
    $MachinePolButton = New-Object System.Windows.Forms.Button
    $MachinePolButton.Text = "Run"
    $MachinePolButton.Location = New-Object System.Drawing.Point(400, 20)
    $MachinePolButton.Size = New-Object System.Drawing.Size(100, 23)
    $MachinePolButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        foreach ($Computer in $Computers) {
            $StatusLabel.Text = "Triggering Machine Policy Retrieval on $Computer..."
            
            $Result = Trigger-MecmClientAction -ComputerName $Computer -ActionType "MachinePolicy"
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered Machine Policy Retrieval on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to trigger Machine Policy Retrieval on $Computer"
            }
        }
    })
    
    $DiscoveryLabel = New-Object System.Windows.Forms.Label
    $DiscoveryLabel.Text = "Discovery Data Collection Cycle"
    $DiscoveryLabel.Location = New-Object System.Drawing.Point(20, 60)
    $DiscoveryLabel.Size = New-Object System.Drawing.Size(300, 20)
    
    $DiscoveryButton = New-Object System.Windows.Forms.Button
    $DiscoveryButton.Text = "Run"
    $DiscoveryButton.Location = New-Object System.Drawing.Point(400, 60)
    $DiscoveryButton.Size = New-Object System.Drawing.Size(100, 23)
    $DiscoveryButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        foreach ($Computer in $Computers) {
            $StatusLabel.Text = "Triggering Discovery Data Collection on $Computer..."
            
            $Result = Trigger-MecmClientAction -ComputerName $Computer -ActionType "DiscoveryData"
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered Discovery Data Collection on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to trigger Discovery Data Collection on $Computer"
            }
        }
    })
    
    $SWInvLabel = New-Object System.Windows.Forms.Label
    $SWInvLabel.Text = "Software Inventory Cycle"
    $SWInvLabel.Location = New-Object System.Drawing.Point(20, 100)
    $SWInvLabel.Size = New-Object System.Drawing.Size(300, 20)
    
    $SWInvButton = New-Object System.Windows.Forms.Button
    $SWInvButton.Text = "Run"
    $SWInvButton.Location = New-Object System.Drawing.Point(400, 100)
    $SWInvButton.Size = New-Object System.Drawing.Size(100, 23)
    $SWInvButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        foreach ($Computer in $Computers) {
            $StatusLabel.Text = "Triggering Software Inventory on $Computer..."
            
            $Result = Trigger-MecmClientAction -ComputerName $Computer -ActionType "SoftwareInventory"
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered Software Inventory on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to trigger Software Inventory on $Computer"
            }
        }
    })
    
    $HWInvLabel = New-Object System.Windows.Forms.Label
    $HWInvLabel.Text = "Hardware Inventory Cycle"
    $HWInvLabel.Location = New-Object System.Drawing.Point(20, 140)
    $HWInvLabel.Size = New-Object System.Drawing.Size(300, 20)
    
    $HWInvButton = New-Object System.Windows.Forms.Button
    $HWInvButton.Text = "Run"
    $HWInvButton.Location = New-Object System.Drawing.Point(400, 140)
    $HWInvButton.Size = New-Object System.Drawing.Size(100, 23)
    $HWInvButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        foreach ($Computer in $Computers) {
            $StatusLabel.Text = "Triggering Hardware Inventory on $Computer..."
            
            $Result = Trigger-MecmClientAction -ComputerName $Computer -ActionType "HardwareInventory"
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered Hardware Inventory on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to trigger Hardware Inventory on $Computer"
            }
        }
    })
    
    $UpdateScanLabel = New-Object System.Windows.Forms.Label
    $UpdateScanLabel.Text = "Software Update Scan Cycle"
    $UpdateScanLabel.Location = New-Object System.Drawing.Point(20, 180)
    $UpdateScanLabel.Size = New-Object System.Drawing.Size(300, 20)
    
    $UpdateScanButton = New-Object System.Windows.Forms.Button
    $UpdateScanButton.Text = "Run"
    $UpdateScanButton.Location = New-Object System.Drawing.Point(400, 180)
    $UpdateScanButton.Size = New-Object System.Drawing.Size(100, 23)
    $UpdateScanButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        foreach ($Computer in $Computers) {
            $StatusLabel.Text = "Triggering Software Update Scan on $Computer..."
            
            $Result = Trigger-MecmClientAction -ComputerName $Computer -ActionType "UpdateScan"
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered Software Update Scan on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to trigger Software Update Scan on $Computer"
            }
        }
    })
    
    $UpdateDeployLabel = New-Object System.Windows.Forms.Label
    $UpdateDeployLabel.Text = "Software Update Deployment Evaluation Cycle"
    $UpdateDeployLabel.Location = New-Object System.Drawing.Point(20, 220)
    $UpdateDeployLabel.Size = New-Object System.Drawing.Size(300, 20)
    
    $UpdateDeployButton = New-Object System.Windows.Forms.Button
    $UpdateDeployButton.Text = "Run"
    $UpdateDeployButton.Location = New-Object System.Drawing.Point(400, 220)
    $UpdateDeployButton.Size = New-Object System.Drawing.Size(100, 23)
    $UpdateDeployButton.Add_Click({
        $Computers = $ComputerTextBox.Text -split "," | ForEach-Object { $_.Trim() }
        
        foreach ($Computer in $Computers) {
            $StatusLabel.Text = "Triggering Software Update Deployment Evaluation on $Computer..."
            
            $Result = Trigger-MecmClientAction -ComputerName $Computer -ActionType "UpdateDeployment"
            
            if ($Result) {
                $StatusLabel.Text = "Successfully triggered Software Update Deployment Evaluation on $Computer"
            }
            else {
                $StatusLabel.Text = "Failed to trigger Software Update Deployment Evaluation on $Computer"
            }
        }
    })
    
    # Add controls to the client actions panel
    $ClientActionsPanel.Controls.Add($MachinePolLabel)
    $ClientActionsPanel.Controls.Add($MachinePolButton)
    $ClientActionsPanel.Controls.Add($DiscoveryLabel)
    $ClientActionsPanel.Controls.Add($DiscoveryButton)
    $ClientActionsPanel.Controls.Add($SWInvLabel)
    $ClientActionsPanel.Controls.Add($SWInvButton)
    $ClientActionsPanel.Controls.Add($HWInvLabel)
    $ClientActionsPanel.Controls.Add($HWInvButton)
    $ClientActionsPanel.Controls.Add($UpdateScanLabel)
    $ClientActionsPanel.Controls.Add($UpdateScanButton)
    $ClientActionsPanel.Controls.Add($UpdateDeployLabel)
    $ClientActionsPanel.Controls.Add($UpdateDeployButton)
    
    $ClientActionsTab.Controls.Add($ClientActionsPanel)
    
    # Add tabs to the tab control
    $TabControl.TabPages.Add($ApplicationsTab)
    $TabControl.TabPages.Add($PackagesTab)
    $TabControl.TabPages.Add($UpdatesTab)
    $TabControl.TabPages.Add($TaskSequencesTab)
    $TabControl.TabPages.Add($ClientActionsTab)
    
    # Add controls to the main form
    $MainForm.Controls.Add($TopPanel)
    $MainForm.Controls.Add($TabControl)
    $MainForm.Controls.Add($StatusStrip)
    
    # Tab selection change event
    $TabControl.Add_SelectedIndexChanged({
        # Call the refresh button click event to populate the current tab
        $RefreshButton.PerformClick()
    })
    
    # Show the form
    $MainForm.ShowDialog() | Out-Null
}

#endregion

# Main script execution
Show-MainForm
