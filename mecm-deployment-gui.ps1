#Requires -Modules ConfigurationManager
<#
.SYNOPSIS
    GUI-based MECM Deployment Automation Tool for managing deployments and collections.
    
.DESCRIPTION
    This script provides a graphical user interface for common MECM tasks including:
    - Deploying applications, packages, software updates, and task sequences to collections
    - Adding and removing computers from collections
    - Renaming collections
    
.NOTES
    Filename: MECM-Deployment-GUI.ps1
    Author: NonExistent
    Requirements: 
        - PowerShell 5.1 or higher
        - ConfigurationManager PowerShell module
        - Administrative rights in MECM
        - Site connection must be established before running script
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$global:SiteCode = ""
$global:SiteServer = ""
$global:Connected = $false

# Main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "MECM Deployment Automation Tool"
$mainForm.Size = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false

# Status strip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Not connected to MECM"
$statusStrip.Items.Add($statusLabel)
$mainForm.Controls.Add($statusStrip)

# Function to update status
function Update-Status {
    param (
        [string]$Text,
        [string]$Color = "Black"
    )
    
    $statusLabel.Text = $Text
    $statusLabel.ForeColor = $Color
}

# Function to show output
function Write-Output {
    param (
        [string]$Text,
        [string]$Color = "Black"
    )
    
    $outputTextBox.SelectionColor = $Color
    $outputTextBox.AppendText("$Text`r`n")
    $outputTextBox.ScrollToCaret()
}

# Function to connect to MECM
function Connect-MECM {
    param (
        [string]$SiteCode,
        [string]$SiteServer
    )
    
    try {
        # Import the ConfigurationManager.psd1 module
        if (-not (Get-Module ConfigurationManager)) {
            Import-Module ConfigurationManager -ErrorAction Stop
        }
        
        # Set the current location to the site code PSDrive
        $PSDrivePath = "$SiteCode`:"
        if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SiteServer -Description "MECM Site" -ErrorAction Stop
        }
        Set-Location $PSDrivePath -ErrorAction Stop
        
        Write-Output -Text "Successfully connected to MECM Site $SiteCode on server $SiteServer" -Color "Green"
        Update-Status -Text "Connected to MECM Site $SiteCode" -Color "Green"
        $global:Connected = $true
        $global:SiteCode = $SiteCode
        $global:SiteServer = $SiteServer
        
        # Enable tabs
        $tabControl.Enabled = $true
        
        return $true
    }
    catch {
        Write-Output -Text "Failed to connect to MECM: $_" -Color "Red"
        Update-Status -Text "Connection failed" -Color "Red"
        return $false
    }
}

# Connection panel
$connectionPanel = New-Object System.Windows.Forms.Panel
$connectionPanel.Location = New-Object System.Drawing.Point(10, 10)
$connectionPanel.Size = New-Object System.Drawing.Size(765, 100)
$connectionPanel.BorderStyle = "FixedSingle"

$connectionLabel = New-Object System.Windows.Forms.Label
$connectionLabel.Text = "Connect to MECM"
$connectionLabel.Location = New-Object System.Drawing.Point(10, 10)
$connectionLabel.Size = New-Object System.Drawing.Size(200, 20)
$connectionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$connectionPanel.Controls.Add($connectionLabel)

$siteCodeLabel = New-Object System.Windows.Forms.Label
$siteCodeLabel.Text = "Site Code:"
$siteCodeLabel.Location = New-Object System.Drawing.Point(10, 40)
$siteCodeLabel.Size = New-Object System.Drawing.Size(100, 20)
$connectionPanel.Controls.Add($siteCodeLabel)

$siteCodeTextBox = New-Object System.Windows.Forms.TextBox
$siteCodeTextBox.Location = New-Object System.Drawing.Point(120, 40)
$siteCodeTextBox.Size = New-Object System.Drawing.Size(150, 20)
$connectionPanel.Controls.Add($siteCodeTextBox)

$siteServerLabel = New-Object System.Windows.Forms.Label
$siteServerLabel.Text = "Site Server:"
$siteServerLabel.Location = New-Object System.Drawing.Point(10, 70)
$siteServerLabel.Size = New-Object System.Drawing.Size(100, 20)
$connectionPanel.Controls.Add($siteServerLabel)

$siteServerTextBox = New-Object System.Windows.Forms.TextBox
$siteServerTextBox.Location = New-Object System.Drawing.Point(120, 70)
$siteServerTextBox.Size = New-Object System.Drawing.Size(300, 20)
$connectionPanel.Controls.Add($siteServerTextBox)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect"
$connectButton.Location = New-Object System.Drawing.Point(450, 40)
$connectButton.Size = New-Object System.Drawing.Size(100, 30)
$connectButton.Add_Click({
    $siteCode = $siteCodeTextBox.Text.Trim()
    $siteServer = $siteServerTextBox.Text.Trim()
    
    if (-not $siteCode -or -not $siteServer) {
        [System.Windows.Forms.MessageBox]::Show("Please enter Site Code and Site Server", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    Connect-MECM -SiteCode $siteCode -SiteServer $siteServer
})
$connectionPanel.Controls.Add($connectButton)

$mainForm.Controls.Add($connectionPanel)

# Output panel
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Output:"
$outputLabel.Location = New-Object System.Drawing.Point(10, 115)
$outputLabel.Size = New-Object System.Drawing.Size(100, 20)
$outputLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.RichTextBox
$outputTextBox.Location = New-Object System.Drawing.Point(10, 140)
$outputTextBox.Size = New-Object System.Drawing.Size(765, 150)
$outputTextBox.ReadOnly = $true
$outputTextBox.BackColor = [System.Drawing.Color]::White
$outputTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$mainForm.Controls.Add($outputTextBox)

# Tab control for operations
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 300)
$tabControl.Size = New-Object System.Drawing.Size(765, 240)
$tabControl.Enabled = $false
$mainForm.Controls.Add($tabControl)

# Application Deployment Tab
$tabApplication = New-Object System.Windows.Forms.TabPage
$tabApplication.Text = "Application Deployment"

$appNameLabel = New-Object System.Windows.Forms.Label
$appNameLabel.Text = "Application Name:"
$appNameLabel.Location = New-Object System.Drawing.Point(10, 20)
$appNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabApplication.Controls.Add($appNameLabel)

$appNameComboBox = New-Object System.Windows.Forms.ComboBox
$appNameComboBox.Location = New-Object System.Drawing.Point(150, 20)
$appNameComboBox.Size = New-Object System.Drawing.Size(300, 20)
$appNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$appNameComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$appNameComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabApplication.Controls.Add($appNameComboBox)

$appRefreshButton = New-Object System.Windows.Forms.Button
$appRefreshButton.Text = "Refresh"
$appRefreshButton.Location = New-Object System.Drawing.Point(460, 20)
$appRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$appRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving applications..." -Color "Blue"
        $appNameComboBox.Items.Clear()
        Get-CMApplication | ForEach-Object {
            $appNameComboBox.Items.Add($_.LocalizedDisplayName)
        }
        Write-Output -Text "Applications retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve applications: $_" -Color "Red"
    }
})
$tabApplication.Controls.Add($appRefreshButton)

$appCollectionLabel = New-Object System.Windows.Forms.Label
$appCollectionLabel.Text = "Collection Name:"
$appCollectionLabel.Location = New-Object System.Drawing.Point(10, 50)
$appCollectionLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabApplication.Controls.Add($appCollectionLabel)

$appCollectionComboBox = New-Object System.Windows.Forms.ComboBox
$appCollectionComboBox.Location = New-Object System.Drawing.Point(150, 50)
$appCollectionComboBox.Size = New-Object System.Drawing.Size(300, 20)
$appCollectionComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$appCollectionComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$appCollectionComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabApplication.Controls.Add($appCollectionComboBox)

$appCollRefreshButton = New-Object System.Windows.Forms.Button
$appCollRefreshButton.Text = "Refresh"
$appCollRefreshButton.Location = New-Object System.Drawing.Point(460, 50)
$appCollRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$appCollRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving device collections..." -Color "Blue"
        $appCollectionComboBox.Items.Clear()
        Get-CMDeviceCollection | ForEach-Object {
            $appCollectionComboBox.Items.Add($_.Name)
        }
        Write-Output -Text "Device collections retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve collections: $_" -Color "Red"
    }
})
$tabApplication.Controls.Add($appCollRefreshButton)

$appPurposeLabel = New-Object System.Windows.Forms.Label
$appPurposeLabel.Text = "Deployment Purpose:"
$appPurposeLabel.Location = New-Object System.Drawing.Point(10, 80)
$appPurposeLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabApplication.Controls.Add($appPurposeLabel)

$appPurposeComboBox = New-Object System.Windows.Forms.ComboBox
$appPurposeComboBox.Location = New-Object System.Drawing.Point(150, 80)
$appPurposeComboBox.Size = New-Object System.Drawing.Size(150, 20)
$appPurposeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$appPurposeComboBox.Items.Add("Available")
$appPurposeComboBox.Items.Add("Required")
$appPurposeComboBox.SelectedIndex = 0
$tabApplication.Controls.Add($appPurposeComboBox)

$appDeadlineCheck = New-Object System.Windows.Forms.CheckBox
$appDeadlineCheck.Text = "Set Deadline"
$appDeadlineCheck.Location = New-Object System.Drawing.Point(10, 110)
$appDeadlineCheck.Size = New-Object System.Drawing.Size(100, 20)
$tabApplication.Controls.Add($appDeadlineCheck)

$appDeadlinePicker = New-Object System.Windows.Forms.DateTimePicker
$appDeadlinePicker.Location = New-Object System.Drawing.Point(150, 110)
$appDeadlinePicker.Size = New-Object System.Drawing.Size(200, 20)
$appDeadlinePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$appDeadlinePicker.CustomFormat = "yyyy-MM-dd HH:mm"
$appDeadlinePicker.Value = (Get-Date).AddDays(7)
$appDeadlinePicker.Enabled = $false
$tabApplication.Controls.Add($appDeadlinePicker)

$appOverrideCheck = New-Object System.Windows.Forms.CheckBox
$appOverrideCheck.Text = "Override Maintenance Window"
$appOverrideCheck.Location = New-Object System.Drawing.Point(360, 110)
$appOverrideCheck.Size = New-Object System.Drawing.Size(200, 20)
$appOverrideCheck.Enabled = $false
$tabApplication.Controls.Add($appOverrideCheck)

$appDeadlineCheck.Add_CheckedChanged({
    $appDeadlinePicker.Enabled = $appDeadlineCheck.Checked
    $appOverrideCheck.Enabled = $appDeadlineCheck.Checked
})

$appDeployButton = New-Object System.Windows.Forms.Button
$appDeployButton.Text = "Deploy Application"
$appDeployButton.Location = New-Object System.Drawing.Point(150, 150)
$appDeployButton.Size = New-Object System.Drawing.Size(150, 30)
$appDeployButton.Add_Click({
    try {
        $appName = $appNameComboBox.Text.Trim()
        $collName = $appCollectionComboBox.Text.Trim()
        $purpose = $appPurposeComboBox.SelectedItem.ToString()
        
        if (-not $appName -or -not $collName) {
            [System.Windows.Forms.MessageBox]::Show("Please enter Application Name and Collection Name", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        Write-Output -Text "Deploying application '$appName' to collection '$collName'..." -Color "Blue"
        
        # Verify application exists
        $application = Get-CMApplication -Name $appName -ErrorAction Stop
        if (-not $application) {
            Write-Output -Text "Application '$appName' not found" -Color "Red"
            return
        }
        
        # Verify collection exists
        $collection = Get-CMCollection -Name $collName -ErrorAction Stop
        if (-not $collection) {
            Write-Output -Text "Collection '$collName' not found" -Color "Red"
            return
        }
        
        # Set deployment parameters
        $deploymentParams = @{
            Application = $appName
            Collection = $collName
            DeployPurpose = $purpose
            DeployAction = "Install"
            UserNotification = "DisplayAll"
            AllowRepairApp = $true
            TimeBaseOn = "LocalTime"
            AvailableDateTime = (Get-Date)
        }
        
        # Add deadline if provided
        if ($appDeadlineCheck.Checked) {
            $deploymentParams.Add("DeadlineDateTime", $appDeadlinePicker.Value)
            $deploymentParams.Add("OverrideServiceWindow", $appOverrideCheck.Checked)
        }
        
        # Create deployment
        $deployment = New-CMApplicationDeployment @deploymentParams -ErrorAction Stop
        
        Write-Output -Text "Successfully deployed application '$appName' to collection '$collName'" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to create application deployment: $_" -Color "Red"
    }
})
$tabApplication.Controls.Add($appDeployButton)

$tabControl.Controls.Add($tabApplication)

# Package Deployment Tab
$tabPackage = New-Object System.Windows.Forms.TabPage
$tabPackage.Text = "Package Deployment"

$pkgNameLabel = New-Object System.Windows.Forms.Label
$pkgNameLabel.Text = "Package Name:"
$pkgNameLabel.Location = New-Object System.Drawing.Point(10, 20)
$pkgNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabPackage.Controls.Add($pkgNameLabel)

$pkgNameComboBox = New-Object System.Windows.Forms.ComboBox
$pkgNameComboBox.Location = New-Object System.Drawing.Point(150, 20)
$pkgNameComboBox.Size = New-Object System.Drawing.Size(300, 20)
$pkgNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$pkgNameComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$pkgNameComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabPackage.Controls.Add($pkgNameComboBox)

$pkgRefreshButton = New-Object System.Windows.Forms.Button
$pkgRefreshButton.Text = "Refresh"
$pkgRefreshButton.Location = New-Object System.Drawing.Point(460, 20)
$pkgRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$pkgRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving packages..." -Color "Blue"
        $pkgNameComboBox.Items.Clear()
        Get-CMPackage | ForEach-Object {
            $pkgNameComboBox.Items.Add($_.Name)
        }
        Write-Output -Text "Packages retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve packages: $_" -Color "Red"
    }
})
$tabPackage.Controls.Add($pkgRefreshButton)

$progNameLabel = New-Object System.Windows.Forms.Label
$progNameLabel.Text = "Program Name:"
$progNameLabel.Location = New-Object System.Drawing.Point(10, 50)
$progNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabPackage.Controls.Add($progNameLabel)

$progNameComboBox = New-Object System.Windows.Forms.ComboBox
$progNameComboBox.Location = New-Object System.Drawing.Point(150, 50)
$progNameComboBox.Size = New-Object System.Drawing.Size(300, 20)
$progNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$progNameComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$progNameComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabPackage.Controls.Add($progNameComboBox)

$pkgNameComboBox.Add_SelectedIndexChanged({
    try {
        $pkgName = $pkgNameComboBox.SelectedItem.ToString()
        if ($pkgName) {
            Write-Output -Text "Retrieving programs for package '$pkgName'..." -Color "Blue"
            $progNameComboBox.Items.Clear()
            Get-CMProgram -PackageName $pkgName | ForEach-Object {
                $progNameComboBox.Items.Add($_.ProgramName)
            }
            Write-Output -Text "Programs retrieved successfully" -Color "Green"
        }
    }
    catch {
        Write-Output -Text "Failed to retrieve programs: $_" -Color "Red"
    }
})

$pkgCollectionLabel = New-Object System.Windows.Forms.Label
$pkgCollectionLabel.Text = "Collection Name:"
$pkgCollectionLabel.Location = New-Object System.Drawing.Point(10, 80)
$pkgCollectionLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabPackage.Controls.Add($pkgCollectionLabel)

$pkgCollectionComboBox = New-Object System.Windows.Forms.ComboBox
$pkgCollectionComboBox.Location = New-Object System.Drawing.Point(150, 80)
$pkgCollectionComboBox.Size = New-Object System.Drawing.Size(300, 20)
$pkgCollectionComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$pkgCollectionComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$pkgCollectionComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabPackage.Controls.Add($pkgCollectionComboBox)

$pkgCollRefreshButton = New-Object System.Windows.Forms.Button
$pkgCollRefreshButton.Text = "Refresh"
$pkgCollRefreshButton.Location = New-Object System.Drawing.Point(460, 80)
$pkgCollRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$pkgCollRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving device collections..." -Color "Blue"
        $pkgCollectionComboBox.Items.Clear()
        Get-CMDeviceCollection | ForEach-Object {
            $pkgCollectionComboBox.Items.Add($_.Name)
        }
        Write-Output -Text "Device collections retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve collections: $_" -Color "Red"
    }
})
$tabPackage.Controls.Add($pkgCollRefreshButton)

$pkgPurposeLabel = New-Object System.Windows.Forms.Label
$pkgPurposeLabel.Text = "Deployment Purpose:"
$pkgPurposeLabel.Location = New-Object System.Drawing.Point(10, 110)
$pkgPurposeLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabPackage.Controls.Add($pkgPurposeLabel)

$pkgPurposeComboBox = New-Object System.Windows.Forms.ComboBox
$pkgPurposeComboBox.Location = New-Object System.Drawing.Point(150, 110)
$pkgPurposeComboBox.Size = New-Object System.Drawing.Size(150, 20)
$pkgPurposeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$pkgPurposeComboBox.Items.Add("Optional")
$pkgPurposeComboBox.Items.Add("Required")
$pkgPurposeComboBox.SelectedIndex = 0
$tabPackage.Controls.Add($pkgPurposeComboBox)

$pkgDeadlineCheck = New-Object System.Windows.Forms.CheckBox
$pkgDeadlineCheck.Text = "Set Deadline"
$pkgDeadlineCheck.Location = New-Object System.Drawing.Point(10, 140)
$pkgDeadlineCheck.Size = New-Object System.Drawing.Size(100, 20)
$tabPackage.Controls.Add($pkgDeadlineCheck)

$pkgDeadlinePicker = New-Object System.Windows.Forms.DateTimePicker
$pkgDeadlinePicker.Location = New-Object System.Drawing.Point(150, 140)
$pkgDeadlinePicker.Size = New-Object System.Drawing.Size(200, 20)
$pkgDeadlinePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$pkgDeadlinePicker.CustomFormat = "yyyy-MM-dd HH:mm"
$pkgDeadlinePicker.Value = (Get-Date).AddDays(7)
$pkgDeadlinePicker.Enabled = $false
$tabPackage.Controls.Add($pkgDeadlinePicker)

$pkgOverrideCheck = New-Object System.Windows.Forms.CheckBox
$pkgOverrideCheck.Text = "Override Maintenance Window"
$pkgOverrideCheck.Location = New-Object System.Drawing.Point(360, 140)
$pkgOverrideCheck.Size = New-Object System.Drawing.Size(200, 20)
$pkgOverrideCheck.Enabled = $false
$tabPackage.Controls.Add($pkgOverrideCheck)

$pkgDeadlineCheck.Add_CheckedChanged({
    $pkgDeadlinePicker.Enabled = $pkgDeadlineCheck.Checked
    $pkgOverrideCheck.Enabled = $pkgDeadlineCheck.Checked
})

$pkgDeployButton = New-Object System.Windows.Forms.Button
$pkgDeployButton.Text = "Deploy Package"
$pkgDeployButton.Location = New-Object System.Drawing.Point(150, 170)
$pkgDeployButton.Size = New-Object System.Drawing.Size(150, 30)
$pkgDeployButton.Add_Click({
    try {
        $pkgName = $pkgNameComboBox.Text.Trim()
        $progName = $progNameComboBox.Text.Trim()
        $collName = $pkgCollectionComboBox.Text.Trim()
        $purpose = $pkgPurposeComboBox.SelectedItem.ToString()
        
        if (-not $pkgName -or -not $progName -or -not $collName) {
            [System.Windows.Forms.MessageBox]::Show("Please enter Package Name, Program Name, and Collection Name", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        Write-Output -Text "Deploying package '$pkgName' with program '$progName' to collection '$collName'..." -Color "Blue"
        
        # Verify package exists
        $package = Get-CMPackage -Name $pkgName -ErrorAction Stop
        if (-not $package) {
            Write-Output -Text "Package '$pkgName' not found" -Color "Red"
            return
        }
        
        # Verify program exists in package
        $program = Get-CMProgram -PackageName $pkgName -ProgramName $progName -ErrorAction Stop
        if (-not $program) {
            Write-Output -Text "Program '$progName' not found in package '$pkgName'" -Color "Red"
            return
        }
        
        # Verify collection exists
        $collection = Get-CMCollection -Name $collName -ErrorAction Stop
        if (-not $collection) {
            Write-Output -Text "Collection '$collName' not found" -Color "Red"
            return
        }
        
        # Set deployment parameters
        $deploymentParams = @{
            Package = $pkgName
            Program = $progName
            Collection = $collName
            DeployPurpose = $purpose
            StandardProgram = $true
            FastNetworkOption = "DownloadContentFromDistributionPointAndRunLocally"
            SlowNetworkOption = "DownloadContentFromDistributionPointAndLocally"
            UserNotification = "DisplayAll"
            TimeBaseOn = "LocalTime"
            AvailableDateTime = (Get-Date)
        }
        
        # Add deadline if provided
        if ($pkgDeadlineCheck.Checked) {
            $deploymentParams.Add("DeadlineDateTime", $pkgDeadlinePicker.Value)
            $deploymentParams.Add("OverrideServiceWindow", $pkgOverrideCheck.Checked)
        }
        
        # Create deployment
        $deployment = New-CMPackageDeployment @deploymentParams -ErrorAction Stop
        
        Write-Output -Text "Successfully deployed package '$pkgName' with program '$progName' to collection '$collName'" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to create package deployment: $_" -Color "Red"
    }
})
$tabPackage.Controls.Add($pkgDeployButton)

$tabControl.Controls.Add($tabPackage)

# Software Update Deployment Tab
$tabSoftwareUpdate = New-Object System.Windows.Forms.TabPage
$tabSoftwareUpdate.Text = "Software Update Deployment"

$sugNameLabel = New-Object System.Windows.Forms.Label
$sugNameLabel.Text = "Update Group Name:"
$sugNameLabel.Location = New-Object System.Drawing.Point(10, 20)
$sugNameLabel.Size = New-Object System.Drawing.Size(130, 20)
$tabSoftwareUpdate.Controls.Add($sugNameLabel)

$sugNameComboBox = New-Object System.Windows.Forms.ComboBox
$sugNameComboBox.Location = New-Object System.Drawing.Point(150, 20)
$sugNameComboBox.Size = New-Object System.Drawing.Size(300, 20)
$sugNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$sugNameComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$sugNameComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabSoftwareUpdate.Controls.Add($sugNameComboBox)

$sugRefreshButton = New-Object System.Windows.Forms.Button
$sugRefreshButton.Text = "Refresh"
$sugRefreshButton.Location = New-Object System.Drawing.Point(460, 20)
$sugRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$sugRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving software update groups..." -Color "Blue"
        $sugNameComboBox.Items.Clear()
        Get-CMSoftwareUpdateGroup | ForEach-Object {
            $sugNameComboBox.Items.Add($_.LocalizedDisplayName)
        }
        Write-Output -Text "Software update groups retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve software update groups: $_" -Color "Red"
    }
})
$tabSoftwareUpdate.Controls.Add($sugRefreshButton)

$suCollectionLabel = New-Object System.Windows.Forms.Label
$suCollectionLabel.Text = "Collection Name:"
$suCollectionLabel.Location = New-Object System.Drawing.Point(10, 50)
$suCollectionLabel.Size = New-Object System.Drawing.Size(130, 20)
$tabSoftwareUpdate.Controls.Add($suCollectionLabel)

$suCollectionComboBox = New-Object System.Windows.Forms.ComboBox
$suCollectionComboBox.Location = New-Object System.Drawing.Point(150, 50)
$suCollectionComboBox.Size = New-Object System.Drawing.Size(300, 20)
$suCollectionComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$suCollectionComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$suCollectionComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabSoftwareUpdate.Controls.Add($suCollectionComboBox)

$suCollRefreshButton = New-Object System.Windows.Forms.Button
$suCollRefreshButton.Text = "Refresh"
$suCollRefreshButton.Location = New-Object System.Drawing.Point(460, 50)
$suCollRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$suCollRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving device collections..." -Color "Blue"
        $suCollectionComboBox.Items.Clear()
        Get-CMDeviceCollection | ForEach-Object {
            $suCollectionComboBox.Items.Add($_.Name)
        }
        Write-Output -Text "Device collections retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve collections: $_" -Color "Red"
    }
})
$tabSoftwareUpdate.Controls.Add($suCollRefreshButton)

$suDownloadCheck = New-Object System.Windows.Forms.CheckBox
$suDownloadCheck.Text = "Download content locally when needed by client"
$suDownloadCheck.Location = New-Object System.Drawing.Point(10, 80)
$suDownloadCheck.Size = New-Object System.Drawing.Size(300, 20)
$suDownloadCheck.Checked = $true
$tabSoftwareUpdate.Controls.Add($suDownloadCheck)

$suDeadlineLabel = New-Object System.Windows.Forms.Label
$suDeadlineLabel.Text = "Deadline:"
$suDeadlineLabel.Location = New-Object System.Drawing.Point(10, 110)
$suDeadlineLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabSoftwareUpdate.Controls.Add($suDeadlineLabel)

$suDeadlinePicker = New-Object System.Windows.Forms.DateTimePicker
$suDeadlinePicker.Location = New-Object System.Drawing.Point(150, 110)
$suDeadlinePicker.Size = New-Object System.Drawing.Size(200, 20)
$suDeadlinePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$suDeadlinePicker.CustomFormat = "yyyy-MM-dd HH:mm"
$suDeadlinePicker.Value = (Get-Date).AddDays(7)
$tabSoftwareUpdate.Controls.Add($suDeadlinePicker)

$suOverrideCheck = New-Object System.Windows.Forms.CheckBox
$suOverrideCheck.Text = "Override Maintenance Window"
$suOverrideCheck.Location = New-Object System.Drawing.Point(360, 110)
$suOverrideCheck.Size = New-Object System.Drawing.Size(200, 20)
$tabSoftwareUpdate.Controls.Add($suOverrideCheck)

$suDeployNameLabel = New-Object System.Windows.Forms.Label
$suDeployNameLabel.Text = "Deployment Name:"
$suDeployNameLabel.Location = New-Object System.Drawing.Point(10, 140)
$suDeployNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabSoftwareUpdate.Controls.Add($suDeployNameLabel)

$suDeployNameTextBox = New-Object System.Windows.Forms.TextBox
$suDeployNameTextBox.Location = New-Object System.Drawing.Point(150, 140)
$suDeployNameTextBox.Size = New-Object System.Drawing.Size(300, 20)
$tabSoftwareUpdate.Controls.Add($suDeployNameTextBox)

$suDeployButton = New-Object System.Windows.Forms.Button
$suDeployButton.Text = "Deploy Updates"
$suDeployButton.Location = New-Object System.Drawing.Point(150, 170)
$suDeployButton.Size = New-Object System.Drawing.Size(150, 30)
$suDeployButton.Add_Click({
    try {
        $sugName = $sugNameComboBox.Text.Trim()
        $collName = $suCollectionComboBox.Text.Trim()
        $deployName = $suDeployNameTextBox.Text.Trim()
        
        if (-not $sugName -or -not $collName -or -not $deployName) {
            [System.Windows.Forms.MessageBox]::Show("Please enter Update Group Name, Collection Name, and Deployment Name", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        Write-Output -Text "Deploying software update group '$sugName' to collection '$collName'..." -Color "Blue"
        
        # Verify update group exists
        $updateGroup = Get-CMSoftwareUpdateGroup -Name $sugName -ErrorAction Stop
        if (-not $updateGroup) {
            Write-Output -Text "Software update group '$sugName' not found" -Color "Red"
            return
        }
        
        # Verify collection exists
        $collection = Get-CMCollection -Name $collName -ErrorAction Stop
        if (-not $collection) {
            Write-Output -Text "Collection '$collName' not found" -Color "Red"
            return
        }
        
        # Set deployment parameters
        $deploymentParams = @{
            SoftwareUpdateGroupName = $sugName
            CollectionName = $collName
            DeploymentName = $deployName
            DeploymentType = "Required"
            EnforcementDeadline = $suDeadlinePicker.Value
            TimeBasedOn = "LocalTime"
            UserNotification = "DisplayAll"
            AvailableDateTime = (Get-Date)
            OverrideServiceWindows = $suOverrideCheck.Checked
            RestartServer = $false
            RestartWorkstation = $false
            SendWakeupPacket = $true
            VerboseLevel = "AllMessages"
        }
        
        # Set download option
        if ($suDownloadCheck.Checked) {
            $deploymentParams.Add("DownloadFromMicrosoftUpdate", $false)
            $deploymentParams.Add("SoftwareInstallation", $true)
        } else {
            $deploymentParams.Add("DownloadFromMicrosoftUpdate", $true)
            $deploymentParams.Add("SoftwareInstallation", $false)
        }
        
        # Create deployment
        $deployment = New-CMSoftwareUpdateDeployment @deploymentParams -ErrorAction Stop
        
        Write-Output -Text "Successfully deployed software update group '$sugName' to collection '$collName'" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to create software update deployment: $_" -Color "Red"
    }
})
$tabSoftwareUpdate.Controls.Add($suDeployButton)

$tabControl.Controls.Add($tabSoftwareUpdate)

# Collection Management Tab
$tabCollection = New-Object System.Windows.Forms.TabPage
$tabCollection.Text = "Collection Management"

$collNameLabel = New-Object System.Windows.Forms.Label
$collNameLabel.Text = "Collection Name:"
$collNameLabel.Location = New-Object System.Drawing.Point(10, 20)
$collNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$tabCollection.Controls.Add($collNameLabel)

$collNameComboBox = New-Object System.Windows.Forms.ComboBox
$collNameComboBox.Location = New-Object System.Drawing.Point(150, 20)
$collNameComboBox.Size = New-Object System.Drawing.Size(300, 20)
$collNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$collNameComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$collNameComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$tabCollection.Controls.Add($collNameComboBox)

$collRefreshButton = New-Object System.Windows.Forms.Button
$collRefreshButton.Text = "Refresh"
$collRefreshButton.Location = New-Object System.Drawing.Point(460, 20)
$collRefreshButton.Size = New-Object System.Drawing.Size(80, 20)
$collRefreshButton.Add_Click({
    try {
        Write-Output -Text "Retrieving device collections..." -Color "Blue"
        $collNameComboBox.Items.Clear()
        Get-CMDeviceCollection | ForEach-Object {
            $collNameComboBox.Items.Add($_.Name)
        }
        Write-Output -Text "Device collections retrieved successfully" -Color "Green"
    }
    catch {
        Write-Output -Text "Failed to retrieve collections: $_" -Color "Red"
    }
})
$tabCollection.Controls.Add($collRefreshButton)

# Add device section
$addDeviceGroupBox = New-Object System.Windows.Forms.GroupBox
$addDeviceGroupBox.Text = "Add Device to Collection"
$addDeviceGroupBox.Location = New-Object System.Drawing.Point(10, 60)
$addDeviceGroupBox.Size = New-Object System.Drawing.Size(355, 120)
$tabCollection.Controls.Add($addDeviceGroupBox)

$deviceNameLabel = New-Object System.Windows.Forms.Label
$deviceNameLabel.Text = "Device Name:"
$deviceNameLabel.Location = New-Object System.Drawing.Point(10, 25)
$deviceNameLabel.Size = New-Object System.Drawing.Size(100, 20)
$addDeviceGroupBox.Controls.Add($deviceNameLabel)

$deviceNameTextBox = New-Object System.Windows.Forms.TextBox
$deviceNameTextBox.Location = New-Object System.Drawing.Point(120, 25)
$deviceNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
$addDeviceGroupBox.Controls.Add($deviceNameTextBox)

$bulkAddCheck = New-Object System.Windows.Forms.CheckBox
$bulkAddCheck.Text = "Bulk Add (one device per line)"
$bulkAddCheck.Location = New-Object System.Drawing.Point(10, 55)
$bulkAddCheck.Size = New-Object System.Drawing.Size(200, 20)
$addDeviceGroupBox.Controls.Add($bulkAddCheck)

$addDeviceButton = New-Object System.Windows.Forms.Button
$addDeviceButton.Text = "Add Device"
$addDeviceButton.Location = New-Object System.Drawing.Point(120, 80)
$addDeviceButton.Size = New-Object System.Drawing.Size(120, 30)
$addDeviceButton.Add_Click({
    try {
        $collName = $collNameComboBox.Text.Trim()
        $deviceName = $deviceNameTextBox.Text.Trim()
        
        if (-not $collName) {
            [System.Windows.Forms.MessageBox]::Show("Please select a Collection", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        if (-not $deviceName) {
            [System.Windows.Forms.MessageBox]::Show("Please enter Device Name(s)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Verify collection exists
        $collection = Get-CMCollection -Name $collName -ErrorAction Stop
        if (-not $collection) {
            Write-Output -Text "Collection '$collName' not found" -Color "Red"
            return
        }
        
        # Check if bulk add is selected
        if ($bulkAddCheck.Checked) {
            $devices = $deviceName -split "`r`n" | Where-Object { $_ -ne "" }
            
            Write-Output -Text "Adding $($devices.Count) devices to collection '$collName'..." -Color "Blue"
            
            foreach ($device in $devices) {
                try {
                    Add-CMDeviceCollectionDirectMembershipRule -CollectionName $collName -ResourceId (Get-CMDevice -Name $device).ResourceID -ErrorAction Stop
                    Write-Output -Text "Successfully added device '$device' to collection '$collName'" -Color "Green"
                }
                catch {
                    Write-Output -Text "Failed to add device '$device': $_" -Color "Red"
                }
            }
        }
        else {
            Write-Output -Text "Adding device '$deviceName' to collection '$collName'..." -Color "Blue"
            
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $collName -ResourceId (Get-CMDevice -Name $deviceName).ResourceID -ErrorAction Stop
            
            Write-Output -Text "Successfully added device '$deviceName' to collection '$collName'" -Color "Green"
        }
    }
    catch {
        Write-Output -Text "Failed to add device(s): $_" -Color "Red"
    }
})
$addDeviceGroupBox.Controls.Add($addDeviceButton)

# Remove device section
$removeDeviceGroupBox = New-Object System.Windows.Forms.GroupBox
$removeDeviceGroupBox.Text = "Remove Device from Collection"
$removeDeviceGroupBox.Location = New-Object System.Drawing.Point(375, 60)
$removeDeviceGroupBox.Size = New-Object System.Drawing.Size(355, 120)
$tabCollection.Controls.Add($removeDeviceGroupBox)

$removeDeviceLabel = New-Object System.Windows.Forms.Label
$removeDeviceLabel.Text = "Device Name:"
$removeDeviceLabel.Location = New-Object System.Drawing.Point(10, 25)
$removeDeviceLabel.Size = New-Object System.Drawing.Size(100, 20)
$removeDeviceGroupBox.Controls.Add($removeDeviceLabel)

$removeDeviceTextBox = New-Object System.Windows.Forms.TextBox
$removeDeviceTextBox.Location = New-Object System.Drawing.Point(120, 25)
$removeDeviceTextBox.Size = New-Object System.Drawing.Size(200, 20)
$removeDeviceGroupBox.Controls.Add($removeDeviceTextBox)

$bulkRemoveCheck = New-Object System.Windows.Forms.CheckBox
$bulkRemoveCheck.Text = "Bulk Remove (one device per line)"
$bulkRemoveCheck.Location = New-Object System.Drawing.Point(10, 55)
$bulkRemoveCheck.Size = New-Object System.Drawing.Size(200, 20)
$removeDeviceGroupBox.Controls.Add($bulkRemoveCheck)

$removeDeviceButton = New-Object System.Windows.Forms.Button
$removeDeviceButton.Text = "Remove Device"
$removeDeviceButton.Location = New-Object System.Drawing.Point(120, 80)
$removeDeviceButton.Size = New-Object System.Drawing.Size(120, 30)
$removeDeviceButton.Add_Click({
    try {
        $collName = $collNameComboBox.Text.Trim()
        $deviceName = $removeDeviceTextBox.Text.Trim()
        
        if (-not $collName) {
            [System.Windows.Forms.MessageBox]::Show("Please select a Collection", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        if (-not $deviceName) {
            [System.Windows.Forms.MessageBox]::Show("Please enter Device Name(s)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Verify collection exists
        $collection = Get-CMCollection -Name $collName -ErrorAction Stop
        if (-not $collection) {
            Write-Output -Text "Collection '$collName' not found" -Color "Red"
            return
        }
        
        # Check if bulk remove is selected
        if ($bulkRemoveCheck.Checked) {
            $devices = $deviceName -split "`r`n" | Where-Object { $_ -ne "" }
            
            Write-Output -Text "Removing $($devices.Count) devices from collection '$collName'..." -Color "Blue"
            
            foreach ($device in $devices) {
                try {
                    $resourceId = (Get-CMDevice -Name $device).ResourceID
                    $rule = Get-CMDeviceCollectionDirectMembershipRule -CollectionName $collName -ResourceId $resourceId -ErrorAction Stop
                    Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $collName -ResourceId $resourceId -Force -ErrorAction Stop
                    Write-Output -Text "Successfully removed device '$device' from collection '$collName'" -Color "Green"
                }
                catch {
                    Write-Output -Text "Failed to remove device '$device': $_" -Color "Red"
                }
            }
        }
        else {
            Write-Output -Text "Removing device '$deviceName' from collection '$collName'..." -Color "Blue"
            
            $resourceId = (Get-CMDevice -Name $deviceName).ResourceID
            $rule = Get-CMDeviceCollectionDirectMembershipRule -CollectionName $collName -ResourceId $resourceId -ErrorAction Stop
            Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $collName -ResourceId $resourceId -Force -ErrorAction Stop
            
            Write-Output -Text "Successfully removed device '$deviceName' from collection '$collName'" -Color "Green"
        }
    }
    catch {
        Write-Output -Text "Failed to remove device(s): $_" -Color "Red"
    }
})
$removeDeviceGroupBox.Controls.Add($removeDeviceButton)

# Rename collection section
$renameCollGroupBox = New-Object System.Windows.Forms.GroupBox
$renameCollGroupBox.Text = "Rename Collection"
$renameCollGroupBox.Location = New-Object System.Drawing.Point(10, 190)
$renameCollGroupBox.Size = New-Object System.Drawing.Size(730, 80)
$tabCollection.Controls.Add($renameCollGroupBox)

$newNameLabel = New-Object System.Windows.Forms.Label
$newNameLabel.Text = "New Name:"
$newNameLabel.Location = New-Object System.Drawing.Point(10, 25)
$newNameLabel.Size = New-Object System.Drawing.Size(100, 20)
$renameCollGroupBox.Controls.Add($newNameLabel)

$newNameTextBox = New-Object System.Windows.Forms.TextBox
$newNameTextBox.Location = New-Object System.Drawing.Point(120, 25)
$newNameTextBox.Size = New-Object System.Drawing.Size(400, 20)
$renameCollGroupBox.Controls.Add($newNameTextBox)

$renameCollButton = New-Object System.Windows.Forms.Button
$renameCollButton.Text = "Rename Collection"
$renameCollButton.Location = New-Object System.Drawing.Point(540, 25)
$renameCollButton.Size = New-Object System.Drawing.Size(150, 30)
$renameCollButton.Add_Click({
    try {
        $collName = $collNameComboBox.Text.Trim()
        $newName = $newNameTextBox.Text.Trim()
        
        if (-not $collName) {
            [System.Windows.Forms.MessageBox]::Show("Please select a Collection", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        if (-not $newName) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a New Name", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        Write-Output -Text "Renaming collection '$collName' to '$newName'..." -Color "Blue"
        
        # Get the collection
        $collection = Get-CMDeviceCollection -Name $collName -ErrorAction Stop
        if (-not $collection) {
            Write-Output -Text "Collection '$collName' not found" -Color "Red"
            return
        }
        
        # Rename the collection
        Set-CMDeviceCollection -InputObject $collection -NewName $newName -ErrorAction Stop
        
        Write-Output -Text "Successfully renamed collection '$collName' to '$newName'" -Color "Green"
        
        # Refresh the collection comboboxes
        $collNameComboBox.Items.Clear()
        $appCollectionComboBox.Items.Clear()
        $pkgCollectionComboBox.Items.Clear()
        $suCollectionComboBox.Items.Clear()
        
        Get-CMDeviceCollection | ForEach-Object {
            $collNameComboBox.Items.Add($_.Name)
            $appCollectionComboBox.Items.Add($_.Name)
            $pkgCollectionComboBox.Items.Add($_.Name)
            $suCollectionComboBox.Items.Add($_.Name)
        }
        
        # Update selected item
        $collNameComboBox.Text = $newName
    }
    catch {
        Write-Output -Text "Failed to rename collection: $_" -Color "Red"
    }
})
$renameCollGroupBox.Controls.Add($renameCollButton)

$tabControl.Controls.Add($tabCollection)

# Add the main form's OnLoad event handler
$mainForm.Add_Load({
    # Display welcome message
    Write-Output -Text "Welcome to MECM Deployment Automation Tool" -Color "Blue"
    Write-Output -Text "Please enter your MECM Site Code and Site Server to connect" -Color "Black"
})

# Show the form
$mainForm.ShowDialog()