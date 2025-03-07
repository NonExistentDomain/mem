#Requires -Modules ConfigurationManager
<#
.SYNOPSIS
    GUI tool for querying installed software across multiple computers using MECM/SCCM.
.DESCRIPTION
    This script provides a graphical interface to search for installed software
    across one or multiple computers in your MECM/SCCM environment.
.NOTES
    Requires the ConfigurationManager PowerShell module and appropriate permissions.
    Run from a PowerShell console with the ConfigurationManager module loaded.
#>

# Load required assemblies for the GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to connect to the MECM site
function Connect-MECMSite {
    param (
        [string]$SiteCode,
        [string]$SiteServer
    )
    
    try {
        # Import the ConfigurationManager module
        Import-Module ConfigurationManager -ErrorAction Stop
        
        # Set the location to the MECM site
        $CMDrive = $SiteCode + ":"
        if (!(Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
        }
        Set-Location $CMDrive -ErrorAction Stop
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to MECM site: $_", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Function to get installed software for the specified computers
function Get-MECMInstalledSoftware {
    param (
        [string[]]$ComputerNames,
        [string[]]$SoftwareNames
    )
    
    $results = @()
    
    foreach ($computer in $ComputerNames) {
        $computerID = Get-CMDevice -Name $computer -Fast | Select-Object -ExpandProperty ResourceID
        
        if ($computerID) {
            foreach ($software in $SoftwareNames) {
                $query = "SELECT 
                            SYS.Name0 as 'ComputerName',
                            SOFT.DisplayName0 as 'SoftwareName',
                            SOFT.Version0 as 'Version',
                            SOFT.Publisher0 as 'Publisher',
                            SOFT.InstallDate0 as 'InstallDate'
                          FROM 
                            v_R_System SYS
                            JOIN v_Add_Remove_Programs SOFT ON SYS.ResourceID = SOFT.ResourceID
                          WHERE 
                            SYS.ResourceID = '$computerID'
                            AND SOFT.DisplayName0 LIKE '%$software%'"
                
                $result = Invoke-CMWmiQuery -Query $query
                
                if ($result) {
                    foreach ($item in $result) {
                        $results += [PSCustomObject]@{
                            ComputerName = $item.ComputerName
                            SoftwareName = $item.SoftwareName
                            Version = $item.Version
                            Publisher = $item.Publisher
                            InstallDate = $item.InstallDate
                        }
                    }
                }
            }
        }
        else {
            $results += [PSCustomObject]@{
                ComputerName = $computer
                SoftwareName = "Computer not found in MECM"
                Version = ""
                Publisher = ""
                InstallDate = ""
            }
        }
    }
    
    return $results
}

# Function to export results to CSV
function Export-ResultsToCSV {
    param (
        $Results,
        $FilePath
    )
    
    try {
        $Results | Export-Csv -Path $FilePath -NoTypeInformation -ErrorAction Stop
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to export results: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MECM Software Query Tool"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# MECM Connection GroupBox
$connectionGroupBox = New-Object System.Windows.Forms.GroupBox
$connectionGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$connectionGroupBox.Size = New-Object System.Drawing.Size(760, 80)
$connectionGroupBox.Text = "MECM Connection"

# Site Code Label and TextBox
$siteCodeLabel = New-Object System.Windows.Forms.Label
$siteCodeLabel.Location = New-Object System.Drawing.Point(10, 25)
$siteCodeLabel.Size = New-Object System.Drawing.Size(70, 20)
$siteCodeLabel.Text = "Site Code:"

$siteCodeTextBox = New-Object System.Windows.Forms.TextBox
$siteCodeTextBox.Location = New-Object System.Drawing.Point(90, 22)
$siteCodeTextBox.Size = New-Object System.Drawing.Size(100, 20)

# Site Server Label and TextBox
$siteServerLabel = New-Object System.Windows.Forms.Label
$siteServerLabel.Location = New-Object System.Drawing.Point(210, 25)
$siteServerLabel.Size = New-Object System.Drawing.Size(70, 20)
$siteServerLabel.Text = "Site Server:"

$siteServerTextBox = New-Object System.Windows.Forms.TextBox
$siteServerTextBox.Location = New-Object System.Drawing.Point(290, 22)
$siteServerTextBox.Size = New-Object System.Drawing.Size(200, 20)

# Connect Button
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = New-Object System.Drawing.Point(500, 21)
$connectButton.Size = New-Object System.Drawing.Size(100, 23)
$connectButton.Text = "Connect"

# Connection Status Label
$connectionStatusLabel = New-Object System.Windows.Forms.Label
$connectionStatusLabel.Location = New-Object System.Drawing.Point(610, 25)
$connectionStatusLabel.Size = New-Object System.Drawing.Size(140, 20)
$connectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
$connectionStatusLabel.Text = "Not Connected"

# Search GroupBox
$searchGroupBox = New-Object System.Windows.Forms.GroupBox
$searchGroupBox.Location = New-Object System.Drawing.Point(10, 100)
$searchGroupBox.Size = New-Object System.Drawing.Size(760, 130)
$searchGroupBox.Text = "Search Criteria"
$searchGroupBox.Enabled = $false

# Computer Names Label and TextBox
$computerNamesLabel = New-Object System.Windows.Forms.Label
$computerNamesLabel.Location = New-Object System.Drawing.Point(10, 25)
$computerNamesLabel.Size = New-Object System.Drawing.Size(110, 20)
$computerNamesLabel.Text = "Computer Names:"

$computerNamesTextBox = New-Object System.Windows.Forms.TextBox
$computerNamesTextBox.Location = New-Object System.Drawing.Point(130, 22)
$computerNamesTextBox.Size = New-Object System.Drawing.Size(600, 20)
$computerNamesTextBox.PlaceholderText = "Enter computer names separated by commas (e.g., PC1, PC2, PC3)"

# Software Names Label and TextBox
$softwareNamesLabel = New-Object System.Windows.Forms.Label
$softwareNamesLabel.Location = New-Object System.Drawing.Point(10, 55)
$softwareNamesLabel.Size = New-Object System.Drawing.Size(110, 20)
$softwareNamesLabel.Text = "Software Names:"

$softwareNamesTextBox = New-Object System.Windows.Forms.TextBox
$softwareNamesTextBox.Location = New-Object System.Drawing.Point(130, 52)
$softwareNamesTextBox.Size = New-Object System.Drawing.Size(600, 20)
$softwareNamesTextBox.PlaceholderText = "Enter software names separated by commas (e.g., Office, Adobe, Chrome)"

# Search Button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(330, 90)
$searchButton.Size = New-Object System.Drawing.Size(100, 23)
$searchButton.Text = "Search"

# Results GroupBox
$resultsGroupBox = New-Object System.Windows.Forms.GroupBox
$resultsGroupBox.Location = New-Object System.Drawing.Point(10, 240)
$resultsGroupBox.Size = New-Object System.Drawing.Size(760, 270)
$resultsGroupBox.Text = "Results"

# Results DataGridView
$resultsDataGridView = New-Object System.Windows.Forms.DataGridView
$resultsDataGridView.Location = New-Object System.Drawing.Point(10, 20)
$resultsDataGridView.Size = New-Object System.Drawing.Size(740, 200)
$resultsDataGridView.AllowUserToAddRows = $false
$resultsDataGridView.AllowUserToDeleteRows = $false
$resultsDataGridView.AllowUserToResizeRows = $false
$resultsDataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$resultsDataGridView.ReadOnly = $true
$resultsDataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$resultsDataGridView.RowHeadersVisible = $false
$resultsDataGridView.BackgroundColor = [System.Drawing.Color]::White

# Export Button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(650, 230)
$exportButton.Size = New-Object System.Drawing.Size(100, 23)
$exportButton.Text = "Export to CSV"
$exportButton.Enabled = $false

# Status Strip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusStrip.Items.Add($statusLabel)

# Add controls to the form
$connectionGroupBox.Controls.Add($siteCodeLabel)
$connectionGroupBox.Controls.Add($siteCodeTextBox)
$connectionGroupBox.Controls.Add($siteServerLabel)
$connectionGroupBox.Controls.Add($siteServerTextBox)
$connectionGroupBox.Controls.Add($connectButton)
$connectionGroupBox.Controls.Add($connectionStatusLabel)

$searchGroupBox.Controls.Add($computerNamesLabel)
$searchGroupBox.Controls.Add($computerNamesTextBox)
$searchGroupBox.Controls.Add($softwareNamesLabel)
$searchGroupBox.Controls.Add($softwareNamesTextBox)
$searchGroupBox.Controls.Add($searchButton)

$resultsGroupBox.Controls.Add($resultsDataGridView)
$resultsGroupBox.Controls.Add($exportButton)

$form.Controls.Add($connectionGroupBox)
$form.Controls.Add($searchGroupBox)
$form.Controls.Add($resultsGroupBox)
$form.Controls.Add($statusStrip)

# Connect Button Click Event
$connectButton.Add_Click({
    $siteCode = $siteCodeTextBox.Text.Trim()
    $siteServer = $siteServerTextBox.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($siteCode) -or [string]::IsNullOrWhiteSpace($siteServer)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter both Site Code and Site Server.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Connecting to MECM site..."
    
    if (Connect-MECMSite -SiteCode $siteCode -SiteServer $siteServer) {
        $connectionStatusLabel.Text = "Connected"
        $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Green
        $searchGroupBox.Enabled = $true
        $statusLabel.Text = "Connected to MECM site. Ready to search."
    }
    else {
        $statusLabel.Text = "Failed to connect to MECM site."
    }
})

# Search Button Click Event
$searchButton.Add_Click({
    $computerNames = $computerNamesTextBox.Text.Trim()
    $softwareNames = $softwareNamesTextBox.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($computerNames) -or [string]::IsNullOrWhiteSpace($softwareNames)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter both Computer Names and Software Names.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $computerNamesArray = $computerNames -split ',' | ForEach-Object { $_.Trim() }
    $softwareNamesArray = $softwareNames -split ',' | ForEach-Object { $_.Trim() }
    
    $statusLabel.Text = "Searching for software..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Clear existing results
    $resultsDataGridView.DataSource = $null
    $resultsDataGridView.Rows.Clear()
    
    # Get results
    $searchResults = Get-MECMInstalledSoftware -ComputerNames $computerNamesArray -SoftwareNames $softwareNamesArray
    
    if ($searchResults.Count -gt 0) {
        $resultsDataGridView.DataSource = [System.Collections.ArrayList]$searchResults
        $exportButton.Enabled = $true
        $statusLabel.Text = "$($searchResults.Count) result(s) found."
    }
    else {
        $statusLabel.Text = "No results found."
        $exportButton.Enabled = $false
    }
    
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# Export Button Click Event
$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Export Results to CSV"
    $saveFileDialog.FileName = "MECM_Software_Query_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $statusLabel.Text = "Exporting results..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        $resultsData = @()
        foreach ($row in $resultsDataGridView.Rows) {
            $resultsData += [PSCustomObject]@{
                ComputerName = $row.Cells["ComputerName"].Value
                SoftwareName = $row.Cells["SoftwareName"].Value
                Version = $row.Cells["Version"].Value
                Publisher = $row.Cells["Publisher"].Value
                InstallDate = $row.Cells["InstallDate"].Value
            }
        }
        
        if (Export-ResultsToCSV -Results $resultsData -FilePath $saveFileDialog.FileName) {
            $statusLabel.Text = "Results exported to: $($saveFileDialog.FileName)"
        }
        else {
            $statusLabel.Text = "Failed to export results."
        }
        
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Show the form
$form.ShowDialog() | Out-Null

# Clean up
$form.Dispose()