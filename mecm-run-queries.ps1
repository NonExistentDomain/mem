#Requires -Modules ConfigurationManager
<#
.SYNOPSIS
    Advanced GUI tool for running any queries in MECM/SCCM.
.DESCRIPTION
    This script provides a graphical interface to run both prebuilt and custom
    queries across one or multiple computers in your MECM/SCCM environment.
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

# Function to get all prebuilt MECM queries
function Get-MECMPrebuiltQueries {
    try {
        $queries = Get-CMQuery | Select-Object Name, Expression, @{Name="Type";Expression={"Prebuilt"}}
        return $queries
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to get prebuilt queries: $_", "Query Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

# Function to execute a WQL query
function Invoke-MECMQuery {
    param (
        [string]$QueryText,
        [string[]]$ComputerNames = @(),
        [switch]$FilterByComputers = $false
    )
    
    try {
        # If filtering by computers is enabled and we have computer names
        if ($FilterByComputers -and $ComputerNames.Count -gt 0) {
            # Get ResourceIDs for the computers
            $computerResourceIDs = @()
            foreach ($computer in $ComputerNames) {
                $computerDevice = Get-CMDevice -Name $computer -Fast
                if ($computerDevice) {
                    $computerResourceIDs += $computerDevice.ResourceID
                }
            }
            
            if ($computerResourceIDs.Count -gt 0) {
                # Add WHERE clause for ResourceIDs if not already in the query
                if ($QueryText -match "WHERE" -or $QueryText -match "where") {
                    $resourceFilter = "AND SYS.ResourceID IN ('" + ($computerResourceIDs -join "','") + "')"
                    $QueryText = $QueryText -replace "(WHERE|where)(.+?)$", "`$1`$2 $resourceFilter"
                }
                else {
                    $resourceFilter = "WHERE SYS.ResourceID IN ('" + ($computerResourceIDs -join "','") + "')"
                    $QueryText = "$QueryText $resourceFilter"
                }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("No valid computers found in MECM.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return @()
            }
        }
        
        # Execute the query
        $results = Invoke-CMWmiQuery -Query $QueryText
        return $results
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to execute query: $_", "Query Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

# Function to execute a prebuilt query
function Invoke-MECMPrebuiltQuery {
    param (
        [string]$QueryName,
        [string[]]$ComputerNames = @(),
        [switch]$FilterByComputers = $false
    )
    
    try {
        # Get the prebuilt query object
        $query = Get-CMQuery -Name $QueryName
        if ($query) {
            # Execute the query with the same filtering logic
            return Invoke-MECMQuery -QueryText $query.Expression -ComputerNames $ComputerNames -FilterByComputers $FilterByComputers
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Prebuilt query '$QueryName' not found.", "Query Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return @()
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to execute prebuilt query: $_", "Query Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
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
$form.Text = "Advanced MECM Query Tool"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# MECM Connection GroupBox
$connectionGroupBox = New-Object System.Windows.Forms.GroupBox
$connectionGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$connectionGroupBox.Size = New-Object System.Drawing.Size(860, 80)
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

# Query Type GroupBox
$queryTypeGroupBox = New-Object System.Windows.Forms.GroupBox
$queryTypeGroupBox.Location = New-Object System.Drawing.Point(10, 100)
$queryTypeGroupBox.Size = New-Object System.Drawing.Size(860, 60)
$queryTypeGroupBox.Text = "Query Type"
$queryTypeGroupBox.Enabled = $false

# Query Type Radio Buttons
$prebuiltQueryRadioButton = New-Object System.Windows.Forms.RadioButton
$prebuiltQueryRadioButton.Location = New-Object System.Drawing.Point(20, 25)
$prebuiltQueryRadioButton.Size = New-Object System.Drawing.Size(150, 20)
$prebuiltQueryRadioButton.Text = "Prebuilt MECM Query"
$prebuiltQueryRadioButton.Checked = $true

$customQueryRadioButton = New-Object System.Windows.Forms.RadioButton
$customQueryRadioButton.Location = New-Object System.Drawing.Point(200, 25)
$customQueryRadioButton.Size = New-Object System.Drawing.Size(150, 20)
$customQueryRadioButton.Text = "Custom WQL Query"

# Apply Computer Filter Checkbox
$filterComputersCheckBox = New-Object System.Windows.Forms.CheckBox
$filterComputersCheckBox.Location = New-Object System.Drawing.Point(400, 25)
$filterComputersCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$filterComputersCheckBox.Text = "Filter Results by Computer Names"

# Prebuilt Query GroupBox
$prebuiltQueryGroupBox = New-Object System.Windows.Forms.GroupBox
$prebuiltQueryGroupBox.Location = New-Object System.Drawing.Point(10, 170)
$prebuiltQueryGroupBox.Size = New-Object System.Drawing.Size(860, 80)
$prebuiltQueryGroupBox.Text = "Prebuilt Query Selection"
$prebuiltQueryGroupBox.Enabled = $false

# Prebuilt Query ComboBox and Refresh Button
$prebuiltQueryLabel = New-Object System.Windows.Forms.Label
$prebuiltQueryLabel.Location = New-Object System.Drawing.Point(10, 25)
$prebuiltQueryLabel.Size = New-Object System.Drawing.Size(100, 20)
$prebuiltQueryLabel.Text = "Select Query:"

$prebuiltQueryComboBox = New-Object System.Windows.Forms.ComboBox
$prebuiltQueryComboBox.Location = New-Object System.Drawing.Point(120, 22)
$prebuiltQueryComboBox.Size = New-Object System.Drawing.Size(500, 20)
$prebuiltQueryComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$refreshQueriesButton = New-Object System.Windows.Forms.Button
$refreshQueriesButton.Location = New-Object System.Drawing.Point(630, 21)
$refreshQueriesButton.Size = New-Object System.Drawing.Size(100, 23)
$refreshQueriesButton.Text = "Refresh Queries"

# Custom Query GroupBox
$customQueryGroupBox = New-Object System.Windows.Forms.GroupBox
$customQueryGroupBox.Location = New-Object System.Drawing.Point(10, 170)
$customQueryGroupBox.Size = New-Object System.Drawing.Size(860, 120)
$customQueryGroupBox.Text = "Custom WQL Query"
$customQueryGroupBox.Enabled = $false
$customQueryGroupBox.Visible = $false

# Custom Query TextBox
$customQueryTextBox = New-Object System.Windows.Forms.TextBox
$customQueryTextBox.Location = New-Object System.Drawing.Point(10, 20)
$customQueryTextBox.Size = New-Object System.Drawing.Size(840, 90)
$customQueryTextBox.Multiline = $true
$customQueryTextBox.ScrollBars = "Vertical"
$customQueryTextBox.PlaceholderText = "Enter your WQL query here (e.g., SELECT * FROM v_R_System)"

# Computer Filter GroupBox
$computerFilterGroupBox = New-Object System.Windows.Forms.GroupBox
$computerFilterGroupBox.Location = New-Object System.Drawing.Point(10, 300)
$computerFilterGroupBox.Size = New-Object System.Drawing.Size(860, 80)
$computerFilterGroupBox.Text = "Computer Filter (Optional)"
$computerFilterGroupBox.Enabled = $false

# Computer Names TextBox
$computerNamesLabel = New-Object System.Windows.Forms.Label
$computerNamesLabel.Location = New-Object System.Drawing.Point(10, 25)
$computerNamesLabel.Size = New-Object System.Drawing.Size(110, 20)
$computerNamesLabel.Text = "Computer Names:"

$computerNamesTextBox = New-Object System.Windows.Forms.TextBox
$computerNamesTextBox.Location = New-Object System.Drawing.Point(130, 22)
$computerNamesTextBox.Size = New-Object System.Drawing.Size(700, 20)
$computerNamesTextBox.PlaceholderText = "Enter computer names separated by commas (e.g., PC1, PC2, PC3)"

# Execute Query Button
$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Location = New-Object System.Drawing.Point(380, 390)
$executeButton.Size = New-Object System.Drawing.Size(120, 30)
$executeButton.Text = "Execute Query"
$executeButton.Enabled = $false

# Results GroupBox
$resultsGroupBox = New-Object System.Windows.Forms.GroupBox
$resultsGroupBox.Location = New-Object System.Drawing.Point(10, 430)
$resultsGroupBox.Size = New-Object System.Drawing.Size(860, 180)
$resultsGroupBox.Text = "Results"

# Results DataGridView
$resultsDataGridView = New-Object System.Windows.Forms.DataGridView
$resultsDataGridView.Location = New-Object System.Drawing.Point(10, 20)
$resultsDataGridView.Size = New-Object System.Drawing.Size(840, 110)
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
$exportButton.Location = New-Object System.Drawing.Point(750, 140)
$exportButton.Size = New-Object System.Drawing.Size(100, 23)
$exportButton.Text = "Export to CSV"
$exportButton.Enabled = $false

# Query Details GroupBox
$queryDetailsGroupBox = New-Object System.Windows.Forms.GroupBox
$queryDetailsGroupBox.Location = New-Object System.Drawing.Point(10, 620)
$queryDetailsGroupBox.Size = New-Object System.Drawing.Size(860, 40)
$queryDetailsGroupBox.Text = "Query Details"

# Query Details Label
$queryDetailsLabel = New-Object System.Windows.Forms.Label
$queryDetailsLabel.Location = New-Object System.Drawing.Point(10, 15)
$queryDetailsLabel.Size = New-Object System.Drawing.Size(840, 20)
$queryDetailsLabel.Text = "No query executed yet."

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

$queryTypeGroupBox.Controls.Add($prebuiltQueryRadioButton)
$queryTypeGroupBox.Controls.Add($customQueryRadioButton)
$queryTypeGroupBox.Controls.Add($filterComputersCheckBox)

$prebuiltQueryGroupBox.Controls.Add($prebuiltQueryLabel)
$prebuiltQueryGroupBox.Controls.Add($prebuiltQueryComboBox)
$prebuiltQueryGroupBox.Controls.Add($refreshQueriesButton)

$customQueryGroupBox.Controls.Add($customQueryTextBox)

$computerFilterGroupBox.Controls.Add($computerNamesLabel)
$computerFilterGroupBox.Controls.Add($computerNamesTextBox)

$resultsGroupBox.Controls.Add($resultsDataGridView)
$resultsGroupBox.Controls.Add($exportButton)

$queryDetailsGroupBox.Controls.Add($queryDetailsLabel)

$form.Controls.Add($connectionGroupBox)
$form.Controls.Add($queryTypeGroupBox)
$form.Controls.Add($prebuiltQueryGroupBox)
$form.Controls.Add($customQueryGroupBox)
$form.Controls.Add($computerFilterGroupBox)
$form.Controls.Add($executeButton)
$form.Controls.Add($resultsGroupBox)
$form.Controls.Add($queryDetailsGroupBox)
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
        $queryTypeGroupBox.Enabled = $true
        $prebuiltQueryGroupBox.Enabled = $true
        $computerFilterGroupBox.Enabled = $true
        $executeButton.Enabled = $true
        
        # Load prebuilt queries
        $statusLabel.Text = "Loading prebuilt queries..."
        $prebuiltQueries = Get-MECMPrebuiltQueries
        $prebuiltQueryComboBox.Items.Clear()
        foreach ($query in $prebuiltQueries) {
            $prebuiltQueryComboBox.Items.Add($query.Name)
        }
        if ($prebuiltQueryComboBox.Items.Count -gt 0) {
            $prebuiltQueryComboBox.SelectedIndex = 0
        }
        
        $statusLabel.Text = "Connected to MECM site. Ready to run queries."
    }
    else {
        $statusLabel.Text = "Failed to connect to MECM site."
    }
})

# Query Type Radio Button Change Events
$prebuiltQueryRadioButton.Add_CheckedChanged({
    if ($prebuiltQueryRadioButton.Checked) {
        $prebuiltQueryGroupBox.Visible = $true
        $prebuiltQueryGroupBox.Enabled = $true
        $customQueryGroupBox.Visible = $false
        $customQueryGroupBox.Enabled = $false
    }
})

$customQueryRadioButton.Add_CheckedChanged({
    if ($customQueryRadioButton.Checked) {
        $prebuiltQueryGroupBox.Visible = $false
        $prebuiltQueryGroupBox.Enabled = $false
        $customQueryGroupBox.Visible = $true
        $customQueryGroupBox.Enabled = $true
    }
})

# Filter Computers CheckBox Change Event
$filterComputersCheckBox.Add_CheckedChanged({
    $computerFilterGroupBox.Enabled = $filterComputersCheckBox.Checked
})

# Refresh Queries Button Click Event
$refreshQueriesButton.Add_Click({
    $statusLabel.Text = "Refreshing prebuilt queries..."
    $prebuiltQueries = Get-MECMPrebuiltQueries
    $prebuiltQueryComboBox.Items.Clear()
    foreach ($query in $prebuiltQueries) {
        $prebuiltQueryComboBox.Items.Add($query.Name)
    }
    if ($prebuiltQueryComboBox.Items.Count -gt 0) {
        $prebuiltQueryComboBox.SelectedIndex = 0
    }
    $statusLabel.Text = "Prebuilt queries refreshed."
})

# Execute Query Button Click Event
$executeButton.Add_Click({
    $filterByComputers = $filterComputersCheckBox.Checked
    $computerNames = @()
    
    if ($filterByComputers) {
        $computerNamesText = $computerNamesTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($computerNamesText)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter computer names to filter by.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $computerNames = $computerNamesText -split ',' | ForEach-Object { $_.Trim() }
    }
    
    $statusLabel.Text = "Executing query..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Clear existing results
    $resultsDataGridView.DataSource = $null
    $resultsDataGridView.Rows.Clear()
    
    # Execute the appropriate query type
    $queryResults = @()
    if ($prebuiltQueryRadioButton.Checked) {
        $selectedQuery = $prebuiltQueryComboBox.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedQuery)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a prebuilt query.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $statusLabel.Text = "Ready"
            return
        }
        
        $queryResults = Invoke-MECMPrebuiltQuery -QueryName $selectedQuery -ComputerNames $computerNames -FilterByComputers:$filterByComputers
        $queryDetailsLabel.Text = "Executed prebuilt query: $selectedQuery"
    }
    else {
        $customQuery = $customQueryTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($customQuery)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a custom WQL query.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $statusLabel.Text = "Ready"
            return
        }
        
        $queryResults = Invoke-MECMQuery -QueryText $customQuery -ComputerNames $computerNames -FilterByComputers:$filterByComputers
        $queryDetailsLabel.Text = "Executed custom query: $($customQuery.Substring(0, [Math]::Min(100, $customQuery.Length)))..."
    }
    
    if ($queryResults.Count -gt 0) {
        $resultsTable = New-Object System.Data.DataTable
        $properties = $queryResults[0].PSObject.Properties | Where-Object { $_.MemberType -eq "NoteProperty" }
        
        foreach ($prop in $properties) {
            [void]$resultsTable.Columns.Add($prop.Name)
        }
        
        foreach ($result in $queryResults) {
            $row = $resultsTable.NewRow()
            foreach ($prop in $properties) {
                $row[$prop.Name] = $result.($prop.Name)
            }
            [void]$resultsTable.Rows.Add($row)
        }
        
        $resultsDataGridView.DataSource = $resultsTable
        $exportButton.Enabled = $true
        $statusLabel.Text = "$($queryResults.Count) result(s) found."
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
    $saveFileDialog.FileName = "MECM_Query_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $statusLabel.Text = "Exporting results..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        $resultsData = @()
        $dataTable = $resultsDataGridView.DataSource
        
        foreach ($row in $dataTable.Rows) {
            $rowData = New-Object PSObject
            foreach ($column in $dataTable.Columns) {
                Add-Member -InputObject $rowData -MemberType NoteProperty -Name $column.ColumnName -Value $row[$column.ColumnName]
            }
            $resultsData += $rowData
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