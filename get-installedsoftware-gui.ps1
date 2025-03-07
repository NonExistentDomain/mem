# Software Inventory Tool
# This script creates a GUI to query installed software on local or remote computers

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Software Inventory Tool"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create computer list group box
$groupBoxComputers = New-Object System.Windows.Forms.GroupBox
$groupBoxComputers.Location = New-Object System.Drawing.Point(20, 20)
$groupBoxComputers.Size = New-Object System.Drawing.Size(320, 200)
$groupBoxComputers.Text = "Computers"
$form.Controls.Add($groupBoxComputers)

# Create computer list textbox
$textBoxComputers = New-Object System.Windows.Forms.TextBox
$textBoxComputers.Location = New-Object System.Drawing.Point(10, 25)
$textBoxComputers.Size = New-Object System.Drawing.Size(300, 150)
$textBoxComputers.Multiline = $true
$textBoxComputers.ScrollBars = "Vertical"
$textBoxComputers.Text = $env:COMPUTERNAME
$groupBoxComputers.Controls.Add($textBoxComputers)

# Create software group box
$groupBoxSoftware = New-Object System.Windows.Forms.GroupBox
$groupBoxSoftware.Location = New-Object System.Drawing.Point(350, 20)
$groupBoxSoftware.Size = New-Object System.Drawing.Size(320, 200)
$groupBoxSoftware.Text = "Software to Query (leave blank for all)"
$form.Controls.Add($groupBoxSoftware)

# Create software list textbox
$textBoxSoftware = New-Object System.Windows.Forms.TextBox
$textBoxSoftware.Location = New-Object System.Drawing.Point(10, 25) 
$textBoxSoftware.Size = New-Object System.Drawing.Size(300, 150)
$textBoxSoftware.Multiline = $true
$textBoxSoftware.ScrollBars = "Vertical"
$groupBoxSoftware.Controls.Add($textBoxSoftware)

# Create options group box
$groupBoxOptions = New-Object System.Windows.Forms.GroupBox
$groupBoxOptions.Location = New-Object System.Drawing.Point(20, 230)
$groupBoxOptions.Size = New-Object System.Drawing.Size(650, 60)
$groupBoxOptions.Text = "Options"
$form.Controls.Add($groupBoxOptions)

# Create credentials checkbox
$checkBoxCredentials = New-Object System.Windows.Forms.CheckBox
$checkBoxCredentials.Location = New-Object System.Drawing.Point(10, 25)
$checkBoxCredentials.Size = New-Object System.Drawing.Size(200, 20)
$checkBoxCredentials.Text = "Use alternate credentials"
$groupBoxOptions.Controls.Add($checkBoxCredentials)

# Create export checkbox
$checkBoxExport = New-Object System.Windows.Forms.CheckBox
$checkBoxExport.Location = New-Object System.Drawing.Point(220, 25)
$checkBoxExport.Size = New-Object System.Drawing.Size(200, 20)
$checkBoxExport.Text = "Export results to CSV"
$groupBoxOptions.Controls.Add($checkBoxExport)

# Create query button
$buttonQuery = New-Object System.Windows.Forms.Button
$buttonQuery.Location = New-Object System.Drawing.Point(430, 25)
$buttonQuery.Size = New-Object System.Drawing.Size(100, 25)
$buttonQuery.Text = "Run Query"
$buttonQuery.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$buttonQuery.ForeColor = [System.Drawing.Color]::White
$buttonQuery.FlatStyle = "Flat"
$groupBoxOptions.Controls.Add($buttonQuery)

# Create results group box
$groupBoxResults = New-Object System.Windows.Forms.GroupBox
$groupBoxResults.Location = New-Object System.Drawing.Point(20, 300)
$groupBoxResults.Size = New-Object System.Drawing.Size(650, 250)
$groupBoxResults.Text = "Results"
$form.Controls.Add($groupBoxResults)

# Create status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(20, 560)
$labelStatus.Size = New-Object System.Drawing.Size(650, 20)
$labelStatus.Text = "Ready"
$form.Controls.Add($labelStatus)

# Create results DataGridView
$dataGridViewResults = New-Object System.Windows.Forms.DataGridView
$dataGridViewResults.Location = New-Object System.Drawing.Point(10, 25)
$dataGridViewResults.Size = New-Object System.Drawing.Size(630, 215)
$dataGridViewResults.AllowUserToAddRows = $false
$dataGridViewResults.AllowUserToDeleteRows = $false
$dataGridViewResults.ReadOnly = $true
$dataGridViewResults.MultiSelect = $true
$dataGridViewResults.AutoSizeColumnsMode = "Fill"
$dataGridViewResults.ColumnHeadersHeightSizeMode = "AutoSize"
$dataGridViewResults.SelectionMode = "FullRowSelect"
$dataGridViewResults.RowHeadersVisible = $false
$dataGridViewResults.BackgroundColor = [System.Drawing.Color]::White
$dataGridViewResults.BorderStyle = "Fixed3D"
$dataGridViewResults.CellBorderStyle = "SingleHorizontal"
$dataGridViewResults.GridColor = [System.Drawing.Color]::LightGray
$groupBoxResults.Controls.Add($dataGridViewResults)

# Query button click event
$buttonQuery.Add_Click({
    # Clear previous results
    $dataGridViewResults.DataSource = $null
    $dataGridViewResults.Rows.Clear()
    $dataGridViewResults.Columns.Clear()
    
    # Get computer list
    $computers = $textBoxComputers.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    
    # Get software list
    $softwareFilters = $textBoxSoftware.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    
    # Validate input
    if ($computers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter at least one computer name.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Update status
    $labelStatus.Text = "Running query on $($computers.Count) computer(s)..."
    $form.Refresh()
    
    # Get credentials if needed
    $credential = $null
    if ($checkBoxCredentials.Checked) {
        $credential = Get-Credential -Message "Enter credentials for remote computer access"
        if ($null -eq $credential) {
            $labelStatus.Text = "Query canceled"
            return
        }
    }
    
    # Create results collection
    $results = New-Object System.Collections.ArrayList
    
    # Query each computer
    foreach ($computer in $computers) {
        try {
            $labelStatus.Text = "Querying $computer..."
            $form.Refresh()
            
            # Build script block for remote execution
            $scriptBlock = {
                $softwareList = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                    Where-Object { $_.DisplayName -ne $null } | 
                    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
                
                # Also get 32-bit software on 64-bit systems
                if (Test-Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*') {
                    $softwareList += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                        Where-Object { $_.DisplayName -ne $null } | 
                        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
                }
                
                # Return results
                $softwareList | Sort-Object DisplayName -Unique
            }
            
            # Execute the query
            if ($computer -eq $env:COMPUTERNAME) {
                $software = Invoke-Command -ScriptBlock $scriptBlock
            } else {
                if ($null -ne $credential) {
                    $software = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock -Credential $credential -ErrorAction Stop
                } else {
                    $software = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock -ErrorAction Stop
                }
            }
            
            # Filter software if specified
            if ($softwareFilters.Count -gt 0) {
                $filteredSoftware = $software | Where-Object {
                    $item = $_
                    $matched = $false
                    foreach ($filter in $softwareFilters) {
                        if ($item.DisplayName -like "*$filter*") {
                            $matched = $true
                            break
                        }
                    }
                    $matched
                }
                $software = $filteredSoftware
            }
            
            # Add computer name to results
            foreach ($item in $software) {
                $result = [PSCustomObject]@{
                    ComputerName = $computer
                    SoftwareName = $item.DisplayName
                    Version = $item.DisplayVersion
                    Publisher = $item.Publisher
                    InstallDate = $item.InstallDate
                }
                [void]$results.Add($result)
            }
            
        } catch {
            $errorMessage = $_.Exception.Message
            $result = [PSCustomObject]@{
                ComputerName = $computer
                SoftwareName = "ERROR: $errorMessage"
                Version = ""
                Publisher = ""
                InstallDate = ""
            }
            [void]$results.Add($result)
        }
    }
    
    # Export results if needed
    if ($checkBoxExport.Checked -and $results.Count -gt 0) {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveFileDialog.Title = "Save Results"
        $saveFileDialog.FileName = "SoftwareInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $results | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            $labelStatus.Text = "Results exported to $($saveFileDialog.FileName)"
        }
    }
    
    # Display results in grid
    if ($results.Count -gt 0) {
        $dataTable = New-Object System.Data.DataTable
        $dataTable.Columns.Add("Computer Name", [string]) | Out-Null
        $dataTable.Columns.Add("Software Name", [string]) | Out-Null
        $dataTable.Columns.Add("Version", [string]) | Out-Null
        $dataTable.Columns.Add("Publisher", [string]) | Out-Null
        $dataTable.Columns.Add("Install Date", [string]) | Out-Null
        
        foreach ($result in $results) {
            $dataTable.Rows.Add($result.ComputerName, $result.SoftwareName, $result.Version, $result.Publisher, $result.InstallDate) | Out-Null
        }
        
        $dataGridViewResults.DataSource = $dataTable
        $labelStatus.Text = "Query complete. Found $($results.Count) software items."
    } else {
        $labelStatus.Text = "Query complete. No software found matching criteria."
    }
})

# Show the form
[void]$form.ShowDialog()