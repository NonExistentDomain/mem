# MECM Content Downloader and Deployment Tool
# This script provides a GUI to download content from MECM and manage deployments

# Import required modules
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to check if MECM console is installed and import module
function Initialize-MECMConnection {
    $consolePath = $null
    
    # Common installation paths for MECM console
    $possiblePaths = @(
        "C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1",
        "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1",
        "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $consolePath = $path
            break
        }
    }
    
    if ($null -eq $consolePath) {
        [System.Windows.MessageBox]::Show("MECM Console not found. Please install the MECM console or provide the path manually.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
    
    try {
        Import-Module $consolePath -ErrorAction Stop
        
        # Connect to the site
        $siteCode = Read-Host "Enter your MECM site code (e.g., ABC)"
        $providerMachineName = Read-Host "Enter your MECM server name"
        
        if (-not $siteCode -or -not $providerMachineName) {
            [System.Windows.MessageBox]::Show("Site code and server name are required.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
        
        $initParams = @{}
        $initParams.Add("ProviderMachineName", $providerMachineName)
        $initParams.Add("SiteCode", $siteCode)
        
        # Connect to the provider
        New-PSDrive -Name $siteCode -PSProvider CMSite -Root $providerMachineName @initParams -ErrorAction Stop | Out-Null
        Set-Location "$($siteCode):\" -ErrorAction Stop
        
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Failed to connect to MECM: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# Function to get content from an application
function Get-MECMApplicationContent {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    try {
        $application = Get-CMApplication -Name $ApplicationName -Fast
        
        if (-not $application) {
            throw "Application not found: $ApplicationName"
        }
        
        $contentLocations = Get-CMDistributionPoint -Application $application
        
        if (-not $contentLocations -or $contentLocations.Count -eq 0) {
            throw "No distribution points found for application: $ApplicationName"
        }
        
        $dpPath = "\\$($contentLocations[0].NetworkOSPath)\SMS_DP$\sccmcontentlib"
        $contentID = $application.ContentID
        
        # Find the content in the content library
        $pkgLib = Get-ChildItem -Path $dpPath -Recurse -Filter "*.INI" | Where-Object { $_.Name -eq "ContentInfo.ini" } | ForEach-Object {
            $content = Get-Content $_.FullName
            if ($content -match $contentID) {
                $_.Directory.FullName
            }
        }
        
        if (-not $pkgLib) {
            throw "Content not found for application: $ApplicationName"
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Copy content to destination
        Copy-Item -Path "$pkgLib\*" -Destination $DestinationPath -Recurse -Force
        
        return $DestinationPath
    }
    catch {
        throw "Error getting application content: $_"
    }
}

# Function to get software update content
function Get-MECMUpdateContent {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UpdateName,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    try {
        $update = Get-CMSoftwareUpdate -Name $UpdateName -Fast
        
        if (-not $update) {
            throw "Software update not found: $UpdateName"
        }
        
        $contentLocations = Get-CMDistributionPoint -SoftwareUpdate $update
        
        if (-not $contentLocations -or $contentLocations.Count -eq 0) {
            throw "No distribution points found for update: $UpdateName"
        }
        
        $dpPath = "\\$($contentLocations[0].NetworkOSPath)\SMS_DP$\sccmcontentlib"
        $contentID = $update.ContentID
        
        # Find the content in the content library
        $updateLib = Get-ChildItem -Path $dpPath -Recurse -Filter "*.INI" | Where-Object { $_.Name -eq "ContentInfo.ini" } | ForEach-Object {
            $content = Get-Content $_.FullName
            if ($content -match $contentID) {
                $_.Directory.FullName
            }
        }
        
        if (-not $updateLib) {
            throw "Content not found for update: $UpdateName"
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Copy content to destination
        Copy-Item -Path "$updateLib\*" -Destination $DestinationPath -Recurse -Force
        
        return $DestinationPath
    }
    catch {
        throw "Error getting software update content: $_"
    }
}

# Function to get package content
function Get-MECMPackageContent {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PackageName,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    try {
        $package = Get-CMPackage -Name $PackageName -Fast
        
        if (-not $package) {
            throw "Package not found: $PackageName"
        }
        
        $contentLocations = Get-CMDistributionPoint -Package $package
        
        if (-not $contentLocations -or $contentLocations.Count -eq 0) {
            throw "No distribution points found for package: $PackageName"
        }
        
        $dpPath = "\\$($contentLocations[0].NetworkOSPath)\SMS_DP$\sccmcontentlib"
        $contentID = $package.PackageID
        
        # Find the content in the content library
        $pkgLib = Get-ChildItem -Path $dpPath -Recurse -Filter "*.INI" | Where-Object { $_.Name -eq "ContentInfo.ini" } | ForEach-Object {
            $content = Get-Content $_.FullName
            if ($content -match $contentID) {
                $_.Directory.FullName
            }
        }
        
        if (-not $pkgLib) {
            throw "Content not found for package: $PackageName"
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Copy content to destination
        Copy-Item -Path "$pkgLib\*" -Destination $DestinationPath -Recurse -Force
        
        return $DestinationPath
    }
    catch {
        throw "Error getting package content: $_"
    }
}

# Function to get task sequence content
function Get-MECMTaskSequenceContent {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TaskSequenceName,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    try {
        $taskSequence = Get-CMTaskSequence -Name $TaskSequenceName
        
        if (-not $taskSequence) {
            throw "Task sequence not found: $TaskSequenceName"
        }
        
        $contentLocations = Get-CMDistributionPoint -TaskSequence $taskSequence
        
        if (-not $contentLocations -or $contentLocations.Count -eq 0) {
            throw "No distribution points found for task sequence: $TaskSequenceName"
        }
        
        $dpPath = "\\$($contentLocations[0].NetworkOSPath)\SMS_DP$\sccmcontentlib"
        $contentID = $taskSequence.PackageID
        
        # Find the content in the content library
        $tsLib = Get-ChildItem -Path $dpPath -Recurse -Filter "*.INI" | Where-Object { $_.Name -eq "ContentInfo.ini" } | ForEach-Object {
            $content = Get-Content $_.FullName
            if ($content -match $contentID) {
                $_.Directory.FullName
            }
        }
        
        if (-not $tsLib) {
            throw "Content not found for task sequence: $TaskSequenceName"
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Copy content to destination
        Copy-Item -Path "$tsLib\*" -Destination $DestinationPath -Recurse -Force
        
        return $DestinationPath
    }
    catch {
        throw "Error getting task sequence content: $_"
    }
}

# Function to install an application on a remote computer
function Install-MECMApplication {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName
    )
    
    try {
        $application = Get-CMApplication -Name $ApplicationName -Fast
        
        if (-not $application) {
            throw "Application not found: $ApplicationName"
        }
        
        # Trigger application installation
        Invoke-CMClientAction -ComputerName $ComputerName -ActionType ApplicationDeploymentEvaluation
        Start-Process "C:\Windows\CCM\SCClient.exe" -ArgumentList "/{00000000-0000-0000-0000-000000000021} {$($application.ModelName)}"
        
        return $true
    }
    catch {
        throw "Error installing application: $_"
    }
}

# Function to uninstall an application on a remote computer
function Uninstall-MECMApplication {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName
    )
    
    try {
        $application = Get-CMApplication -Name $ApplicationName -Fast
        
        if (-not $application) {
            throw "Application not found: $ApplicationName"
        }
        
        # Trigger application uninstallation
        Invoke-CMClientAction -ComputerName $ComputerName -ActionType ApplicationDeploymentEvaluation
        Start-Process "C:\Windows\CCM\SCClient.exe" -ArgumentList "/{00000000-0000-0000-0000-000000000022} {$($application.ModelName)}"
        
        return $true
    }
    catch {
        throw "Error uninstalling application: $_"
    }
}

# Function to install a software update on a remote computer
function Install-MECMUpdate {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $true)]
        [string]$UpdateName
    )
    
    try {
        $update = Get-CMSoftwareUpdate -Name $UpdateName -Fast
        
        if (-not $update) {
            throw "Software update not found: $UpdateName"
        }
        
        # Trigger update installation
        Invoke-CMClientAction -ComputerName $ComputerName -ActionType SoftwareUpdateScan
        Invoke-CMClientAction -ComputerName $ComputerName -ActionType SoftwareUpdateDeploymentEvaluation
        
        return $true
    }
    catch {
        throw "Error installing update: $_"
    }
}

# Function to run a task sequence on a remote computer
function Start-MECMTaskSequence {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $true)]
        [string]$TaskSequenceName
    )
    
    try {
        $taskSequence = Get-CMTaskSequence -Name $TaskSequenceName
        
        if (-not $taskSequence) {
            throw "Task sequence not found: $TaskSequenceName"
        }
        
        # Trigger task sequence
        Invoke-CMClientAction -ComputerName $ComputerName -ActionType TaskSequenceEvaluation
        
        return $true
    }
    catch {
        throw "Error running task sequence: $_"
    }
}

# Function to create GUI
function Show-MECMGUI {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "MECM Content Downloader and Deployment Tool"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    
    # Create tab control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(765, 540)
    
    # Create Applications tab
    $tabApplications = New-Object System.Windows.Forms.TabPage
    $tabApplications.Text = "Applications"
    
    # Create updates tab
    $tabUpdates = New-Object System.Windows.Forms.TabPage
    $tabUpdates.Text = "Software Updates"
    
    # Create packages tab
    $tabPackages = New-Object System.Windows.Forms.TabPage
    $tabPackages.Text = "Packages"
    
    # Create task sequences tab
    $tabTaskSequences = New-Object System.Windows.Forms.TabPage
    $tabTaskSequences.Text = "Task Sequences"
    
    # Add tabs to tab control
    $tabControl.Controls.Add($tabApplications)
    $tabControl.Controls.Add($tabUpdates)
    $tabControl.Controls.Add($tabPackages)
    $tabControl.Controls.Add($tabTaskSequences)
    
    # Add tab control to form
    $form.Controls.Add($tabControl)
    
    # Create controls for Applications tab
    $lblAppSearch = New-Object System.Windows.Forms.Label
    $lblAppSearch.Text = "Search Applications:"
    $lblAppSearch.Location = New-Object System.Drawing.Point(10, 10)
    $lblAppSearch.Size = New-Object System.Drawing.Size(120, 20)
    $tabApplications.Controls.Add($lblAppSearch)
    
    $txtAppSearch = New-Object System.Windows.Forms.TextBox
    $txtAppSearch.Location = New-Object System.Drawing.Point(130, 10)
    $txtAppSearch.Size = New-Object System.Drawing.Size(250, 20)
    $tabApplications.Controls.Add($txtAppSearch)
    
    $btnAppSearch = New-Object System.Windows.Forms.Button
    $btnAppSearch.Text = "Search"
    $btnAppSearch.Location = New-Object System.Drawing.Point(390, 10)
    $btnAppSearch.Size = New-Object System.Drawing.Size(80, 20)
    $tabApplications.Controls.Add($btnAppSearch)
    
    $lstApplications = New-Object System.Windows.Forms.ListView
    $lstApplications.View = "Details"
    $lstApplications.Location = New-Object System.Drawing.Point(10, 40)
    $lstApplications.Size = New-Object System.Drawing.Size(730, 250)
    $lstApplications.FullRowSelect = $true
    $lstApplications.Columns.Add("Name", 300)
    $lstApplications.Columns.Add("Version", 100)
    $lstApplications.Columns.Add("Publisher", 150)
    $lstApplications.Columns.Add("ContentID", 180)
    $tabApplications.Controls.Add($lstApplications)
    
    $lblAppTarget = New-Object System.Windows.Forms.Label
    $lblAppTarget.Text = "Target Computer(s):"
    $lblAppTarget.Location = New-Object System.Drawing.Point(10, 300)
    $lblAppTarget.Size = New-Object System.Drawing.Size(120, 20)
    $tabApplications.Controls.Add($lblAppTarget)
    
    $txtAppTarget = New-Object System.Windows.Forms.TextBox
    $txtAppTarget.Location = New-Object System.Drawing.Point(130, 300)
    $txtAppTarget.Size = New-Object System.Drawing.Size(250, 20)
    $txtAppTarget.Text = $env:COMPUTERNAME
    $tabApplications.Controls.Add($txtAppTarget)
    
    $lblAppTargetInfo = New-Object System.Windows.Forms.Label
    $lblAppTargetInfo.Text = "(Separate multiple computers with commas)"
    $lblAppTargetInfo.Location = New-Object System.Drawing.Point(390, 300)
    $lblAppTargetInfo.Size = New-Object System.Drawing.Size(250, 20)
    $tabApplications.Controls.Add($lblAppTargetInfo)
    
    $lblAppDestination = New-Object System.Windows.Forms.Label
    $lblAppDestination.Text = "Destination Path:"
    $lblAppDestination.Location = New-Object System.Drawing.Point(10, 330)
    $lblAppDestination.Size = New-Object System.Drawing.Size(120, 20)
    $tabApplications.Controls.Add($lblAppDestination)
    
    $txtAppDestination = New-Object System.Windows.Forms.TextBox
    $txtAppDestination.Location = New-Object System.Drawing.Point(130, 330)
    $txtAppDestination.Size = New-Object System.Drawing.Size(250, 20)
    $txtAppDestination.Text = "C:\MECM_Downloads\Applications"
    $tabApplications.Controls.Add($txtAppDestination)
    
    $btnAppDownload = New-Object System.Windows.Forms.Button
    $btnAppDownload.Text = "Download Content"
    $btnAppDownload.Location = New-Object System.Drawing.Point(10, 370)
    $btnAppDownload.Size = New-Object System.Drawing.Size(150, 30)
    $tabApplications.Controls.Add($btnAppDownload)
    
    $btnAppInstall = New-Object System.Windows.Forms.Button
    $btnAppInstall.Text = "Install Application"
    $btnAppInstall.Location = New-Object System.Drawing.Point(170, 370)
    $btnAppInstall.Size = New-Object System.Drawing.Size(150, 30)
    $tabApplications.Controls.Add($btnAppInstall)
    
    $btnAppUninstall = New-Object System.Windows.Forms.Button
    $btnAppUninstall.Text = "Uninstall Application"
    $btnAppUninstall.Location = New-Object System.Drawing.Point(330, 370)
    $btnAppUninstall.Size = New-Object System.Drawing.Size(150, 30)
    $tabApplications.Controls.Add($btnAppUninstall)
    
    $txtAppLog = New-Object System.Windows.Forms.TextBox
    $txtAppLog.Location = New-Object System.Drawing.Point(10, 410)
    $txtAppLog.Size = New-Object System.Drawing.Size(730, 90)
    $txtAppLog.Multiline = $true
    $txtAppLog.ScrollBars = "Vertical"
    $txtAppLog.ReadOnly = $true
    $tabApplications.Controls.Add($txtAppLog)
    
    # Create controls for Software Updates tab
    $lblUpdateSearch = New-Object System.Windows.Forms.Label
    $lblUpdateSearch.Text = "Search Updates:"
    $lblUpdateSearch.Location = New-Object System.Drawing.Point(10, 10)
    $lblUpdateSearch.Size = New-Object System.Drawing.Size(120, 20)
    $tabUpdates.Controls.Add($lblUpdateSearch)
    
    $txtUpdateSearch = New-Object System.Windows.Forms.TextBox
    $txtUpdateSearch.Location = New-Object System.Drawing.Point(130, 10)
    $txtUpdateSearch.Size = New-Object System.Drawing.Size(250, 20)
    $tabUpdates.Controls.Add($txtUpdateSearch)
    
    $btnUpdateSearch = New-Object System.Windows.Forms.Button
    $btnUpdateSearch.Text = "Search"
    $btnUpdateSearch.Location = New-Object System.Drawing.Point(390, 10)
    $btnUpdateSearch.Size = New-Object System.Drawing.Size(80, 20)
    $tabUpdates.Controls.Add($btnUpdateSearch)
    
    $lstUpdates = New-Object System.Windows.Forms.ListView
    $lstUpdates.View = "Details"
    $lstUpdates.Location = New-Object System.Drawing.Point(10, 40)
    $lstUpdates.Size = New-Object System.Drawing.Size(730, 250)
    $lstUpdates.FullRowSelect = $true
    $lstUpdates.Columns.Add("Name", 350)
    $lstUpdates.Columns.Add("Article ID", 100)
    $lstUpdates.Columns.Add("Type", 100)
    $lstUpdates.Columns.Add("ContentID", 180)
    $tabUpdates.Controls.Add($lstUpdates)
    
    $lblUpdateTarget = New-Object System.Windows.Forms.Label
    $lblUpdateTarget.Text = "Target Computer(s):"
    $lblUpdateTarget.Location = New-Object System.Drawing.Point(10, 300)
    $lblUpdateTarget.Size = New-Object System.Drawing.Size(120, 20)
    $tabUpdates.Controls.Add($lblUpdateTarget)
    
    $txtUpdateTarget = New-Object System.Windows.Forms.TextBox
    $txtUpdateTarget.Location = New-Object System.Drawing.Point(130, 300)
    $txtUpdateTarget.Size = New-Object System.Drawing.Size(250, 20)
    $txtUpdateTarget.Text = $env:COMPUTERNAME
    $tabUpdates.Controls.Add($txtUpdateTarget)
    
    $lblUpdateTargetInfo = New-Object System.Windows.Forms.Label
    $lblUpdateTargetInfo.Text = "(Separate multiple computers with commas)"
    $lblUpdateTargetInfo.Location = New-Object System.Drawing.Point(390, 300)
    $lblUpdateTargetInfo.Size = New-Object System.Drawing.Size(250, 20)
    $tabUpdates.Controls.Add($lblUpdateTargetInfo)
    
    $lblUpdateDestination = New-Object System.Windows.Forms.Label
    $lblUpdateDestination.Text = "Destination Path:"
    $lblUpdateDestination.Location = New-Object System.Drawing.Point(10, 330)
    $lblUpdateDestination.Size = New-Object System.Drawing.Size(120, 20)
    $tabUpdates.Controls.Add($lblUpdateDestination)
    
    $txtUpdateDestination = New-Object System.Windows.Forms.TextBox
    $txtUpdateDestination.Location = New-Object System.Drawing.Point(130, 330)
    $txtUpdateDestination.Size = New-Object System.Drawing.Size(250, 20)
    $txtUpdateDestination.Text = "C:\MECM_Downloads\Updates"
    $tabUpdates.Controls.Add($txtUpdateDestination)
    
    $btnUpdateDownload = New-Object System.Windows.Forms.Button
    $btnUpdateDownload.Text = "Download Content"
    $btnUpdateDownload.Location = New-Object System.Drawing.Point(10, 370)
    $btnUpdateDownload.Size = New-Object System.Drawing.Size(150, 30)
    $tabUpdates.Controls.Add($btnUpdateDownload)
    
    $btnUpdateInstall = New-Object System.Windows.Forms.Button
    $btnUpdateInstall.Text = "Install Update"
    $btnUpdateInstall.Location = New-Object System.Drawing.Point(170, 370)
    $btnUpdateInstall.Size = New-Object System.Drawing.Size(150, 30)
    $tabUpdates.Controls.Add($btnUpdateInstall)
    
    $txtUpdateLog = New-Object System.Windows.Forms.TextBox
    $txtUpdateLog.Location = New-Object System.Drawing.Point(10, 410)
    $txtUpdateLog.Size = New-Object System.Drawing.Size(730, 90)
    $txtUpdateLog.Multiline = $true
    $txtUpdateLog.ScrollBars = "Vertical"
    $txtUpdateLog.ReadOnly = $true
    $tabUpdates.Controls.Add($txtUpdateLog)
    
    # Create controls for Packages tab
    $lblPkgSearch = New-Object System.Windows.Forms.Label
    $lblPkgSearch.Text = "Search Packages:"
    $lblPkgSearch.Location = New-Object System.Drawing.Point(10, 10)
    $lblPkgSearch.Size = New-Object System.Drawing.Size(120, 20)
    $tabPackages.Controls.Add($lblPkgSearch)
    
    $txtPkgSearch = New-Object System.Windows.Forms.TextBox
    $txtPkgSearch.Location = New-Object System.Drawing.Point(130, 10)
    $txtPkgSearch.Size = New-Object System.Drawing.Size(250, 20)
    $tabPackages.Controls.Add($txtPkgSearch)
    
    $btnPkgSearch = New-Object System.Windows.Forms.Button
    $btnPkgSearch.Text = "Search"
    $btnPkgSearch.Location = New-Object System.Drawing.Point(390, 10)
    $btnPkgSearch.Size = New-Object System.Drawing.Size(80, 20)
    $tabPackages.Controls.Add($btnPkgSearch)
    
    $lstPackages = New-Object System.Windows.Forms.ListView
    $lstPackages.View = "Details"
    $lstPackages.Location = New-Object System.Drawing.Point(10, 40)
    $lstPackages.Size = New-Object System.Drawing.Size(730, 250)
    $lstPackages.FullRowSelect = $true
    $lstPackages.Columns.Add("Name", 300)
    $lstPackages.Columns.Add("Version", 100)
    $lstPackages.Columns.Add("Package ID", 150)
    $lstPackages.Columns.Add("Description", 180)
    $tabPackages.Controls.Add($lstPackages)
    
    $lblPkgDestination = New-Object System.Windows.Forms.Label
    $lblPkgDestination.Text = "Destination Path:"
    $lblPkgDestination.Location = New-Object System.Drawing.Point(10, 330)
    $lblPkgDestination.Size = New-Object System.Drawing.Size(120, 20)
    $tabPackages.Controls.Add($lblPkgDestination)
    
    $txtPkgDestination = New-Object System.Windows.Forms.TextBox
    $txtPkgDestination.Location = New-Object System.Drawing.Point(130, 330)
    $txtPkgDestination.Size = New-Object System.Drawing.Size(250, 20)
    $txtPkgDestination.Text = "C:\MECM_Downloads\Packages"
    $tabPackages.Controls.Add($txtPkgDestination)
    
    $btnPkgDownload = New-Object System.Windows.Forms.Button
    $btnPkgDownload.Text = "Download Content"
    $btnPkgDownload.Location = New-Object System.Drawing.Point(10, 370)
    $btnPkgDownload.Size = New-Object System.Drawing.Size(150, 30)
    $tabPackages.Controls.Add($btnPkgDownload)
    
    $txtPkgLog = New-Object System.Windows.Forms.TextBox
    $txtPkgLog.Location = New-Object System.Drawing.Point(10, 410)
    $txtPkgLog.Size = New-Object System.Drawing.Size(730, 90)
    $txtPkgLog.Multiline = $true
    $txtPkgLog.ScrollBars = "Vertical"
    $txtPkgLog.ReadOnly = $true
    $tabPackages.Controls.Add($txtPkgLog)
    
    # Create controls for Task Sequences tab
    $lblTSSearch = New-Object System.Windows.Forms.Label
    $lblTSSearch.Text = "Search Task Sequences:"
    $lblTSSearch.Location = New-Object System.Drawing.Point(10, 10)
    $lblTSSearch.Size = New-Object System.Drawing.Size(120, 20)
    $tabTaskSequences.Controls.Add($lblTSSearch)
    
    $txtTSSearch = New-Object System.Windows.Forms.TextBox
    $txtTSSearch.Location = New-Object System.Drawing.Point(130, 10)
    $txtTSSearch.Size = New-Object System.Drawing.Size(250, 20)
    $tabTaskSequences.Controls.Add($txtTSSearch)
    
    $btnTSSearch = New-Object System.Windows.Forms.Button
    $btnTSSearch.Text = "Search"
    $btnTSSearch.Location = New-Object System.Drawing.Point(390, 10)
    $btnTSSearch.Size = New-Object System.Drawing.Size(80, 20)
    $tabTaskSequences.Controls.Add($btnTSSearch)
    
    $lstTaskSequences = New-Object System.Windows.Forms.ListView
    $lstTaskSequences.View = "Details"
    $lstTaskSequences.Location = New-Object System.Drawing.Point(10, 40)
    $lstTaskSequences.Size = New-Object System.Drawing.Size(730, 250)
    $lstTaskSequences.FullRowSelect = $true
    $lstTaskSequences.Columns.Add("Name", 300)
    $lstTaskSequences.Columns.Add("Description", 250)
    $lstTaskSequences.Columns.Add("Package ID", 180)
    $tabTaskSequences.Controls.Add($lstTaskSequences)
    
    $lblTSTarget = New-Object System.Windows.Forms.Label
    $lblTSTarget.Text = "Target Computer(s):"
    $lblTSTarget = New-Object System.Windows.Forms.Label
    $lblTSTarget.Text = "Target Computer(s):"
    $lblTSTarget.Location = New-Object System.Drawing.Point(10, 300)
    $lblTSTarget.Size = New-Object System.Drawing.Size(120, 20)
    $tabTaskSequences.Controls.Add($lblTSTarget)
    
    $txtTSTarget = New-Object System.Windows.Forms.TextBox
    $txtTSTarget.Location = New-Object System.Drawing.Point(130, 300)
    $txtTSTarget.Size = New-Object System.Drawing.Size(250, 20)
    $txtTSTarget.Text = $env:COMPUTERNAME
    $tabTaskSequences.Controls.Add($txtTSTarget)
    
    $lblTSTargetInfo = New-Object System.Windows.Forms.Label
    $lblTSTargetInfo.Text = "(Separate multiple computers with commas)"
    $lblTSTargetInfo.Location = New-Object System.Drawing.Point(390, 300)
    $lblTSTargetInfo.Size = New-Object System.Drawing.Size(250, 20)
    $tabTaskSequences.Controls.Add($lblTSTargetInfo)
    
    $lblTSDestination = New-Object System.Windows.Forms.Label
    $lblTSDestination.Text = "Destination Path:"
    $lblTSDestination.Location = New-Object System.Drawing.Point(10, 330)
    $lblTSDestination.Size = New-Object System.Drawing.Size(120, 20)
    $tabTaskSequences.Controls.Add($lblTSDestination)
    
    $txtTSDestination = New-Object System.Windows.Forms.TextBox
    $txtTSDestination.Location = New-Object System.Drawing.Point(130, 330)
    $txtTSDestination.Size = New-Object System.Drawing.Size(250, 20)
    $txtTSDestination.Text = "C:\MECM_Downloads\TaskSequences"
    $tabTaskSequences.Controls.Add($txtTSDestination)
    
    $btnTSDownload = New-Object System.Windows.Forms.Button
    $btnTSDownload.Text = "Download Content"
    $btnTSDownload.Location = New-Object System.Drawing.Point(10, 370)
    $btnTSDownload.Size = New-Object System.Drawing.Size(150, 30)
    $tabTaskSequences.Controls.Add($btnTSDownload)
    
    $btnTSRun = New-Object System.Windows.Forms.Button
    $btnTSRun.Text = "Run Task Sequence"
    $btnTSRun.Location = New-Object System.Drawing.Point(170, 370)
    $btnTSRun.Size = New-Object System.Drawing.Size(150, 30)
    $tabTaskSequences.Controls.Add($btnTSRun)
    
    $txtTSLog = New-Object System.Windows.Forms.TextBox
    $txtTSLog.Location = New-Object System.Drawing.Point(10, 410)
    $txtTSLog.Size = New-Object System.Drawing.Size(730, 90)
    $txtTSLog.Multiline = $true
    $txtTSLog.ScrollBars = "Vertical"
    $txtTSLog.ReadOnly = $true
    $tabTaskSequences.Controls.Add($txtTSLog)
    
    # Event handlers for Applications tab
    $btnAppSearch.Add_Click({
        $txtAppLog.Text = "Searching for applications..."
        $lstApplications.Items.Clear()
        
        try {
            $apps = Get-CMApplication -Name "*$($txtAppSearch.Text)*" -Fast
            
            if ($apps) {
                foreach ($app in $apps) {
                    $item = New-Object System.Windows.Forms.ListViewItem($app.LocalizedDisplayName)
                    $item.SubItems.Add($app.SoftwareVersion)
                    $item.SubItems.Add($app.Manufacturer)
                    $item.SubItems.Add($app.ContentID)
                    $lstApplications.Items.Add($item)
                }
                $txtAppLog.Text = "Found $($apps.Count) applications."
            }
            else {
                $txtAppLog.Text = "No applications found."
            }
        }
        catch {
            $txtAppLog.Text = "Error searching for applications: $_"
        }
    })
    
    $btnAppDownload.Add_Click({
        if ($lstApplications.SelectedItems.Count -gt 0) {
            $appName = $lstApplications.SelectedItems[0].Text
            
            $txtAppLog.Text = "Downloading content for application: $appName to $($txtAppDestination.Text)"
            
            try {
                $path = Get-MECMApplicationContent -ApplicationName $appName -DestinationPath $txtAppDestination.Text
                $txtAppLog.Text = "Downloaded content for application: $appName to $path"
            }
            catch {
                $txtAppLog.Text = "Error downloading application content: $_"
            }
        }
        else {
            $txtAppLog.Text = "Please select an application."
        }
    })
    
    $btnAppInstall.Add_Click({
        if ($lstApplications.SelectedItems.Count -gt 0) {
            $appName = $lstApplications.SelectedItems[0].Text
            $computers = $txtAppTarget.Text -split ','
            
            $txtAppLog.Text = "Installing application: $appName on $($computers.Count) computer(s)..."
            
            foreach ($computer in $computers) {
                $computer = $computer.Trim()
                try {
                    $result = Install-MECMApplication -ComputerName $computer -ApplicationName $appName
                    $txtAppLog.AppendText("`r`nInitiated installation of $appName on $computer.")
                }
                catch {
                    $txtAppLog.AppendText("`r`nError installing $appName on ${computer}: $_")
                }
            }
        }
        else {
            $txtAppLog.Text = "Please select an application."
        }
    })
    
    $btnAppUninstall.Add_Click({
        if ($lstApplications.SelectedItems.Count -gt 0) {
            $appName = $lstApplications.SelectedItems[0].Text
            $computers = $txtAppTarget.Text -split ','
            
            $txtAppLog.Text = "Uninstalling application: $appName from $($computers.Count) computer(s)..."
            
            foreach ($computer in $computers) {
                $computer = $computer.Trim()
                try {
                    $result = Uninstall-MECMApplication -ComputerName $computer -ApplicationName $appName
                    $txtAppLog.AppendText("`r`nInitiated uninstallation of $appName from $computer.")
                }
                catch {
                    $txtAppLog.AppendText("`r`nError uninstalling $appName from ${computer}: $_")
                }
            }
        }
        else {
            $txtAppLog.Text = "Please select an application."
        }
    })
    
    # Event handlers for Software Updates tab
    $btnUpdateSearch.Add_Click({
        $txtUpdateLog.Text = "Searching for updates..."
        $lstUpdates.Items.Clear()
        
        try {
            $updates = Get-CMSoftwareUpdate -Name "*$($txtUpdateSearch.Text)*" -Fast
            
            if ($updates) {
                foreach ($update in $updates) {
                    $item = New-Object System.Windows.Forms.ListViewItem($update.LocalizedDisplayName)
                    $item.SubItems.Add($update.ArticleID)
                    $item.SubItems.Add($update.UpdateClassification)
                    $item.SubItems.Add($update.CI_ID)
                    $lstUpdates.Items.Add($item)
                }
                $txtUpdateLog.Text = "Found $($updates.Count) updates."
            }
            else {
                $txtUpdateLog.Text = "No updates found."
            }
        }
        catch {
            $txtUpdateLog.Text = "Error searching for updates: $_"
        }
    })
    
    $btnUpdateDownload.Add_Click({
        if ($lstUpdates.SelectedItems.Count -gt 0) {
            $updateName = $lstUpdates.SelectedItems[0].Text
            
            $txtUpdateLog.Text = "Downloading content for update: $updateName to $($txtUpdateDestination.Text)"
            
            try {
                $path = Get-MECMUpdateContent -UpdateName $updateName -DestinationPath $txtUpdateDestination.Text
                $txtUpdateLog.Text = "Downloaded content for update: $updateName to $path"
            }
            catch {
                $txtUpdateLog.Text = "Error downloading update content: $_"
            }
        }
        else {
            $txtUpdateLog.Text = "Please select an update."
        }
    })
    
    $btnUpdateInstall.Add_Click({
        if ($lstUpdates.SelectedItems.Count -gt 0) {
            $updateName = $lstUpdates.SelectedItems[0].Text
            $computers = $txtUpdateTarget.Text -split ','
            
            $txtUpdateLog.Text = "Installing update: $updateName on $($computers.Count) computer(s)..."
            
            foreach ($computer in $computers) {
                $computer = $computer.Trim()
                try {
                    $result = Install-MECMUpdate -ComputerName $computer -UpdateName $updateName
                    $txtUpdateLog.AppendText("`r`nInitiated installation of $updateName on $computer.")
                }
                catch {
                    $txtUpdateLog.AppendText("`r`nError installing $updateName on ${$computer}: $_")
                }
            }
        }
        else {
            $txtUpdateLog.Text = "Please select an update."
        }
    })
    
    # Event handlers for Packages tab
    $btnPkgSearch.Add_Click({
        $txtPkgLog.Text = "Searching for packages..."
        $lstPackages.Items.Clear()
        
        try {
            $packages = Get-CMPackage -Name "*$($txtPkgSearch.Text)*" -Fast
            
            if ($packages) {
                foreach ($package in $packages) {
                    $item = New-Object System.Windows.Forms.ListViewItem($package.Name)
                    $item.SubItems.Add($package.Version)
                    $item.SubItems.Add($package.PackageID)
                    $item.SubItems.Add($package.Description)
                    $lstPackages.Items.Add($item)
                }
                $txtPkgLog.Text = "Found $($packages.Count) packages."
            }
            else {
                $txtPkgLog.Text = "No packages found."
            }
        }
        catch {
            $txtPkgLog.Text = "Error searching for packages: $_"
        }
    })
    
    $btnPkgDownload.Add_Click({
        if ($lstPackages.SelectedItems.Count -gt 0) {
            $packageName = $lstPackages.SelectedItems[0].Text
            
            $txtPkgLog.Text = "Downloading content for package: $packageName to $($txtPkgDestination.Text)"
            
            try {
                $path = Get-MECMPackageContent -PackageName $packageName -DestinationPath $txtPkgDestination.Text
                $txtPkgLog.Text = "Downloaded content for package: $packageName to $path"
            }
            catch {
                $txtPkgLog.Text = "Error downloading package content: $_"
            }
        }
        else {
            $txtPkgLog.Text = "Please select a package."
        }
    })
    
    # Event handlers for Task Sequences tab
    $btnTSSearch.Add_Click({
        $txtTSLog.Text = "Searching for task sequences..."
        $lstTaskSequences.Items.Clear()
        
        try {
            $taskSequences = Get-CMTaskSequence -Name "*$($txtTSSearch.Text)*"
            
            if ($taskSequences) {
                foreach ($ts in $taskSequences) {
                    $item = New-Object System.Windows.Forms.ListViewItem($ts.Name)
                    $item.SubItems.Add($ts.Description)
                    $item.SubItems.Add($ts.PackageID)
                    $lstTaskSequences.Items.Add($item)
                }
                $txtTSLog.Text = "Found $($taskSequences.Count) task sequences."
            }
            else {
                $txtTSLog.Text = "No task sequences found."
            }
        }
        catch {
            $txtTSLog.Text = "Error searching for task sequences: $_"
        }
    })
    
    $btnTSDownload.Add_Click({
        if ($lstTaskSequences.SelectedItems.Count -gt 0) {
            $tsName = $lstTaskSequences.SelectedItems[0].Text
            
            $txtTSLog.Text = "Downloading content for task sequence: $tsName to $($txtTSDestination.Text)"
            
            try {
                $path = Get-MECMTaskSequenceContent -TaskSequenceName $tsName -DestinationPath $txtTSDestination.Text
                $txtTSLog.Text = "Downloaded content for task sequence: $tsName to $path"
            }
            catch {
                $txtTSLog.Text = "Error downloading task sequence content: $_"
            }
        }
        else {
            $txtTSLog.Text = "Please select a task sequence."
        }
    })
    
    $btnTSRun.Add_Click({
        if ($lstTaskSequences.SelectedItems.Count -gt 0) {
            $tsName = $lstTaskSequences.SelectedItems[0].Text
            $computers = $txtTSTarget.Text -split ','
            
            $txtTSLog.Text = "Running task sequence: $tsName on $($computers.Count) computer(s)..."
            
            foreach ($computer in $computers) {
                $computer = $computer.Trim()
                try {
                    $result = Start-MECMTaskSequence -ComputerName $computer -TaskSequenceName $tsName
                    $txtTSLog.AppendText("`r`nInitiated task sequence $tsName on $computer.")
                }
                catch {
                    $txtTSLog.AppendText("`r`nError running task sequence $tsName on ${computer}: $_")
                }
            }
        }
        else {
            $txtTSLog.Text = "Please select a task sequence."
        }
    })
    
    # Show form
    $form.Add_Shown({
        # Initialize connection to MECM
        if (-not (Initialize-MECMConnection)) {
            $form.Close()
        }
    })
    
    [void]$form.ShowDialog()
}

# Start the GUI
Show-MECMGUI
