
# Requires ActiveDirectory and ConfigurationManager modules
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
$SiteCode = "PS1"  # Adjust to your SCCM site code
$SCCMServer = "SCCM-SERVER.domain.com"  # Adjust to your SCCM server
if (Test-Path "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1") {
    Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
    Set-Location "$SiteCode`:"
}

# Function to verify if a computer is truly offline
function Test-ComputerOfflineStatus {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [int]$ADThresholdMinutes = 60,  # Consider offline if no AD logon in last X minutes
        [int]$MECMThresholdMinutes = 60  # Consider offline if no MECM heartbeat in last X minutes
    )

    # Initialize result object
    $result = [PSCustomObject]@{
        ComputerName    = $ComputerName
        IPAddress       = $null
        Pingable        = $false
        ADLastOnline    = $null
        ADStatus        = "Unknown"
        MECMLastOnline  = $null
        MECMStatus      = "Unknown"
        IdentityVerified = $false
        Conclusion      = "Unknown"
    }

    # Step 1: Resolve IP and check ping (for context, not definitive)
    try {
        $dnsResult = Resolve-DnsName -Name $ComputerName -ErrorAction Stop
        $result.IPAddress = $dnsResult.IPAddress
        $result.Pingable = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "DNS resolution failed for $ComputerName : $_"
    }

    # Step 2: Check AD last logon (authoritative for domain authentication)
    if (Get-Module ActiveDirectory -ErrorAction SilentlyContinue) {
        $adComputer = Get-ADComputer -Filter "Name -eq '$ComputerName'" `
                                    -Properties LastLogonTimeStamp -ErrorAction SilentlyContinue
        if ($adComputer -and $adComputer.LastLogonTimeStamp) {
            $result.ADLastOnline = [DateTime]::FromFileTime($adComputer.LastLogonTimeStamp)
            $adTimeDiff = (Get-Date) - $result.ADLastOnline
            if ($adTimeDiff.TotalMinutes -gt $ADThresholdMinutes) {
                $result.ADStatus = "Offline (No recent AD authentication)"
            } else {
                $result.ADStatus = "Online (Recent AD authentication)"
            }
        } else {
            $result.ADStatus = "No AD data or never logged on"
        }
    } else {
        Write-Warning "ActiveDirectory module not available"
    }

    # Step 3: Check MECM last active time (more precise heartbeat data)
    if (Get-Command Get-CMDevice -ErrorAction SilentlyContinue) {
        $mecmDevice = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue
        if ($mecmDevice -and $mecmDevice.LastActiveTime) {
            $result.MECMLastOnline = $mecmDevice.LastActiveTime
            $mecmTimeDiff = (Get-Date) - $result.MECMLastOnline
            if ($mecmTimeDiff.TotalMinutes -gt $MECMThresholdMinutes) {
                $result.MECMStatus = "Offline (No recent MECM heartbeat)"
            } else {
                $result.MECMStatus = "Online (Recent MECM heartbeat)"
            }
        } else {
            $result.MECMStatus = "No MECM data"
        }
    } else {
        Write-Warning "MECM module not available or not connected to site"
    }

    # Step 4: Direct identity check (optional, if pingable and tools inconclusive)
    if ($result.Pingable -and ($result.ADStatus -eq "Offline" -or $result.MECMStatus -eq "Offline")) {
        try {
            # Attempt WMI query to verify hostname matches
            $wmiResult = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName `
                                      -ErrorAction Stop
            if ($wmiResult.Name -eq $ComputerName) {
                $result.IdentityVerified = $true
                $result.Conclusion = "Online (Identity confirmed via WMI)"
            } else {
                $result.IdentityVerified = $false
                $result.Conclusion = "Offline (Pingable IP belongs to another device)"
            }
        } catch {
            $result.Conclusion = "Offline (No WMI response, likely not the target computer)"
        }
    }

    # Step 5: Final conclusion based on available data
    if (-not $result.Conclusion -or $result.Conclusion -eq "Unknown") {
        if ($result.MECMLastOnline -and $result.MECMStatus -eq "Offline") {
            $result.Conclusion = "Offline (Per MECM heartbeat)"
        } elseif ($result.ADLastOnline -and $result.ADStatus -eq "Offline") {
            $result.Conclusion = "Offline (Per AD last logon)"
        } elseif ($result.Pingable) {
            $result.Conclusion = "Possibly Online (Pingable but identity unverified)"
        } else {
            $result.Conclusion = "Offline (Not pingable and no recent management data)"
        }
    }

    # Output result
    $result
}

# Example usage
Test-ComputerOfflineStatus -ComputerName "WORKSTATION01" -ADThresholdMinutes 60 -MECMThresholdMinutes 60

















# mem

# Requires ActiveDirectory and ConfigurationManager modules
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
$SiteCode = "PS1"  # Adjust to your SCCM site code
$SCCMServer = "SCCM-SERVER.domain.com"  # Adjust to your SCCM server
if (Test-Path "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1") {
    Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
    Set-Location "$SiteCode`:"
}

# Function to verify if a computer is truly offline
function Test-ComputerOfflineStatus {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [int]$ADThresholdMinutes = 60,  # Consider offline if no AD logon in last X minutes
        [int]$MECMThresholdMinutes = 60  # Consider offline if no MECM heartbeat in last X minutes
    )

    # Initialize result object
    $result = [PSCustomObject]@{
        ComputerName    = $ComputerName
        IPAddress       = $null
        Pingable        = $false
        ADLastOnline    = $null
        ADStatus        = "Unknown"
        MECMLastOnline  = $null
        MECMStatus      = "Unknown"
        IdentityVerified = $false
        Conclusion      = "Unknown"
    }

    # Step 1: Resolve IP and check ping (for context, not definitive)
    try {
        $dnsResult = Resolve-DnsName -Name $ComputerName -ErrorAction Stop
        $result.IPAddress = $dnsResult.IPAddress
        $result.Pingable = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "DNS resolution failed for $ComputerName : $_"
    }

    # Step 2: Check AD last logon (authoritative for domain authentication)
    if (Get-Module ActiveDirectory -ErrorAction SilentlyContinue) {
        $adComputer = Get-ADComputer -Filter "Name -eq '$ComputerName'" `
                                    -Properties LastLogonTimeStamp -ErrorAction SilentlyContinue
        if ($adComputer -and $adComputer.LastLogonTimeStamp) {
            $result.ADLastOnline = [DateTime]::FromFileTime($adComputer.LastLogonTimeStamp)
            $adTimeDiff = (Get-Date) - $result.ADLastOnline
            if ($adTimeDiff.TotalMinutes -gt $ADThresholdMinutes) {
                $result.ADStatus = "Offline (No recent AD authentication)"
            } else {
                $result.ADStatus = "Online (Recent AD authentication)"
            }
        } else {
            $result.ADStatus = "No AD data or never logged on"
        }
    } else {
        Write-Warning "ActiveDirectory module not available"
    }

    # Step 3: Check MECM last active time (more precise heartbeat data)
    if (Get-Command Get-CMDevice -ErrorAction SilentlyContinue) {
        $mecmDevice = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue
        if ($mecmDevice -and $mecmDevice.LastActiveTime) {
            $result.MECMLastOnline = $mecmDevice.LastActiveTime
            $mecmTimeDiff = (Get-Date) - $result.MECMLastOnline
            if ($mecmTimeDiff.TotalMinutes -gt $MECMThresholdMinutes) {
                $result.MECMStatus = "Offline (No recent MECM heartbeat)"
            } else {
                $result.MECMStatus = "Online (Recent MECM heartbeat)"
            }
        } else {
            $result.MECMStatus = "No MECM data"
        }
    } else {
        Write-Warning "MECM module not available or not connected to site"
    }

    # Step 4: Direct identity check (optional, if pingable and tools inconclusive)
    if ($result.Pingable -and ($result.ADStatus -eq "Offline" -or $result.MECMStatus -eq "Offline")) {
        try {
            # Attempt WMI query to verify hostname matches
            $wmiResult = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName `
                                      -ErrorAction Stop
            if ($wmiResult.Name -eq $ComputerName) {
                $result.IdentityVerified = $true
                $result.Conclusion = "Online (Identity confirmed via WMI)"
            } else {
                $result.IdentityVerified = $false
                $result.Conclusion = "Offline (Pingable IP belongs to another device)"
            }
        } catch {
            $result.Conclusion = "Offline (No WMI response, likely not the target computer)"
        }
    }

    # Step 5: Final conclusion based on available data
    if (-not $result.Conclusion -or $result.Conclusion -eq "Unknown") {
        if ($result.MECMLastOnline -and $result.MECMStatus -eq "Offline") {
            $result.Conclusion = "Offline (Per MECM heartbeat)"
        } elseif ($result.ADLastOnline -and $result.ADStatus -eq "Offline") {
            $result.Conclusion = "Offline (Per AD last logon)"
        } elseif ($result.Pingable) {
            $result.Conclusion = "Possibly Online (Pingable but identity unverified)"
        } else {
            $result.Conclusion = "Offline (Not pingable and no recent management data)"
        }
    }

    # Output result
    $result
}

# Example usage
Test-ComputerOfflineStatus -ComputerName "WORKSTATION01" -ADThresholdMinutes 60 -MECMThresholdMinutes 60


























Import-Module ActiveDirectory
$SiteCode = "PS1"
$SCCMServer = "SCCM-SERVER.domain.com"
Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
Set-Location "$SiteCode`:"

function Get-ComputerLastOnline {
    param (
        [string]$ComputerName = "WORKSTATION01"
    )
    
    # AD data
    $adData = Get-ADComputer -Filter "Name -eq '$ComputerName'" -Properties LastLogonTimeStamp
    $adLastOnline = if ($adData.LastLogonTimeStamp) { [DateTime]::FromFileTime($adData.LastLogonTimeStamp) } else { "Never" }
    
    # MECM data
    $mecmData = Get-CMDevice -Name $ComputerName
    $mecmLastOnline = $mecmData.LastActiveTime
    
    [PSCustomObject]@{
        ComputerName    = $ComputerName
        ADLastOnline    = $adLastOnline
        MECMLastOnline  = $mecmLastOnline
    }
}

Get-ComputerLastOnline -ComputerName "WORKSTATION01"







#ACTIVE DIRECTORY
# Requires the ActiveDirectory module (install with: Install-Module -Name ActiveDirectory)
Import-Module ActiveDirectory

# Function to get last online time from AD
function Get-ADComputerLastOnline {
    param (
        [string]$ComputerName = "*"  # Wildcard for all computers, or specify a name
    )
    
    # Query AD for computer objects
    $computers = Get-ADComputer -Filter "Name -like '$ComputerName'" `
                               -Properties LastLogonTimeStamp, Name, DistinguishedName
    
    # Process results
    $results = foreach ($computer in $computers) {
        $lastLogon = if ($computer.LastLogonTimeStamp) {
            [DateTime]::FromFileTime($computer.LastLogonTimeStamp)
        } else {
            "Never logged on"
        }
        
        [PSCustomObject]@{
            ComputerName      = $computer.Name
            LastOnline        = $lastLogon
            DistinguishedName = $computer.DistinguishedName
        }
    }
    
    # Output results
    $results | Sort-Object LastOnline -Descending
}

# Example usage
Get-ADComputerLastOnline -ComputerName "WORKSTATION01"  # Single computer
# Get-ADComputerLastOnline  # All computers in the domain




















#MECM AND CONFIG MANAGER
# Requires ActiveDirectory and ConfigurationManager modules
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
$SiteCode = "PS1"  # Adjust to your SCCM site code
$SCCMServer = "SCCM-SERVER.domain.com"  # Adjust to your SCCM server
if (Test-Path "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1") {
    Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
    Set-Location "$SiteCode`:"
}

# Function to verify if a computer is truly offline
function Test-ComputerOfflineStatus {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [int]$ADThresholdMinutes = 60,  # Consider offline if no AD logon in last X minutes
        [int]$MECMThresholdMinutes = 60  # Consider offline if no MECM heartbeat in last X minutes
    )

    # Initialize result object
    $result = [PSCustomObject]@{
        ComputerName    = $ComputerName
        IPAddress       = $null
        Pingable        = $false
        ADLastOnline    = $null
        ADStatus        = "Unknown"
        MECMLastOnline  = $null
        MECMStatus      = "Unknown"
        IdentityVerified = $false
        Conclusion      = "Unknown"
    }

    # Step 1: Resolve IP and check ping (for context, not definitive)
    try {
        $dnsResult = Resolve-DnsName -Name $ComputerName -ErrorAction Stop
        $result.IPAddress = $dnsResult.IPAddress
        $result.Pingable = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "DNS resolution failed for $ComputerName : $_"
    }

    # Step 2: Check AD last logon (authoritative for domain authentication)
    if (Get-Module ActiveDirectory -ErrorAction SilentlyContinue) {
        $adComputer = Get-ADComputer -Filter "Name -eq '$ComputerName'" `
                                    -Properties LastLogonTimeStamp -ErrorAction SilentlyContinue
        if ($adComputer -and $adComputer.LastLogonTimeStamp) {
            $result.ADLastOnline = [DateTime]::FromFileTime($adComputer.LastLogonTimeStamp)
            $adTimeDiff = (Get-Date) - $result.ADLastOnline
            if ($adTimeDiff.TotalMinutes -gt $ADThresholdMinutes) {
                $result.ADStatus = "Offline (No recent AD authentication)"
            } else {
                $result.ADStatus = "Online (Recent AD authentication)"
            }
        } else {
            $result.ADStatus = "No AD data or never logged on"
        }
    } else {
        Write-Warning "ActiveDirectory module not available"
    }

    # Step 3: Check MECM last active time (more precise heartbeat data)
    if (Get-Command Get-CMDevice -ErrorAction SilentlyContinue) {
        $mecmDevice = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue
        if ($mecmDevice -and $mecmDevice.LastActiveTime) {
            $result.MECMLastOnline = $mecmDevice.LastActiveTime
            $mecmTimeDiff = (Get-Date) - $result.MECMLastOnline
            if ($mecmTimeDiff.TotalMinutes -gt $MECMThresholdMinutes) {
                $result.MECMStatus = "Offline (No recent MECM heartbeat)"
            } else {
                $result.MECMStatus = "Online (Recent MECM heartbeat)"
            }
        } else {
            $result.MECMStatus = "No MECM data"
        }
    } else {
        Write-Warning "MECM module not available or not connected to site"
    }

    # Step 4: Direct identity check (optional, if pingable and tools inconclusive)
    if ($result.Pingable -and ($result.ADStatus -eq "Offline" -or $result.MECMStatus -eq "Offline")) {
        try {
            # Attempt WMI query to verify hostname matches
            $wmiResult = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName `
                                      -ErrorAction Stop
            if ($wmiResult.Name -eq $ComputerName) {
                $result.IdentityVerified = $true
                $result.Conclusion = "Online (Identity confirmed via WMI)"
            } else {
                $result.IdentityVerified = $false
                $result.Conclusion = "Offline (Pingable IP belongs to another device)"
            }
        } catch {
            $result.Conclusion = "Offline (No WMI response, likely not the target computer)"
        }
    }

    # Step 5: Final conclusion based on available data
    if (-not $result.Conclusion -or $result.Conclusion -eq "Unknown") {
        if ($result.MECMLastOnline -and $result.MECMStatus -eq "Offline") {
            $result.Conclusion = "Offline (Per MECM heartbeat)"
        } elseif ($result.ADLastOnline -and $result.ADStatus -eq "Offline") {
            $result.Conclusion = "Offline (Per AD last logon)"
        } elseif ($result.Pingable) {
            $result.Conclusion = "Possibly Online (Pingable but identity unverified)"
        } else {
            $result.Conclusion = "Offline (Not pingable and no recent management data)"
        }
    }

    # Output result
    $result
}

# Example usage
Test-ComputerOfflineStatus -ComputerName "WORKSTATION01" -ADThresholdMinutes 60 -MECMThresholdMinutes 60












#AD+MECM
Import-Module ActiveDirectory
$SiteCode = "PS1"
$SCCMServer = "SCCM-SERVER.domain.com"
Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
Set-Location "$SiteCode`:"

function Get-ComputerLastOnline {
    param (
        [string]$ComputerName = "WORKSTATION01"
    )
    
    # AD data
    $adData = Get-ADComputer -Filter "Name -eq '$ComputerName'" -Properties LastLogonTimeStamp
    $adLastOnline = if ($adData.LastLogonTimeStamp) { [DateTime]::FromFileTime($adData.LastLogonTimeStamp) } else { "Never" }
    
    # MECM data
    $mecmData = Get-CMDevice -Name $ComputerName
    $mecmLastOnline = $mecmData.LastActiveTime
    
    [PSCustomObject]@{
        ComputerName    = $ComputerName
        ADLastOnline    = $adLastOnline
        MECMLastOnline  = $mecmLastOnline
    }
}

Get-ComputerLastOnline -ComputerName "WORKSTATION01"












#SCCM
# Requires access to the SCCM PowerShell module (run from SCCM server or console)
# Adjust the site code and server name as needed
$SiteCode = "PS1"  # Your SCCM site code
$SCCMServer = "SCCM-SERVER.domain.com"  # Your SCCM server

# Import the SCCM module
Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1"
Set-Location "$SiteCode`:"

# Function to get last online time from MECM
function Get-MECMComputerLastOnline {
    param (
        [string]$ComputerName = "*"  # Wildcard for all, or specify a name
    )
    
    # Query SCCM for device info
    $devices = Get-CMDevice -Name $ComputerName | Select-Object Name, LastActiveTime, IsActive
    
    # Process results
    $results = foreach ($device in $devices) {
        [PSCustomObject]@{
            ComputerName   = $device.Name
            LastOnline     = $device.LastActiveTime
            IsActive       = $device.IsActive
        }
    }
    
    # Output results
    $results | Sort-Object LastOnline -Descending
}

# Example usage
Get-MECMComputerLastOnline -ComputerName "WORKSTATION01"  # Single computer
# Get-MECMComputerLastOnline  # All computers
















#SPLUNK
# Requires Splunk PowerShell module or REST API credentials
$SplunkServer = "https://splunk.domain.com:8089"
$Username = "admin"
$Password = "yourpassword" | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)

# Function to query Splunk for last online time
function Get-SplunkComputerLastOnline {
    param (
        [string]$ComputerName = "WORKSTATION01"
    )
    
    # Splunk search query (example: last logon event)
    $query = "search host=$ComputerName EventCode=4624 earliest=-30d | head 1 | table _time, host"
    
    # Invoke Splunk REST API
    $response = Invoke-RestMethod -Uri "$SplunkServer/services/search/jobs/export" `
                                  -Method Post `
                                  -Credential $Credential `
                                  -Body @{
                                      search = $query
                                      output_mode = "json"
                                  }
    
    # Parse result
    $lastOnline = ($response.results | ConvertFrom-Json)._time
    [PSCustomObject]@{
        ComputerName = $ComputerName
        LastOnline   = $lastOnline
    }
}

# Example usage
Get-SplunkComputerLastOnline -ComputerName "WORKSTATION01"

















#NESSUS
# Requires Nessus API access
$NessusServer = "https://nessus.domain.com:8834"
$AccessKey = "your_access_key"
$SecretKey = "your_secret_key"

# Function to get last online time from Nessus
function Get-NessusComputerLastOnline {
    param (
        [string]$ComputerName = "WORKSTATION01"
    )
    
    $headers = @{
        "X-ApiKeys" = "accessKey=$AccessKey;secretKey=$SecretKey"
    }
    
    # Query Nessus for host details
    $response = Invoke-RestMethod -Uri "$NessusServer/scans" `
                                 -Method Get `
                                 -Headers $headers
    
    # Find the computer in scan history (simplified)
    $lastScan = $response.scans | Where-Object { $_.name -like "*$ComputerName*" } | 
                Select-Object -First 1 | Select-Object -ExpandProperty last_modification_date
    
    [PSCustomObject]@{
        ComputerName = $ComputerName
        LastOnline   = [DateTime]::FromUnixTime($lastScan)
    }
}

# Example usage
Get-NessusComputerLastOnline -ComputerName "WORKSTATION01"
















#EPO API
# Requires ePO API access
$ePOServer = "https://epo.domain.com:8443"
$Username = "admin"
$Password = "yourpassword" | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)

# Function to get last online time from ePO
function Get-TrellixComputerLastOnline {
    param (
        [string]$ComputerName = "WORKSTATION01"
    )
    
    # ePO API call (example endpoint)
    $response = Invoke-RestMethod -Uri "$ePOServer/remote/system.find?searchText=$ComputerName" `
                                 -Credential $Credential `
                                 -Method Get
    
    $lastOnline = $response | Where-Object { $_."EPOComputerProperties.ComputerName" -eq $ComputerName } | 
                  Select-Object -ExpandProperty "EPOComputerProperties.LastUpdate"
    
    [PSCustomObject]@{
        ComputerName = $ComputerName
        LastOnline   = $lastOnline
    }
}

# Example usage
Get-TrellixComputerLastOnline -ComputerName "WORKSTATION01"

















