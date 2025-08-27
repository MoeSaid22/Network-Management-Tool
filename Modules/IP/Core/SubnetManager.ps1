class SubnetDataStore {
    hidden [string]$DataFile
    hidden [System.Collections.Generic.List[SubnetEntry]]$Entries

    SubnetDataStore() {
        # Initialize Entries list in constructor
        $this.Entries = [System.Collections.Generic.List[SubnetEntry]]::new()
    }

    [void] SetDataPath([string]$path) {
        $this.DataFile = $path
        $this.LoadData()
    }

    [void] LoadData() {
        try {
            if (Test-Path $this.DataFile) {
                $jsonData = Get-Content $this.DataFile -Raw | ConvertFrom-Json
                $this.Entries.Clear()
                
                if ($jsonData) {
                    foreach ($item in $jsonData) {
                        if ($item -and $item.IP_Subnet) {
                            $entry = [SubnetEntry]::new(
                                $item.IP_Subnet,
                                $item.VLAN_ID,
                                $item.VLAN_Name,
                                $item.Site_Name
                            )
                            $entry.ID = $item.ID
                            $this.Entries.Add($entry)
                        }
                    }
                }
            }
        } catch {
            $this.Entries.Clear()
        }
    }

    [void] SaveData() {
        try {
            if ($this.Entries.Count -eq 0) {
                if (Test-Path $this.DataFile) {
                    Remove-Item $this.DataFile -Force
                }
            } else {
                $this.Entries | ConvertTo-Json -Depth 3 | Set-Content $this.DataFile -Encoding UTF8
            }
        } catch {
        }
    }

    [SubnetEntry[]] GetAllEntries() {
        if ($this.Entries -eq $null -or $this.Entries.Count -eq 0) {
            return @()
        }
        return $this.Entries.ToArray()
    }

    [bool] AddEntry([SubnetEntry]$entry) {
        try {
            if ($entry -eq $null) { return $false }
            
            # Check for duplicates
            foreach ($existing in $this.Entries) {
                if ($existing.IP_Subnet -eq $entry.IP_Subnet) {
                    return $false
                }
            }
            
            $entry.ID = $this.GetNextAvailableId()
            $this.Entries.Add($entry)
            $this.SaveData()
            return $true
        } catch {
            return $false
        }
    }

    [bool] DeleteEntries([int[]]$ids) {
        try {
            if ($ids -eq $null -or $ids.Count -eq 0) { return $false }
            
            $countBefore = $this.Entries.Count
            $newList = [System.Collections.Generic.List[SubnetEntry]]::new()
            
            foreach ($entry in $this.Entries) {
                if ($entry.ID -notin $ids) {
                    $newList.Add($entry)
                }
            }
            
            $this.Entries = $newList
            $this.SaveData()
            return $this.Entries.Count -lt $countBefore
        } catch {
            return $false
        }
    }

    hidden [int] GetNextAvailableId() {
        if ($this.Entries.Count -eq 0) { return 1 }
        $maxId = 0
        foreach ($entry in $this.Entries) {
            if ($entry.ID -gt $maxId) { $maxId = $entry.ID }
        }
        return $maxId + 1
    }
}

function Update-SubnetDataGridWithSearch {
    try {
        $searchTerm = if ($txtSearch -ne $null) { $txtSearch.Text } else { "" }
        
        # Get all data from the data store
        $allData = @($subnetDataStore.GetAllEntries())
        
        # Filter data if search term exists
        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
            $searchTerm = $searchTerm.Trim().ToLower()
            $filteredData = $allData | Where-Object {
                $_.IP_Subnet.ToLower().Contains($searchTerm) -or
                $_.VLAN_ID.ToString().Contains($searchTerm) -or
                $_.VLAN_Name.ToLower().Contains($searchTerm) -or
                $_.Site_Name.ToLower().Contains($searchTerm)
            }
            # Force array conversion to ensure we always have an IEnumerable
            $allData = @($filteredData)
        }
        
        # Sort by ID numerically and update DataGrid
        if ($allData.Count -gt 0) {
            $allData = @($allData | Sort-Object -Property @{Expression={[int]$_.ID}; Ascending=$true})
        }
        
        # Safely update DataGrid - ensure we always pass an array/collection
        if ($dgSubnets -ne $null) {
            # Convert to ArrayList to ensure proper IEnumerable interface
            $dataGridSource = New-Object System.Collections.ArrayList
            foreach ($item in $allData) {
                $null = $dataGridSource.Add($item)
            }
            $dgSubnets.ItemsSource = $dataGridSource
        }
        
        # Update status bar only if values changed
        $newCount = $allData.Count
        if ($txtStatusBarSubnets -ne $null) {
            $txtStatusBarSubnets.Text = "Total Subnets: $newCount"
        }
        if ($txtStatusBarSelected -ne $null) {
            $txtStatusBarSelected.Text = "Selected: None"
        }
        
    } catch {
        
        # Ensure DataGrid is not null even on error
        if ($dgSubnets -ne $null) {
            $dgSubnets.ItemsSource = New-Object System.Collections.ArrayList
        }
    }
}

function Test-IPSubnetFormat {
    param([string]$IPSubnet)
    
    try {
        # Check if input is null or empty
        if ([string]::IsNullOrWhiteSpace($IPSubnet)) {
            return $false
        }
        
        # Basic format check
        if ($IPSubnet -notmatch '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})/(\d{1,2})$') {
            return $false
        }
        
        # Split IP and CIDR parts
        $parts = $IPSubnet.Split('/')
        if ($parts.Count -ne 2) {
            return $false
        }
        
        $ipPart = $parts[0]
        $cidrPart = $null
        
        # Safely parse CIDR
        if (-not [int]::TryParse($parts[1], [ref]$cidrPart)) {
            return $false
        }
        
        # Validate CIDR range (0-32 for IPv4)
        if ($cidrPart -lt 0 -or $cidrPart -gt 32) {
            return $false
        }
        
        # Validate each IP octet (0-255)
        $octets = $ipPart.Split('.')
        if ($octets.Count -ne 4) {
            return $false
        }
        
        foreach ($octet in $octets) {
            $octetValue = $null
            if (-not [int]::TryParse($octet, [ref]$octetValue)) {
                return $false
            }
            if ($octetValue -lt 0 -or $octetValue -gt 255) {
                return $false
            }
        }
        
        # Additional validation: Parse as IPAddress to catch edge cases
        $ipAddress = $null
        if (-not [System.Net.IPAddress]::TryParse($ipPart, [ref]$ipAddress)) {
            return $false
        }
        
        # Validate that it's a proper network address (not a host address)
        # Calculate network address and compare with input
        $ip = [System.Net.IPAddress]::Parse($ipPart)
        $mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl (32 - $cidrPart))
        $maskBytes = [BitConverter]::GetBytes($mask)
        $ipBytes = $ip.GetAddressBytes()
        
        # Calculate network address
        $networkBytes = @()
        for ($i = 0; $i -lt 4; $i++) {
            $networkBytes += ($ipBytes[$i] -band $maskBytes[$i])
        }
        
        # Check if provided IP matches the network address
        for ($i = 0; $i -lt 4; $i++) {
            if ($ipBytes[$i] -ne $networkBytes[$i]) {
                # This is a host address, not a network address
                return $false
            }
        }
        
        return $true
    } catch {
        return $false
    }
}

function Test-VlanId {
    param([string]$VlanId)
    
    try {
        if ([string]::IsNullOrWhiteSpace($VlanId)) {
            return $false
        }
        
        $result = $null
        if ([int]::TryParse($VlanId.Trim(), [ref]$result)) {
            return $result -ge 1 -and $result -le 4094  # Valid VLAN range
        }
        return $false
        
    } catch {
        return $false
    }
}

function Show-IPValidationError {
    param(
        [string]$Message,
        [string]$Title = "Validation Error"
    )
    
    # Use the import status control since txtBlkStatus might not exist
    if ($txtBlkImportStatus -ne $null) {
        try {
            $txtBlkImportStatus.Text = $Message
            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
        } catch {
        }
    }
    
    # Show dialog
    Show-CustomDialog $Message $Title "OK" "Error"
}

function Remove-CSVInjection {
    param([string]$Value)
    
    try {
        if ([string]::IsNullOrWhiteSpace($Value)) { 
            return "" 
        }
        
        # Remove dangerous characters that could cause CSV injection
        $dangerous = @('=', '+', '-', '@', '|', '%')
        foreach ($char in $dangerous) {
            if ($Value.StartsWith($char)) {
                $Value = "'" + $Value  # Prefix with quote to neutralize
            }
        }
        
        # Remove control characters and limit length
        $Value = $Value -replace '[\x00-\x1F\x7F]', ''
        return $Value.Substring(0, [Math]::Min(255, $Value.Length))
    } catch {
        return ""
    }
}

function Add-SubnetEntry {
    param (
        [string]$IpSubnet,
        [int]$VlanId,
        [string]$VlanName,
        [string]$SiteName
    )
    
    try {
        # Validate inputs are not null
        if ([string]::IsNullOrWhiteSpace($IpSubnet)) {
            throw "IP Subnet cannot be empty"
        }
        if ([string]::IsNullOrWhiteSpace($VlanName)) {
            throw "VLAN Name cannot be empty"
        }
        if ([string]::IsNullOrWhiteSpace($SiteName)) {
            throw "Site Name cannot be empty"
        }
        if ($VlanId -le 0) {
            throw "VLAN ID must be a positive number"
        }
        
        # Check if subnetDataStore exists
        if ($subnetDataStore -eq $null) {
            throw "Subnet data store is not initialized"
        }
        
        $entry = [SubnetEntry]::new($IpSubnet, $VlanId, $VlanName, $SiteName)
        
        if ($subnetDataStore.AddEntry($entry)) {
            $successMsg = "Entry added successfully! (ID: $($entry.ID))"
            
            # Safely update status if control exists
            if ($txtBlkImportStatus -ne $null) {
                try {
                    $txtBlkImportStatus.Text = $successMsg
                    $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                } catch {
                }
            }
            return $true
        } else {
            Show-IPValidationError "Error: Subnet already exists"
            return $false
        }
    } catch {
        Show-IPValidationError "Error: $($_.Exception.Message)"
        return $false
    }
}

function Lookup-IpAddress {
    param (
        [string]$IpAddress
    )

    try {
        # Always hide results at start
        if ($grpLookupResults -ne $null) {
            $grpLookupResults.Visibility = "Collapsed"
        }
        
        # Clear previous results
        if ($txtBlkSearchedIp -ne $null) { $txtBlkSearchedIp.Text = "" }
        if ($txtBlkMatchedSubnet -ne $null) { $txtBlkMatchedSubnet.Text = "" }
        if ($txtBlkVlanId -ne $null) { $txtBlkVlanId.Text = "" }
        if ($txtBlkVlanName -ne $null) { $txtBlkVlanName.Text = "" }
        if ($txtBlkSiteName -ne $null) { $txtBlkSiteName.Text = "" }

        if ([string]::IsNullOrEmpty($IpAddress)) {
            Show-CustomDialog "Please enter an IP address" "Input Required" "OK" "Warning"
            return
        }

        try {
            $ip = [System.Net.IPAddress]::Parse($IpAddress)
        } catch {
            Show-CustomDialog "Invalid IP address format" "Input Error" "OK" "Error"
            return
        }

        $foundMatch = $false
        $allEntries = $subnetDataStore.GetAllEntries()
        
        foreach ($entry in $allEntries) {
            if ($entry -eq $null) { continue }
            
            $subnetCidr = $entry.IP_Subnet
            if (-not $subnetCidr -or $subnetCidr -notmatch '^(\d{1,3}\.){3}\d{1,3}/\d{1,2}$') {
                continue
            }

            $parts = $subnetCidr.Split('/')
            try {
                $subnetIp = [System.Net.IPAddress]::Parse($parts[0])
                $prefixLength = [int]$parts[1]
                
                $mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl (32 - $prefixLength))
                $maskBytes = [BitConverter]::GetBytes($mask)
                
                $ipBytes = $ip.GetAddressBytes()
                $subnetIpBytes = $subnetIp.GetAddressBytes()
                
                $isInSubnet = $true
                for ($i = 0; $i -lt 4; $i++) {
                    if (($ipBytes[$i] -band $maskBytes[$i]) -ne ($subnetIpBytes[$i] -band $maskBytes[$i])) {
                        $isInSubnet = $false
                        break
                    }
                }

                if ($isInSubnet) {
                    # Only show results on successful match
                    if ($grpLookupResults -ne $null) { $grpLookupResults.Visibility = "Visible" }
                    if ($txtBlkSearchedIp -ne $null) { $txtBlkSearchedIp.Text = $IpAddress }
                    if ($txtBlkMatchedSubnet -ne $null) { $txtBlkMatchedSubnet.Text = $subnetCidr }
                    if ($txtBlkVlanId -ne $null) { $txtBlkVlanId.Text = $entry.VLAN_ID }
                    if ($txtBlkVlanName -ne $null) { $txtBlkVlanName.Text = $entry.VLAN_Name }
                    if ($txtBlkSiteName -ne $null) { $txtBlkSiteName.Text = $entry.Site_Name }
                    $foundMatch = $true
                    break
                }
            } catch {
                continue
            }
        }

        if (-not $foundMatch) {
            Show-CustomDialog "IP address $IpAddress not found in the network database." "Not Found" "OK" "Information"
        }
    } catch {
        Show-CustomDialog "Error during IP lookup: $_" "Error" "OK" "Error"
    }
}

