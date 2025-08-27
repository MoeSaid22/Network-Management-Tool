class ValidationUtility {
    static [bool] ValidateIP([string]$IPAddress) {
        if ([string]::IsNullOrWhiteSpace($IPAddress)) { return $true }
        try { 
            $null = [System.Net.IPAddress]::Parse($IPAddress.Trim())
            return $true 
        } catch { 
            return $false 
        }
    }
    
    static [void] ValidateDeviceIPs([SiteEntry]$Site) {
        foreach ($switch in $Site.Switches) {
            if (-not [string]::IsNullOrWhiteSpace($switch.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($switch.ManagementIP)) {
                    throw "Invalid Switch IP: $($switch.ManagementIP)"
                }
            }
        }
        
        foreach ($ap in $Site.AccessPoints) {
            if (-not [string]::IsNullOrWhiteSpace($ap.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($ap.ManagementIP)) {
                    throw "Invalid Access Point IP: $($ap.ManagementIP)"
                }
            }
        }
        
        foreach ($ups in $Site.UPSDevices) {
            if (-not [string]::IsNullOrWhiteSpace($ups.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($ups.ManagementIP)) {
                    throw "Invalid UPS IP: $($ups.ManagementIP)"
                }
            }
        }
        
        foreach ($cctv in $Site.CCTVDevices) {
            if (-not [string]::IsNullOrWhiteSpace($cctv.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($cctv.ManagementIP)) {
                    throw "Invalid CCTV IP: $($cctv.ManagementIP)"
                }
            }
        }
        foreach ($printer in $Site.PrinterDevices) {
        if (-not [string]::IsNullOrWhiteSpace($printer.ManagementIP)) {
            if (-not [ValidationUtility]::ValidateIP($printer.ManagementIP)) {
                throw "Invalid Printer IP: $($printer.ManagementIP)"
                }
            }
        }
    }
}

function Validate-SiteBasicInfo {
    param(
        [SiteEntry]$Site,
        [object]$StatusControl = $null,
        [int]$ExcludeSiteID = -1  # For edit mode - exclude current site from duplicate checks
    )
    
    try {
        # Validate required fields
        if ([string]::IsNullOrWhiteSpace($Site.SiteCode)) {
            $errorMsg = "Site Code is required and cannot be empty."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }

        if ([string]::IsNullOrWhiteSpace($Site.SiteSubnet)) {
            $errorMsg = "Site Subnet is required and cannot be empty."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        if ([string]::IsNullOrWhiteSpace($Site.SiteName)) {
            $errorMsg = "Site Name is required and cannot be empty."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # Validate subnet format
        if ($Site.SiteSubnet -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $octets = $Site.SiteSubnet.Split('.')
            $validOctets = $true
            foreach ($octet in $octets) {
                if ([int]$octet -lt 0 -or [int]$octet -gt 255) {
                    $validOctets = $false
                    break
                }
            }
            
            if (-not $validOctets) {
                $errorMsg = "Invalid subnet format. Each octet must be between 0-255."
                if ($StatusControl) {
                    $StatusControl.Text = $errorMsg
                    $StatusControl.Foreground = [System.Windows.Media.Brushes]::Red
                }
                throw $errorMsg
            }
        } else {
            $errorMsg = "Invalid subnet format. Please use format like: XXX.XX.XXX.XXX"
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # Check for duplicates
        $allSites = $siteDataStore.GetAllEntries()
        
        # Check duplicate Site Code (exclude current site if editing)
        $duplicateSiteCode = $allSites | Where-Object { 
            $_.ID -ne $ExcludeSiteID -and $_.SiteCode -eq $Site.SiteCode 
        }
        if ($duplicateSiteCode) {
            $errorMsg = "Site code '$($Site.SiteCode)' already exists in another site."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # Check duplicate Site Subnet (exclude current site if editing)
        $duplicateSubnet = $allSites | Where-Object { 
            $_.ID -ne $ExcludeSiteID -and $_.SiteSubnet -eq $Site.SiteSubnet 
        }
        if ($duplicateSubnet) {
            $errorMsg = "Site subnet '$($Site.SiteSubnet)' already exists in another site."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # If we get here, validation passed
        return $true
        
    } catch {
        # Re-throw the error for the calling function to handle
        throw $_
    }
}

# Helper function to update field only if import has data AND it's different
function Update-FieldIfNotEmpty {
    param($ExistingObject, $ImportObject, $PropertyName, [ref]$ChangesRef, [ref]$ChangeDetailsRef)
    
    $importValue = $ImportObject.$PropertyName
    $existingValue = $ExistingObject.$PropertyName
    
    if (-not [string]::IsNullOrWhiteSpace($importValue)) {
        # Convert both to strings for comparison and trim whitespace
        $importStr = $importValue.ToString().Trim()
        $existingStr = if ($existingValue) { $existingValue.ToString().Trim() } else { "" }
        
        if ($existingStr -ne $importStr) {
            $ExistingObject.$PropertyName = $importValue
            $ChangesRef.Value++
            
            # Add human-readable change description for ALL possible fields
            $fieldDisplayName = switch ($PropertyName) {
                # Basic Info Fields
                "SiteCode" { "Site Code" }
                "SiteSubnet" { "Site Subnet" }
                "SiteSubnetCode" { "Site Subnet Code" }
                "SiteName" { "Site Name" }
                "SiteAddress" { "Site Address" }
                "MainContactName" { "Main Contact Name" }
                "MainContactPhone" { "Main Contact Phone" }
                "SecondContactName" { "Second Contact Name" }
                "SecondContactPhone" { "Second Contact Phone" }
                
                # Firewall Fields
                "FirewallIP" { "Firewall IP" }
                "FirewallName" { "Firewall Name" }
                "FirewallVersion" { "Firewall Version" }
                "FirewallSN" { "Firewall Serial Number" }
                
                # Circuit Fields
                "Vendor" { "Vendor" }
                "CircuitType" { "Circuit Type" }
                "CircuitID" { "Circuit ID" }
                "DownloadSpeed" { "Download Speed" }
                "UploadSpeed" { "Upload Speed" }
                "IPAddress" { "IP Address" }
                "SubnetMask" { "Subnet Mask" }
                "DefaultGateway" { "Default Gateway" }
                "DNS1" { "Primary DNS" }
                "DNS2" { "Secondary DNS" }
                "RouterModel" { "Router Model" }
                "RouterName" { "Router Name" }
                "RouterSN" { "Router Serial Number" }
                "PPPoEUsername" { "PPPoE Username" }
                "PPPoEPassword" { "PPPoE Password" }
                "ModemModel" { "Modem Model" }
                "ModemName" { "Modem Name" }
                "ModemSN" { "Modem Serial Number" }
                
                # Device Fields
                "ManagementIP" { "Management IP" }
                "Name" { "Device Name" }
                "AssetTag" { "Asset Tag" }
                "Version" { "Version" }
                "SerialNumber" { "Serial Number" }
                
                # VLAN Fields
                "VLAN100_Servers" { "VLAN 100 (Servers)" }
                "VLAN101_NetworkDevices" { "VLAN 101 (Network Devices)" }
                "VLAN102_UserDevices" { "VLAN 102 (User Devices)" }
                "VLAN103_UserDevices2" { "VLAN 103 (User Devices 2)" }
                "VLAN104_VOIP" { "VLAN 104 (VOIP)" }
                "VLAN105_WiFiCorp" { "VLAN 105 (WiFi Corporate)" }
                "VLAN106_WiFiBYOD" { "VLAN 106 (WiFi BYOD)" }
                "VLAN107_WiFiGuest" { "VLAN 107 (WiFi Guest)" }
                "VLAN108_Spare" { "VLAN 108 (Spare)" }
                "VLAN109_DMZ" { "VLAN 109 (DMZ)" }
                "VLAN110_CCTV" { "VLAN 110 (CCTV)" }
                
                default { $PropertyName }
            }
            
            $ChangeDetailsRef.Value += $fieldDisplayName
            return $true
        }
    }
    return $false
}

function Update-SiteWithNewData {
    param(
        [SiteEntry]$ExistingSite,
        [SiteEntry]$ImportSite
    )
    
    $changesCount = 0
    $changeDetails = @()
    
    # Pass both changes count and details by reference
    $changesRef = [ref]$changesCount
    $changeDetailsRef = [ref]$changeDetails
    
    # Update basic info
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteSubnet" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteSubnetCode" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteAddress" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "MainContactName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "MainContactPhone" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SecondContactName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SecondContactPhone" $changesRef $changeDetailsRef

    # Update firewall info
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallIP" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallVersion" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallSN" $changesRef $changeDetailsRef
    
    # Update device counts if import has higher counts
    if ($ImportSite.SwitchCount -gt $ExistingSite.SwitchCount) {
        $ExistingSite.SwitchCount = $ImportSite.SwitchCount
        $changesCount++
    }
    if ($ImportSite.APCount -gt $ExistingSite.APCount) {
        $ExistingSite.APCount = $ImportSite.APCount
        $changesCount++
    }
    if ($ImportSite.UPSCount -gt $ExistingSite.UPSCount) {
        $ExistingSite.UPSCount = $ImportSite.UPSCount
        $changesCount++
    }
    if ($ImportSite.CCTVCount -gt $ExistingSite.CCTVCount) {
        $ExistingSite.CCTVCount = $ImportSite.CCTVCount
        $changesCount++
    }
    
    # Update devices - expand arrays if needed and update individual devices
    # Switches
    while ($ExistingSite.Switches.Count -lt $ImportSite.Switches.Count) {
        $ExistingSite.Switches.Add([SwitchInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.Switches.Count; $i++) {
        if ($i -lt $ExistingSite.Switches.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Name"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "AssetTag" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Asset Tag"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "Version" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Version"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "SerialNumber" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Serial Number"
            }
        }
    }

    # Access Points
    while ($ExistingSite.AccessPoints.Count -lt $ImportSite.AccessPoints.Count) {
        $ExistingSite.AccessPoints.Add([AccessPointInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.AccessPoints.Count; $i++) {
        if ($i -lt $ExistingSite.AccessPoints.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Name"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "AssetTag" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Asset Tag"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "Version" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Version"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "SerialNumber" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Serial Number"
            }
        }
    }

    # UPS Devices
    while ($ExistingSite.UPSDevices.Count -lt $ImportSite.UPSDevices.Count) {
        $ExistingSite.UPSDevices.Add([UPSInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.UPSDevices.Count; $i++) {
        if ($i -lt $ExistingSite.UPSDevices.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.UPSDevices[$i] $ImportSite.UPSDevices[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "UPS $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.UPSDevices[$i] $ImportSite.UPSDevices[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "UPS $deviceNum Name"
            }
        }
    }

    # CCTV Devices
    while ($ExistingSite.CCTVDevices.Count -lt $ImportSite.CCTVDevices.Count) {
        $ExistingSite.CCTVDevices.Add([CCTVInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.CCTVDevices.Count; $i++) {
        if ($i -lt $ExistingSite.CCTVDevices.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.CCTVDevices[$i] $ImportSite.CCTVDevices[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Camera $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.CCTVDevices[$i] $ImportSite.CCTVDevices[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Camera $deviceNum Name"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.CCTVDevices[$i] $ImportSite.CCTVDevices[$i] "SerialNumber" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Camera $deviceNum Serial Number"
            }
        }
    }

    # Update ALL circuit fields
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "Vendor" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "CircuitType" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "CircuitID" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DownloadSpeed" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "UploadSpeed" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "IPAddress" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "SubnetMask" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DefaultGateway" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DNS1" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DNS2" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "RouterModel" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "RouterName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "RouterSN" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "PPPoEUsername" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "PPPoEPassword" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "ModemModel" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "ModemName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "ModemSN" $changesRef $changeDetailsRef

    # Update backup circuit if import has backup data
    if ($ImportSite.HasBackupCircuit) {
        if (-not $ExistingSite.HasBackupCircuit) {
            $ExistingSite.HasBackupCircuit = $true
            $changesCount++
            $changeDetails += "Added Backup Circuit"
        }
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "Vendor" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "CircuitType" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "CircuitID" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DownloadSpeed" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "UploadSpeed" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "IPAddress" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "SubnetMask" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DefaultGateway" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DNS1" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DNS2" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "RouterModel" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "RouterName" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "RouterSN" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "PPPoEUsername" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "PPPoEPassword" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "ModemModel" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "ModemName" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "ModemSN" $changesRef $changeDetailsRef
    }

    # Update ALL VLAN fields
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN100_Servers" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN101_NetworkDevices" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN102_UserDevices" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN103_UserDevices2" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN104_VOIP" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN105_WiFiCorp" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN106_WiFiBYOD" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN107_WiFiGuest" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN108_Spare" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN109_DMZ" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN110_CCTV" $changesRef $changeDetailsRef
    
    # Update display properties
    $ExistingSite.UpdateDisplayProperties()
    
    # Return both the site and whether changes were made
    return @{
        Site = $ExistingSite
        HasChanges = ($changesCount -gt 0)
        ChangesCount = $changesCount
        ChangeDetails = $changeDetails
    }
}
