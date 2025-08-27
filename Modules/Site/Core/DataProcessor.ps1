# Site Data Processing Module  
# Contains functions for processing and manipulating site data

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
    
    # Pass changes count by reference to all calls
    $changesRef = [ref]$changesCount
    
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
            
            # Check each field and add device number to change description
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                # Replace the last added item with device-specific description
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