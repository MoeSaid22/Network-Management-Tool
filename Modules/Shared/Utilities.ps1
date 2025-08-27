# Shared Utilities Module
# Contains common utility functions used across the application

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

function Get-SafeValue {
    param($Value)
    
    if ($null -eq $Value) { 
        return "" 
    }
    return $Value.ToString()
}