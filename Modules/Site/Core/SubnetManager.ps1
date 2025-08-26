function Update-VLANsAndIPsFromSubnet {
    param(
        [string]$SubnetInput,
        [hashtable]$VLANControls,
        [object]$DeviceManager,
        [object]$FirewallIPControl,
        [object]$SiteSubnetCodeControl
    )
    
    if ([string]::IsNullOrWhiteSpace($SubnetInput)) { return }
    
    # Parse subnet (e.g., "10.107.0.0" -> "10.107")
    if ($SubnetInput -match '^(\d+\.\d+)\.') {
        $baseSubnet = $matches[1]
        
        # Auto-populate VLAN fields
        if ($VLANControls.VLAN100) { $VLANControls.VLAN100.Text = "$baseSubnet.10.0" }
        if ($VLANControls.VLAN101) { $VLANControls.VLAN101.Text = "$baseSubnet.20.0" }
        if ($VLANControls.VLAN102) { $VLANControls.VLAN102.Text = "$baseSubnet.102.0" }
        if ($VLANControls.VLAN103) { $VLANControls.VLAN103.Text = "$baseSubnet.103.0" }
        if ($VLANControls.VLAN104) { $VLANControls.VLAN104.Text = "$baseSubnet.40.0" }
        if ($VLANControls.VLAN105) { $VLANControls.VLAN105.Text = "$baseSubnet.50.0" }
        if ($VLANControls.VLAN106) { $VLANControls.VLAN106.Text = "$baseSubnet.60.0" }
        if ($VLANControls.VLAN107) { $VLANControls.VLAN107.Text = "$baseSubnet.70.0" }
        if ($VLANControls.VLAN108) { $VLANControls.VLAN108.Text = "$baseSubnet.80.0" }
        if ($VLANControls.VLAN109) { $VLANControls.VLAN109.Text = "$baseSubnet.90.0" }
        if ($VLANControls.VLAN110) { $VLANControls.VLAN110.Text = "$baseSubnet.110.0" }
        
        # Auto-fill firewall IP
        if ($FirewallIPControl) {
            $firewallIP = "$baseSubnet.20.1"
            $FirewallIPControl.Text = $firewallIP
        }
        
        # Auto-fill device IPs using device manager
        $DeviceManager.UpdateDeviceIPsFromSubnet('Switch', $baseSubnet)
        $DeviceManager.UpdateDeviceIPsFromSubnet('AccessPoint', $baseSubnet)
        $DeviceManager.UpdateDeviceIPsFromSubnet('UPS', $baseSubnet)
        $DeviceManager.UpdateDeviceIPsFromSubnet('CCTV', $baseSubnet)
    }
    
    # Auto-populate site subnet code
    if ($SiteSubnetCodeControl -and $SubnetInput -match '^(\d+)\.(\d+)\.') {
        $secondOctet = $matches[2]
        $siteSubnetCode = [int]$secondOctet
        
        # Only auto-fill if the field is empty
        if ([string]::IsNullOrWhiteSpace($SiteSubnetCodeControl.Text)) {
            $SiteSubnetCodeControl.Text = $siteSubnetCode.ToString()
        }
    }
}

function Update-DeviceNamesFromSiteCode {
    param(
        [string]$SiteCode,
        [object]$DeviceManager,
        [object]$FirewallNameControl
    )
    
    if ([string]::IsNullOrWhiteSpace($SiteCode)) { return }
    
    # Update all device types using the device manager
    $DeviceManager.UpdateDeviceNamesFromSiteCode('Switch', $SiteCode)
    $DeviceManager.UpdateDeviceNamesFromSiteCode('AccessPoint', $SiteCode)
    $DeviceManager.UpdateDeviceNamesFromSiteCode('UPS', $SiteCode)
    $DeviceManager.UpdateDeviceNamesFromSiteCode('CCTV', $SiteCode)
    
    # Update firewall name (not managed by DeviceManager)
    if ($FirewallNameControl -and -not [string]::IsNullOrWhiteSpace($SiteCode)) {
        $siteCodeUpper = $SiteCode.Trim().ToUpper()
        $FirewallNameControl.Text = "$siteCodeUpper-FW"
    }
}