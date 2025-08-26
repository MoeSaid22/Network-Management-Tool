class SiteDataStore {
    hidden [string]$DataFile = "$PSScriptRoot\site_data.json"
    hidden [System.Collections.Generic.List[SiteEntry]]$Entries

    SiteDataStore() {
        $this.LoadData()
    }

    # Load site data from JSON file
    [void] LoadData() {   
    if (Test-Path $this.DataFile) {
            try {
                $jsonData = Get-Content $this.DataFile | ConvertFrom-Json
                $this.Entries = [System.Collections.Generic.List[SiteEntry]]::new()
                foreach ($item in $jsonData) {
                    $site = [SiteEntry]::new()
                    $site.ID = $item.ID
                    
                    # Basic Info
                    $site.SiteCode = $item.SiteCode
                    $site.SiteSubnet = Get-SafeValue $item.SiteSubnet
                    $site.SiteSubnetCode = $item.SiteSubnetCode
                    $site.SiteName = $item.SiteName
                    $site.SiteAddress = $item.SiteAddress
                    $site.MainContactName = $item.MainContactName
                    $site.MainContactPhone = $item.MainContactPhone
                    $site.SecondContactName = $item.SecondContactName
                    $site.SecondContactPhone = $item.SecondContactPhone
                    
                    # Network Equipment
                    $site.SwitchCount = $item.SwitchCount
                    $site.Switches = [System.Collections.Generic.List[SwitchInfo]]::new()
                    if ($item.Switches) {
                        foreach ($switchItem in $item.Switches) {
                            $switch = [SwitchInfo]::new()
                            $switch.ManagementIP = Get-SafeValue $switchItem.ManagementIP
                            $switch.Name = Get-SafeValue $switchItem.Name
                            $switch.AssetTag = Get-SafeValue $switchItem.AssetTag
                            $switch.Version = Get-SafeValue $switchItem.Version
                            $switch.SerialNumber = Get-SafeValue $switchItem.SerialNumber
                            $site.Switches.Add($switch)
                        }
                    }

                    # Access Points
                    $site.APCount = if ($item.APCount) { $item.APCount } else { 0 }
                    $site.AccessPoints = [System.Collections.Generic.List[AccessPointInfo]]::new()
                    if ($item.AccessPoints) {
                        foreach ($apItem in $item.AccessPoints) {
                            $ap = [AccessPointInfo]::new()
                            $ap.ManagementIP = Get-SafeValue $apItem.ManagementIP
                            $ap.Name = Get-SafeValue $apItem.Name
                            $ap.AssetTag = Get-SafeValue $apItem.AssetTag
                            $ap.Version = Get-SafeValue $apItem.Version
                            $ap.SerialNumber = Get-SafeValue $apItem.SerialNumber
                            $site.AccessPoints.Add($ap)
                        }
                    }

                    # UPS
                    $site.UPSCount = if ($item.UPSCount) { $item.UPSCount } else { 0 }
                    $site.UPSDevices = [System.Collections.Generic.List[UPSInfo]]::new()
                    if ($item.UPSDevices) {
                        foreach ($upsItem in $item.UPSDevices) {
                            $ups = [UPSInfo]::new()
                            $ups.ManagementIP = Get-SafeValue $upsItem.ManagementIP
                            $ups.Name = Get-SafeValue $upsItem.Name
                            $ups.AssetTag = Get-SafeValue $upsItem.AssetTag
                            $ups.Version = Get-SafeValue $upsItem.Version
                            $ups.SerialNumber = Get-SafeValue $upsItem.SerialNumber
                            $site.UPSDevices.Add($ups)
                        }
                    }

                    # CCTV
                    $site.CCTVCount = if ($item.CCTVCount) { $item.CCTVCount } else { 0 }
                    $site.CCTVDevices = [System.Collections.Generic.List[CCTVInfo]]::new()
                    if ($item.CCTVDevices) {
                        foreach ($cctvItem in $item.CCTVDevices) {
                            $cctv = [CCTVInfo]::new()
                            $cctv.ManagementIP = Get-SafeValue $cctvItem.ManagementIP
                            $cctv.Name = Get-SafeValue $cctvItem.Name
                            $cctv.SerialNumber = Get-SafeValue $cctvItem.SerialNumber
                            $site.CCTVDevices.Add($cctv)
                        }
                    }

                    # Printer
                    $site.PrinterCount = if ($item.PrinterCount) { $item.PrinterCount } else { 0 }
                    $site.PrinterDevices = [System.Collections.Generic.List[PrinterInfo]]::new()
                    if ($item.PrinterDevices) {
                        foreach ($printerItem in $item.PrinterDevices) {
                            $printer = [PrinterInfo]::new()
                            $printer.ManagementIP = Get-SafeValue $printerItem.ManagementIP
                            $printer.Name = Get-SafeValue $printerItem.Name
                            $printer.Model = Get-SafeValue $printerItem.Model
                            $printer.SerialNumber = Get-SafeValue $printerItem.SerialNumber
                            $site.PrinterDevices.Add($printer)
                        }
                    }
                    
                    $site.FirewallIP = Get-SafeValue $item.FirewallIP
                    $site.FirewallName = Get-SafeValue $item.FirewallName
                    $site.FirewallVersion = Get-SafeValue $item.FirewallVersion
                    $site.FirewallSN = Get-SafeValue $item.FirewallSN
                    
                    # Circuits
                    if ($item.PrimaryCircuit) {
                        $site.PrimaryCircuit.Vendor = Get-SafeValue $item.PrimaryCircuit.Vendor
                        $site.PrimaryCircuit.CircuitType = Get-SafeValue $item.PrimaryCircuit.CircuitType
                        $site.PrimaryCircuit.CircuitID = Get-SafeValue $item.PrimaryCircuit.CircuitID
                        $site.PrimaryCircuit.DownloadSpeed = Get-SafeValue $item.PrimaryCircuit.DownloadSpeed
                        $site.PrimaryCircuit.UploadSpeed = Get-SafeValue $item.PrimaryCircuit.UploadSpeed
                        $site.PrimaryCircuit.IPAddress = Get-SafeValue $item.PrimaryCircuit.IPAddress
                        $site.PrimaryCircuit.SubnetMask = Get-SafeValue $item.PrimaryCircuit.SubnetMask
                        $site.PrimaryCircuit.DefaultGateway = Get-SafeValue $item.PrimaryCircuit.DefaultGateway
                        $site.PrimaryCircuit.DNS1 = Get-SafeValue $item.PrimaryCircuit.DNS1
                        $site.PrimaryCircuit.DNS2 = Get-SafeValue $item.PrimaryCircuit.DNS2
                        $site.PrimaryCircuit.RouterModel = Get-SafeValue $item.PrimaryCircuit.RouterModel
                        $site.PrimaryCircuit.RouterName = Get-SafeValue $item.PrimaryCircuit.RouterName
                        $site.PrimaryCircuit.RouterSN = Get-SafeValue $item.PrimaryCircuit.RouterSN
                        $site.PrimaryCircuit.PPPoEUsername = Get-SafeValue $item.PrimaryCircuit.PPPoEUsername
                        $site.PrimaryCircuit.PPPoEPassword = Get-SafeValue $item.PrimaryCircuit.PPPoEPassword
                        $site.PrimaryCircuit.HasModem = if ($item.PrimaryCircuit.HasModem) { $item.PrimaryCircuit.HasModem } else { $false }
                        $site.PrimaryCircuit.ModemModel = Get-SafeValue $item.PrimaryCircuit.ModemModel
                        $site.PrimaryCircuit.ModemName = Get-SafeValue $item.PrimaryCircuit.ModemName
                        $site.PrimaryCircuit.ModemSN = Get-SafeValue $item.PrimaryCircuit.ModemSN
                    }
                    
                    $site.HasBackupCircuit = if ($item.HasBackupCircuit) { $item.HasBackupCircuit } else { $false }
                    if ($item.BackupCircuit -and $site.HasBackupCircuit) {
                        $site.BackupCircuit.Vendor = Get-SafeValue $item.BackupCircuit.Vendor
                        $site.BackupCircuit.CircuitType = Get-SafeValue $item.BackupCircuit.CircuitType
                        $site.BackupCircuit.CircuitID = Get-SafeValue $item.BackupCircuit.CircuitID
                        $site.BackupCircuit.DownloadSpeed = Get-SafeValue $item.BackupCircuit.DownloadSpeed
                        $site.BackupCircuit.UploadSpeed = Get-SafeValue $item.BackupCircuit.UploadSpeed
                        $site.BackupCircuit.IPAddress = Get-SafeValue $item.BackupCircuit.IPAddress
                        $site.BackupCircuit.SubnetMask = Get-SafeValue $item.BackupCircuit.SubnetMask
                        $site.BackupCircuit.DefaultGateway = Get-SafeValue $item.BackupCircuit.DefaultGateway
                        $site.BackupCircuit.DNS1 = Get-SafeValue $item.BackupCircuit.DNS1
                        $site.BackupCircuit.DNS2 = Get-SafeValue $item.BackupCircuit.DNS2
                        $site.BackupCircuit.RouterModel = Get-SafeValue $item.BackupCircuit.RouterModel
                        $site.BackupCircuit.RouterName = Get-SafeValue $item.BackupCircuit.RouterName
                        $site.BackupCircuit.RouterSN = Get-SafeValue $item.BackupCircuit.RouterSN
                        $site.BackupCircuit.PPPoEUsername = Get-SafeValue $item.BackupCircuit.PPPoEUsername
                        $site.BackupCircuit.PPPoEPassword = Get-SafeValue $item.BackupCircuit.PPPoEPassword
                        $site.BackupCircuit.HasModem = if ($item.BackupCircuit.HasModem) { $item.BackupCircuit.HasModem } else { $false }
                        $site.BackupCircuit.ModemModel = Get-SafeValue $item.BackupCircuit.ModemModel
                        $site.BackupCircuit.ModemName = Get-SafeValue $item.BackupCircuit.ModemName
                        $site.BackupCircuit.ModemSN = Get-SafeValue $item.BackupCircuit.ModemSN
                    }
                    
                    # VLANs
                    if ($item.VLANs) {
                        $site.VLANs.VLAN100_Servers = Get-SafeValue $item.VLANs.VLAN100_Servers
                        $site.VLANs.VLAN101_NetworkDevices = Get-SafeValue $item.VLANs.VLAN101_NetworkDevices
                        $site.VLANs.VLAN102_UserDevices = Get-SafeValue $item.VLANs.VLAN102_UserDevices
                        $site.VLANs.VLAN103_UserDevices2 = Get-SafeValue $item.VLANs.VLAN103_UserDevices2
                        $site.VLANs.VLAN104_VOIP = Get-SafeValue $item.VLANs.VLAN104_VOIP
                        $site.VLANs.VLAN105_WiFiCorp = Get-SafeValue $item.VLANs.VLAN105_WiFiCorp
                        $site.VLANs.VLAN106_WiFiBYOD = Get-SafeValue $item.VLANs.VLAN106_WiFiBYOD
                        $site.VLANs.VLAN107_WiFiGuest = Get-SafeValue $item.VLANs.VLAN107_WiFiGuest
                        $site.VLANs.VLAN108_Spare = Get-SafeValue $item.VLANs.VLAN108_Spare
                        $site.VLANs.VLAN109_DMZ = Get-SafeValue $item.VLANs.VLAN109_DMZ
                        $site.VLANs.VLAN110_CCTV = Get-SafeValue $item.VLANs.VLAN110_CCTV
                    }
                    
                    $site.UpdateDisplayProperties()
                    $this.Entries.Add($site)
                }
            } catch {
                [System.Windows.MessageBox]::Show("Error loading site data: $_", "Data Load Error", "OK", "Warning")
                $this.Entries = [System.Collections.Generic.List[SiteEntry]]::new()
                $this.SaveData()
            }
        } else {
            $this.Entries = [System.Collections.Generic.List[SiteEntry]]::new()
            $this.SaveData()
        }
    }

    # Save site data to JSON file
    [void] SaveData() {
    try {        
        if ($this.Entries.Count -eq 0) {
            if (Test-Path $this.DataFile) {
                Remove-Item $this.DataFile -Force
            }
        } else {
            $this.Entries | ConvertTo-Json -Depth 10 | Set-Content $this.DataFile
        }
    } catch {
        [System.Windows.MessageBox]::Show("Error saving site data: $_", "Data Save Error", "OK", "Error")
    }
}


    # Get all site entries
    [SiteEntry[]] GetAllEntries() {
        return $this.Entries.ToArray()
    }

    # Add new site entry with duplicate validation
    [bool] AddEntry([SiteEntry]$entry) {
    # Check for duplicate Site Code
    if ($this.Entries.SiteCode -contains $entry.SiteCode) {
        throw "Site code '$($entry.SiteCode)' already exists"
    }
    
    # Check for duplicate Site Subnet
    if ($this.Entries.SiteSubnet -contains $entry.SiteSubnet) {
        throw "Site subnet '$($entry.SiteSubnet)' already exists"
    }
    
    $entry.ID = $this.GetNextAvailableId()
    $entry.UpdateDisplayProperties()
    $this.Entries.Add($entry)
    $this.SaveData()
    return $true
    }

    # Update existing site entry
    [bool] UpdateEntry([SiteEntry]$entry) {
        $existingIndex = -1
        for ($i = 0; $i -lt $this.Entries.Count; $i++) {
            if ($this.Entries[$i].ID -eq $entry.ID) {
                $existingIndex = $i
                break
            }
        }
        
        if ($existingIndex -ge 0) {
            $entry.UpdateDisplayProperties()
            $this.Entries[$existingIndex] = $entry
            $this.SaveData()
            return $true
        }
        return $false
    }

    # Delete multiple site entries by ID
    [bool] DeleteEntries([int[]]$ids) {
        $countBefore = $this.Entries.Count
        $newEntries = [System.Collections.Generic.List[SiteEntry]]::new()
        
        foreach ($entry in $this.Entries) {
            if ($entry.ID -notin $ids) {
                $newEntries.Add($entry)
            }
        }
        
        $this.Entries = $newEntries
        
        if ($this.Entries.Count -lt $countBefore) {
            $this.SaveData()
            return $true
        }
        return $false
    }

    # Get next available ID for new entries
    hidden [int] GetNextAvailableId() {
        if ($this.Entries.Count -eq 0) { return 1 }
        $maxId = ($this.Entries.ID | Measure-Object -Maximum).Maximum
        for ($i = 1; $i -le $maxId; $i++) {
            if ($i -notin $this.Entries.ID) { return $i }
        }
        return $maxId + 1
    }
}

function Get-SafeValue {
    param([object]$Value)
    if ($Value) { return $Value.ToString() } else { return "" }
}