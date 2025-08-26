class FieldMappingManager {
    [hashtable]$MappingGroups
    [object]$MainWindow
    
    FieldMappingManager([object]$mainWindow) {
        $this.MainWindow = $mainWindow
        $this.InitializeMappingGroups()
    }
    
    [void] InitializeMappingGroups() {
        $this.MappingGroups = @{
            'BasicInfo' = @(
                @{Control = 'txtSiteCode'; Property = 'SiteCode'; Required = $true; Type = 'Text'},
                @{Control = 'txtSiteSubnet'; Property = 'SiteSubnet'; Required = $true; Type = 'Text'},
                @{Control = 'txtSiteSubnetCode'; Property = 'SiteSubnetCode'; Required = $false; Type = 'Text'},
                @{Control = 'txtSiteNameManage'; Property = 'SiteName'; Required = $true; Type = 'Text'},
                @{Control = 'txtSiteAddress'; Property = 'SiteAddress'; Required = $false; Type = 'Text'},
                @{Control = 'txtMainContactName'; Property = 'MainContactName'; Required = $false; Type = 'Text'},
                @{Control = 'txtMainContactPhone'; Property = 'MainContactPhone'; Required = $false; Type = 'Text'},
                @{Control = 'txtSecondContactName'; Property = 'SecondContactName'; Required = $false; Type = 'Text'},
                @{Control = 'txtSecondContactPhone'; Property = 'SecondContactPhone'; Required = $false; Type = 'Text'}
            )
            'Firewall' = @(
                @{Control = 'txtFirewallIP'; Property = 'FirewallIP'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtFirewallName'; Property = 'FirewallName'; Required = $false; Type = 'Text'},
                @{Control = 'txtFirewallVersion'; Property = 'FirewallVersion'; Required = $false; Type = 'Text'},
                @{Control = 'txtFirewallSN'; Property = 'FirewallSN'; Required = $false; Type = 'Text'}
            )
            'VLANs' = @(
                @{Control = 'txtVlan100'; Property = 'VLAN100_Servers'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan101'; Property = 'VLAN101_NetworkDevices'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan102'; Property = 'VLAN102_UserDevices'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan103'; Property = 'VLAN103_UserDevices2'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan104'; Property = 'VLAN104_VOIP'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan105'; Property = 'VLAN105_WiFiCorp'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan106'; Property = 'VLAN106_WiFiBYOD'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan107'; Property = 'VLAN107_WiFiGuest'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan108'; Property = 'VLAN108_Spare'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan109'; Property = 'VLAN109_DMZ'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan110'; Property = 'VLAN110_CCTV'; Required = $false; Type = 'Text'}
            )
            'PrimaryCircuit' = @(
                @{Control = 'txtPrimaryVendor'; Property = 'Vendor'; Required = $false; Type = 'Text'},
                @{Control = 'cmbPrimaryCircuitType'; Property = 'CircuitType'; Required = $false; Type = 'ComboBox'},
                @{Control = 'txtPrimaryCircuitID'; Property = 'CircuitID'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryDownloadSpeed'; Property = 'DownloadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryUploadSpeed'; Property = 'UploadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryIPAddress'; Property = 'IPAddress'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimarySubnetMask'; Property = 'SubnetMask'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryDefaultGateway'; Property = 'DefaultGateway'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimaryDNS1'; Property = 'DNS1'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimaryDNS2'; Property = 'DNS2'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimaryRouterModel'; Property = 'RouterModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryRouterName'; Property = 'RouterName'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryRouterSN'; Property = 'RouterSN'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryPPPoEUsername'; Property = 'PPPoEUsername'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryPPPoEPassword'; Property = 'PPPoEPassword'; Required = $false; Type = 'Text'},
                @{Control = 'chkPrimaryHasModem'; Property = 'HasModem'; Required = $false; Type = 'CheckBox'},
                @{Control = 'txtPrimaryModemModel'; Property = 'ModemModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryModemName'; Property = 'ModemName'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryModemSN'; Property = 'ModemSN'; Required = $false; Type = 'Text'}
            )
            'BackupCircuit' = @(
                @{Control = 'txtBackupVendor'; Property = 'Vendor'; Required = $false; Type = 'Text'},
                @{Control = 'cmbBackupCircuitType'; Property = 'CircuitType'; Required = $false; Type = 'ComboBox'},
                @{Control = 'txtBackupCircuitID'; Property = 'CircuitID'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupDownloadSpeed'; Property = 'DownloadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupUploadSpeed'; Property = 'UploadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupIPAddress'; Property = 'IPAddress'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupSubnetMask'; Property = 'SubnetMask'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupDefaultGateway'; Property = 'DefaultGateway'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupDNS1'; Property = 'DNS1'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupDNS2'; Property = 'DNS2'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupRouterModel'; Property = 'RouterModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupRouterName'; Property = 'RouterName'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupRouterSN'; Property = 'RouterSN'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupPPPoEUsername'; Property = 'PPPoEUsername'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupPPPoEPassword'; Property = 'PPPoEPassword'; Required = $false; Type = 'Text'},
                @{Control = 'chkBackupHasModem'; Property = 'HasModem'; Required = $false; Type = 'CheckBox'},
                @{Control = 'txtBackupModemModel'; Property = 'ModemModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupModemName'; Property = 'ModemName'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupModemSN'; Property = 'ModemSN'; Required = $false; Type = 'Text'}
            )
        }
    }
    
    # Validate all mapping groups
    [bool] ValidateAllMappings([object]$dataObject) {
        try {
            # Validate basic info
            $this.ValidateMappingGroup('BasicInfo', $dataObject)
            
            # Validate firewall
            $this.ValidateMappingGroup('Firewall', $dataObject)
            
            # Validate circuits
            $this.ValidateMappingGroup('PrimaryCircuit', $dataObject.PrimaryCircuit)
            if ($dataObject.HasBackupCircuit) {
                $this.ValidateMappingGroup('BackupCircuit', $dataObject.BackupCircuit)
            }
            
            # Validate VLANs
            $this.ValidateMappingGroup('VLANs', $dataObject.VLANs)
            
            return $true
        }
        catch {
            throw $_
        }
    }
    
    # Validate specific mapping group
    [void] ValidateMappingGroup([string]$groupName, [object]$dataObject) {
        $group = $this.MappingGroups[$groupName]
        foreach ($mapping in $group) {
            # Check required fields
            if ($mapping.Required) {
                $value = $dataObject.($mapping.Property)
                if ([string]::IsNullOrWhiteSpace($value)) {
                    throw "Required field missing: $($mapping.Property)"
                }
            }
            
            # Validate field format
            if ($mapping.ContainsKey('Validator')) {
                $value = $dataObject.($mapping.Property)
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    if (-not $this.ValidateField($value, $mapping.Validator)) {
                        throw "Invalid $($mapping.Validator) format: $($mapping.Property) = $value"
                    }
                }
            }
        }
    }
    
    # Field validation
    [bool] ValidateField([string]$value, [string]$validatorType) {
        if ($validatorType -eq 'IP') {
            return [ValidationUtility]::ValidateIP($value)
        }
        return $true
    }
    
    # Set all mappings to UI
    [void] SetAllMappings([object]$dataObject) {
        $this.SetMappingGroup('BasicInfo', $dataObject)
        $this.SetMappingGroup('Firewall', $dataObject)
        $this.SetMappingGroup('VLANs', $dataObject.VLANs)
        $this.SetMappingGroup('PrimaryCircuit', $dataObject.PrimaryCircuit)
        $this.SetMappingGroup('BackupCircuit', $dataObject.BackupCircuit)
    }
    
    # Get all mappings from UI
    [void] GetAllMappings([object]$dataObject) {
        $this.GetMappingGroup('BasicInfo', $dataObject)
        $this.GetMappingGroup('Firewall', $dataObject)
        $this.GetMappingGroup('VLANs', $dataObject.VLANs)
        $this.GetMappingGroup('PrimaryCircuit', $dataObject.PrimaryCircuit)
        $this.GetMappingGroup('BackupCircuit', $dataObject.BackupCircuit)
    }
    
    # Clear all mappings
    [void] ClearAllMappings() {
        $this.ClearMappingGroup('BasicInfo')
        $this.ClearMappingGroup('Firewall')
        $this.ClearMappingGroup('VLANs')
        $this.ClearMappingGroup('PrimaryCircuit')
        $this.ClearMappingGroup('BackupCircuit')
    }
    
    # Set specific mapping group
    [void] SetMappingGroup([string]$groupName, [object]$dataObject) {
        $group = $this.MappingGroups[$groupName]
        foreach ($mapping in $group) {
            $control = $this.MainWindow.FindName($mapping.Control)
            if ($control) {
                $value = Get-SafeValue $dataObject.($mapping.Property)
                
                switch ($mapping.Type) {
                    'Text' { $control.Text = $value }
                    'CheckBox' { $control.IsChecked = [bool]$value }
                    'ComboBox' { $this.SetComboBoxSelection($control, $value) }
                }
            }
        }
    }
    
    # Get specific mapping group
[void] GetMappingGroup([string]$groupName, [object]$dataObject) {
    $group = $this.MappingGroups[$groupName]    
    foreach ($mapping in $group) {
        $control = $this.MainWindow.FindName($mapping.Control)
        if ($control) {
            $controlValue = ""
            switch ($mapping.Type) {
                'Text' { 
                    $controlValue = $control.Text.Trim()
                    $dataObject.($mapping.Property) = $controlValue
                }
                'CheckBox' { 
                    $controlValue = $control.IsChecked
                    $dataObject.($mapping.Property) = $controlValue
                }
                'ComboBox' { 
                    if ($control.SelectedItem) {
                        $controlValue = $control.SelectedItem.Content
                        $dataObject.($mapping.Property) = $controlValue
                    }
                }
            }
        } else {
        }
    }
}
    
    # Clear specific mapping group
    [void] ClearMappingGroup([string]$groupName) {
        $group = $this.MappingGroups[$groupName]
        foreach ($mapping in $group) {
            $control = $this.MainWindow.FindName($mapping.Control)
            if ($control) {
                switch ($mapping.Type) {
                    'Text' { $control.Text = "" }
                    'CheckBox' { $control.IsChecked = $false }
                    'ComboBox' { $control.SelectedIndex = -1 }
                }
            }
        }
    }
        
    # Helper method for ComboBox selection
    [void] SetComboBoxSelection([System.Windows.Controls.ComboBox]$ComboBox, [string]$Value) {
        Set-ComboBoxValue $ComboBox $Value -ByContent
    }
}