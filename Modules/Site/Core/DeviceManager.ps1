class DevicePanelManager {
    [hashtable]$Configurations
    [hashtable]$StackPanels
    [hashtable]$ComboBoxes
    [object]$MainWindow
    
    DevicePanelManager([object]$mainWindow) {
        $this.MainWindow = $mainWindow
        $this.InitializeConfigurations()
        $this.InitializeUIReferences()
    }
    
    [void] InitializeConfigurations() {
        $this.Configurations = @{
            'Switch' = [DeviceConfiguration]::new(
                'Switch',
                'SWT', 
                '.20',
                5,
                10,
                @('ManagementIP', 'Name', 'AssetTag', 'Version', 'SerialNumber'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                    'AssetTag' = 'Asset Tag:'
                    'Version' = 'Version:'
                    'SerialNumber' = 'Serial Number:'
                },
                'Switch {0}'
            )
            'AccessPoint' = [DeviceConfiguration]::new(
                'AccessPoint',
                'AP',
                '.20',
                100,
                10,
                @('ManagementIP', 'Name', 'AssetTag', 'Version', 'SerialNumber'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                    'AssetTag' = 'Asset Tag:'
                    'Version' = 'Version:'
                    'SerialNumber' = 'Serial Number:'
                },
                'Access Point {0}'
            )
            'UPS' = [DeviceConfiguration]::new(
                'UPS',
                'UPS',
                '.102',
                100,
                5,
                @('ManagementIP', 'Name'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                },
                'UPS {0}'
            )
            'CCTV' = [DeviceConfiguration]::new(
                'CCTV',
                'CAM',
                '.110',
                50,
                15,
                @('ManagementIP', 'Name', 'SerialNumber'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                    'SerialNumber' = 'Serial Number:'
                },
                'Camera {0}'
            )
            'Printer' = [DeviceConfiguration]::new(
            'Printer',
            'PRT',
            '.102',
            50,
            6,
            @('ManagementIP', 'Name', 'Model', 'SerialNumber'),
            @{
                'ManagementIP' = 'Management IP:'
                'Name' = 'Name:'
                'Model' = 'Model:'
                'SerialNumber' = 'Serial Number:'
            },
            'Printer {0}'
            )
        }
    }
    
    [void] InitializeUIReferences() {
        $this.StackPanels = @{
            'Switch' = $this.MainWindow.FindName("stkSwitches")
            'AccessPoint' = $this.MainWindow.FindName("stkAccessPoints") 
            'UPS' = $this.MainWindow.FindName("stkUPS")
            'CCTV' = $this.MainWindow.FindName("stkCCTV")
            'Printer' = $this.MainWindow.FindName("stkPrinter")
        }
        
        $this.ComboBoxes = @{
            'Switch' = $this.MainWindow.FindName("cmbSwitchCount")
            'AccessPoint' = $this.MainWindow.FindName("cmbAPCount")
            'UPS' = $this.MainWindow.FindName("cmbUPSCount") 
            'CCTV' = $this.MainWindow.FindName("cmbCCTVCount")
            'Printer' = $this.MainWindow.FindName("cmbPrinterCount")
        }
    }
    
        
    # Universal panel update
    [void] UpdateDevicePanels([string]$deviceType, [int]$count) {
        try {
            $stackPanel = $this.StackPanels[$deviceType]

            if (-not $stackPanel) { return }
            
            # Save existing data
            $existingData = @()
            Write-Host "DEBUG: Got existing data, calling RestoreDeviceData for $deviceType"
            
            # Clear existing panels
            $stackPanel.Children.Clear()
            $stackPanel.RowDefinitions.Clear()
            $stackPanel.ColumnDefinitions.Clear()
            
            if ($count -eq 0) { return }
            
            # Calculate layout
            $numRows = [Math]::Ceiling($count / 2)
            
            # Setup grid layout
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = New-Object System.Windows.GridLength(1, 'Star')
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = New-Object System.Windows.GridLength(1, 'Star')
            $stackPanel.ColumnDefinitions.Add($col1)
            $stackPanel.ColumnDefinitions.Add($col2)
            
            for ($r = 0; $r -lt $numRows; $r++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = New-Object System.Windows.GridLength(1, 'Auto')
                $stackPanel.RowDefinitions.Add($row)
            }
            
            # Create panels
            for ($i = 1; $i -le $count; $i++) {
                Write-Host "DEBUG: Starting loop iteration $i for $deviceType"
                $config = $this.Configurations[$deviceType]

                $controlPrefix = if ($this -is [EditDevicePanelManager]) { "txtEdit" } else { "txt" }
                $panel = [UniversalDevicePanelFactory]::CreateDevicePanel($config, $i, $controlPrefix)
                Write-Host "DEBUG: Created panel $i for $deviceType, panel is null: $($panel -eq $null)"
                if ($panel) {
                    # Position in grid
                    $row = [Math]::Floor(($i - 1) / 2)
                    $col = ($i - 1) % 2
                    
                    [System.Windows.Controls.Grid]::SetRow($panel, $row)
                    [System.Windows.Controls.Grid]::SetColumn($panel, $col)
                    $panel.Margin = New-Object System.Windows.Thickness(0,0,10,10)
                    
                    $stackPanel.Children.Add($panel) | Out-Null
                }
            }
            
            # Use universal data restorer
            $config = $this.Configurations[$deviceType]

            $controlPrefix = if ($this -is [EditDevicePanelManager]) { "txtEdit" } else { "txt" }
            [UniversalDataCollector]::RestoreDeviceData($config, $stackPanel, $existingData, $count, $controlPrefix)
            
        } catch {
            [System.Windows.MessageBox]::Show("Error updating $deviceType panels: $_", "Panel Update Error", "OK", "Error")
        }
    }
    
    # Universal data collection
    [object] GetDeviceDataFromUI([string]$deviceType) {
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $comboBox = $this.ComboBoxes[$deviceType]
        
        # Determine device type class name
        $className = switch ($deviceType) {
            'Switch' { 'SwitchInfo' }
            'AccessPoint' { 'AccessPointInfo' }
            'UPS' { 'UPSInfo' }
            'CCTV' { 'CCTVInfo' }
        }
        
        $devices = New-Object "System.Collections.Generic.List[$className]"
        Write-Host "DEBUG: Creating list for className: '$className', DeviceType: '$($Config.Type)'"
        $deviceCount = if ($comboBox.SelectedItem) { [int]$comboBox.SelectedItem.Content } else { 0 }
        
        for ($i = 1; $i -le $deviceCount; $i++) {
            $device = New-Object $className
            
            foreach ($groupBox in $stackPanel.Children) {
                if ($groupBox.Header -eq ($config.HeaderTemplate -f $i)) {
                    foreach ($field in $config.Fields) {
                        $controlName = "txt$deviceType$i$field"
                        $control = $this.FindControlInPanel($groupBox, $controlName)
                        if ($control) {
                            $device.$field = $control.Text
                        }
                    }
                    break
                }
            }
            $devices.Add($device)
        }
        return $devices
    }

    [void] RestoreDeviceData([string]$deviceType, [array]$existingData, [int]$newCount) {
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $controlPrefix = if ($this -is [EditDevicePanelManager]) { "txtEdit" } else { "txt" }
        [UniversalDataCollector]::RestoreDeviceData($config, $stackPanel, $existingData, $newCount, $controlPrefix)
    }
    
    # Helper method to find control in panel
    [object] FindControlInPanel([object]$groupBox, [string]$controlName) {
        $grid = $groupBox.Content
        foreach ($control in $grid.Children) {
            if ($control.Name -eq $controlName) {
                return $control
            }
        }
        return $null
    }
    
    
    # Universal auto-naming
    [void] UpdateDeviceNamesFromSiteCode([string]$deviceType, [string]$siteCode) {
        if ([string]::IsNullOrWhiteSpace($siteCode)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $siteCode = $siteCode.Trim().ToUpper()
        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = $matches[1]
                $paddedNumber = $deviceNumber.PadLeft(3, '0')
                $deviceName = "$siteCode-$($config.Prefix)-$paddedNumber"
                
                $nameControl = $this.FindControlInPanel($groupBox, "txt$deviceType${deviceNumber}Name")
                if ($nameControl) {
                    $nameControl.Text = $deviceName
                }
            }
        }
    }
    
    # Universal IP auto-population  
    [void] UpdateDeviceIPsFromSubnet([string]$deviceType, [string]$baseSubnet) {
        if ([string]::IsNullOrWhiteSpace($baseSubnet)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = [int]$matches[1]
                $deviceIP = "$baseSubnet$($config.VLANSubnet).$($deviceNumber + $config.IPStartOffset - 1)"
                
                $ipControl = $this.FindControlInPanel($groupBox, "txt$deviceType${deviceNumber}ManagementIP")
                if ($ipControl -and [string]::IsNullOrWhiteSpace($ipControl.Text)) {
                    $ipControl.Text = $deviceIP
                }
            }
        }
    }
}

class EditDevicePanelManager : DevicePanelManager {
    EditDevicePanelManager([object]$editWindow) : base($editWindow) {
        $this.InitializeEditUIReferences()
    }
    
    [void] InitializeEditUIReferences() {
        $this.StackPanels = @{
            'Switch' = $this.MainWindow.FindName("stkEditSwitches")
            'AccessPoint' = $this.MainWindow.FindName("stkEditAccessPoints") 
            'UPS' = $this.MainWindow.FindName("stkEditUPS")
            'CCTV' = $this.MainWindow.FindName("stkEditCCTV")
            'Printer' = $this.MainWindow.FindName("stkEditPrinter")
        }
        
        $this.ComboBoxes = @{
            'Switch' = $this.MainWindow.FindName("cmbEditSwitchCount")
            'AccessPoint' = $this.MainWindow.FindName("cmbEditAPCount")
            'UPS' = $this.MainWindow.FindName("cmbEditUPSCount") 
            'CCTV' = $this.MainWindow.FindName("cmbEditCCTVCount")
            'Printer' = $this.MainWindow.FindName("cmbEditPrinterCount")
        }
    }
    
    # Override UpdateDeviceNamesFromSiteCode to use Edit naming convention
    [void] UpdateDeviceNamesFromSiteCode([string]$deviceType, [string]$siteCode) {
        if ([string]::IsNullOrWhiteSpace($siteCode)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $siteCode = $siteCode.Trim().ToUpper()
        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = $matches[1]
                $paddedNumber = $deviceNumber.PadLeft(3, '0')
                $deviceName = "$siteCode-$($config.Prefix)-$paddedNumber"
                
                # Use Edit naming convention
                $nameControl = $this.FindControlInPanel($groupBox, "txtEdit$deviceType${deviceNumber}Name")
                if ($nameControl) {
                    $nameControl.Text = $deviceName
                }
            }
        }
    }
    
    # Override UpdateDeviceIPsFromSubnet to use Edit naming convention  
    [void] UpdateDeviceIPsFromSubnet([string]$deviceType, [string]$baseSubnet) {
        if ([string]::IsNullOrWhiteSpace($baseSubnet)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = [int]$matches[1]
                $deviceIP = "$baseSubnet$($config.VLANSubnet).$($deviceNumber + $config.IPStartOffset - 1)"
                
                # Use Edit naming convention
                $ipControl = $this.FindControlInPanel($groupBox, "txtEdit$deviceType${deviceNumber}ManagementIP")
                if ($ipControl) {
                $ipControl.Text = $deviceIP
                }
            }
        }
    }
}

class UniversalDevicePanelFactory {
    static [System.Windows.Controls.GroupBox] CreateDevicePanel([DeviceConfiguration]$Config, [int]$DeviceNumber, [string]$ControlPrefix = "") {
        try {
            $groupBox = New-Object System.Windows.Controls.GroupBox
            $groupBox.Header = $Config.HeaderTemplate -f $DeviceNumber
            $groupBox.Margin = New-Object System.Windows.Thickness(0,0,0,10)
            
            $grid = New-Object System.Windows.Controls.Grid
            $grid.Margin = New-Object System.Windows.Thickness(5)
            
            # Create 2 columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = New-Object System.Windows.GridLength(1, 'Auto')
            $col2 = New-Object System.Windows.Controls.ColumnDefinition  
            $col2.Width = New-Object System.Windows.GridLength(1, 'Star')
            $grid.ColumnDefinitions.Add($col1)
            $grid.ColumnDefinitions.Add($col2)
            
            # Create rows for fields
            for ($i = 0; $i -lt $Config.Fields.Count; $i++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = New-Object System.Windows.GridLength(1, 'Auto')
                $grid.RowDefinitions.Add($row)
            }
            
            # Add fields dynamically
            for ($i = 0; $i -lt $Config.Fields.Count; $i++) {
                $field = $Config.Fields[$i]
                $label = $Config.FieldLabels[$field]
                
                # Create label
                $lblControl = New-Object System.Windows.Controls.Label
                $lblControl.Content = $label
                [System.Windows.Controls.Grid]::SetRow($lblControl, $i)
                [System.Windows.Controls.Grid]::SetColumn($lblControl, 0)
                $grid.Children.Add($lblControl) | Out-Null
                
                # Create textbox with configurable prefix
                $txtControl = New-Object System.Windows.Controls.TextBox
                $txtControl.Name = "$ControlPrefix$($Config.Type)$DeviceNumber$field"
                $txtControl.Margin = New-Object System.Windows.Thickness(0,2,0,2)
                [System.Windows.Controls.Grid]::SetRow($txtControl, $i)
                [System.Windows.Controls.Grid]::SetColumn($txtControl, 1)
                $grid.Children.Add($txtControl) | Out-Null
            }
            
            $groupBox.Content = $grid
            return $groupBox
            
        } catch {
            [System.Windows.MessageBox]::Show("Error creating $($Config.Type) panel: $_", "Panel Creation Error", "OK", "Error")
            return $null
        }
    }
}

class DeviceConfiguration {
    [string]$Type
    [string]$Prefix
    [string]$VLANSubnet
    [int]$IPStartOffset
    [int]$MaxCount
    [string[]]$Fields
    [hashtable]$FieldLabels
    [string]$HeaderTemplate
    
    DeviceConfiguration([string]$type, [string]$prefix, [string]$vlanSubnet, [int]$ipOffset, [int]$maxCount, [string[]]$fields, [hashtable]$labels, [string]$headerTemplate) {
        $this.Type = $type
        $this.Prefix = $prefix
        $this.VLANSubnet = $vlanSubnet
        $this.IPStartOffset = $ipOffset
        $this.MaxCount = $maxCount
        $this.Fields = $fields
        $this.FieldLabels = $labels
        $this.HeaderTemplate = $headerTemplate
    }
}

class UniversalDataCollector {
    static [object] CollectDeviceData([DeviceConfiguration]$Config, [object]$StackPanel, [object]$ComboBox, [string]$ControlPrefix = "txt") {
        # Determine device type class name
        $className = switch ($Config.Type) {
            'Switch' { 'SwitchInfo' }
            'AccessPoint' { 'AccessPointInfo' }
            'UPS' { 'UPSInfo' }
            'CCTV' { 'CCTVInfo' }
            'Printer' { 'PrinterInfo' }
        }
        
        $devices = New-Object "System.Collections.Generic.List[$className]"
        Write-Host "DEBUG: Creating list for className: '$className', DeviceType: '$($Config.Type)'"
        $deviceCount = if ($ComboBox.SelectedItem) { [int]$ComboBox.SelectedItem.Content } else { 0 }
        
        for ($i = 1; $i -le $deviceCount; $i++) {
            $device = New-Object $className
            
            foreach ($groupBox in $StackPanel.Children) {
                if ($groupBox.Header -eq ($Config.HeaderTemplate -f $i)) {
                    foreach ($field in $Config.Fields) {
                        $controlName = "$ControlPrefix$($Config.Type)$i$field"
                        $control = [UniversalDataCollector]::FindControlInPanel($groupBox, $controlName)
                        if ($control) {
                            $device.$field = $control.Text.Trim()
                        }
                    }
                    break
                }
            }
            $devices.Add($device)
        }
        return $devices
    }
    
    # Helper method to find control in panel
    static [object] FindControlInPanel([object]$GroupBox, [string]$ControlName) {
        $grid = $GroupBox.Content
        foreach ($control in $grid.Children) {
            if ($control.Name -eq $ControlName) {
                return $control
            }
        }
        return $null
    }
    
    # Universal data populator
    static [void] PopulateDevicePanels([DeviceConfiguration]$Config, [object]$StackPanel, [array]$DeviceList, [string]$ControlPrefix = "txt") {
        if (-not $DeviceList -or $DeviceList.Count -eq 0) { return }
        
        for ($i = 0; $i -lt $DeviceList.Count; $i++) {
            $deviceNum = $i + 1
            $device = $DeviceList[$i]
            
            foreach ($groupBox in $StackPanel.Children) {
                if ($groupBox.Header -eq ($Config.HeaderTemplate -f $deviceNum)) {
                    foreach ($field in $Config.Fields) {
                        $controlName = "$ControlPrefix$($Config.Type)$deviceNum$field"
                        $control = [UniversalDataCollector]::FindControlInPanel($groupBox, $controlName)
                        if ($control -and $device.$field) {
                            $control.Text = $device.$field
                        }
                    }
                    break
                }
            }
        }
    }
    
    # Universal data restoration
    static [void] RestoreDeviceData([DeviceConfiguration]$Config, [object]$StackPanel, [array]$ExistingData, [int]$NewCount, [string]$ControlPrefix = "txt") {
        if (-not $ExistingData) { return }
        
        $maxRestore = [Math]::Min($ExistingData.Count, $NewCount)
        
        for ($i = 0; $i -lt $maxRestore; $i++) {
            $deviceNum = $i + 1
            $deviceData = $ExistingData[$i]
            
            foreach ($groupBox in $StackPanel.Children) {
                if ($groupBox.Header -eq ($Config.HeaderTemplate -f $deviceNum)) {
                    foreach ($field in $Config.Fields) {
                        $controlName = "$ControlPrefix$($Config.Type)$deviceNum$field"
                        $control = [UniversalDataCollector]::FindControlInPanel($groupBox, $controlName)
                        if ($control -and $deviceData.$field) {
                            $control.Text = $deviceData.$field
                        }
                    }
                    break
                }
            }
        }
    }
}

function Handle-DeviceCountChanged {
    param(
        [string]$DeviceType,
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    
    if ($ComboBox.SelectedItem) {
        $count = [int]$ComboBox.SelectedItem.Content
        $script:DeviceManager.UpdateDevicePanels($DeviceType, $count)
        Write-Host "DEBUG: About to update panels for DeviceType: '$DeviceType', count: $count"
        
        # Auto-populate names and IPs if site code/subnet exists
        if (-not [string]::IsNullOrWhiteSpace($txtSiteCode.Text)) {
            $script:DeviceManager.UpdateDeviceNamesFromSiteCode($DeviceType, $txtSiteCode.Text)
        }
        if (-not [string]::IsNullOrWhiteSpace($txtSiteSubnet.Text)) {
            if ($txtSiteSubnet.Text -match '^(\d+\.\d+)\.') {
                $script:DeviceManager.UpdateDeviceIPsFromSubnet($DeviceType, $matches[1])
            }
        }
    }
}