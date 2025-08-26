function Show-SiteDetails {
    param([SiteEntry]$Site)
    
    try {
        $stkSiteDetails.Items.Clear()
        
        # Create 10 TabItems
        $basicTab = New-Object System.Windows.Controls.TabItem
        $basicTab.Header = "Basic Info"
        
        $switchTab = New-Object System.Windows.Controls.TabItem
        $switchTab.Header = "Switches"
        
        $apTab = New-Object System.Windows.Controls.TabItem
        $apTab.Header = "Access Points"
        
        $firewallTab = New-Object System.Windows.Controls.TabItem
        $firewallTab.Header = "Firewall"
        
        $primaryTab = New-Object System.Windows.Controls.TabItem
        $primaryTab.Header = "Primary Circuit"
        
        $backupTab = New-Object System.Windows.Controls.TabItem
        $backupTab.Header = "Backup Circuit"
        
        $vlanTab = New-Object System.Windows.Controls.TabItem
        $vlanTab.Header = "VLANs"
        
        $cctvTab = New-Object System.Windows.Controls.TabItem
        $cctvTab.Header = "CCTV"
        
        $upsTab = New-Object System.Windows.Controls.TabItem
        $upsTab.Header = "UPS"
        
        $printerTab = New-Object System.Windows.Controls.TabItem
        $printerTab.Header = "Printer"
        
        # ============================================================================
        # 1. BASIC INFO TAB
        # ============================================================================
        $basicContent = New-Object System.Windows.Controls.ScrollViewer
        $basicContent.VerticalScrollBarVisibility = "Auto"
        $basicGrid = New-Object System.Windows.Controls.Grid
        $basicGrid.Margin = "10"
        
        # Create columns for basic info
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $basicGrid.ColumnDefinitions.Add($col1)
        $basicGrid.ColumnDefinitions.Add($col2)
        
        # Basic info fields
        $basicInfoFields = @(
            @("Site Code:", $Site.SiteCode),
            @("Site Subnet:", $Site.SiteSubnet),
            @("Site Subnet Code:", $Site.SiteSubnetCode),
            @("Site Name:", $Site.SiteName),
            @("Site Address:", $Site.SiteAddress),
            @("Main Contact Name:", $Site.MainContactName),
            @("Main Contact Phone:", $Site.MainContactPhone),
            @("Second Contact Name:", $Site.SecondContactName),
            @("Second Contact Phone:", $Site.SecondContactPhone)
        )
        
        for ($i = 0; $i -lt $basicInfoFields.Count; $i++) {
            $row = New-Object System.Windows.Controls.RowDefinition
            $row.Height = "Auto"
            $basicGrid.RowDefinitions.Add($row)
            
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $basicInfoFields[$i][0]
            $label.FontWeight = "Bold"
            [System.Windows.Controls.Grid]::SetRow($label, $i)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            $basicGrid.Children.Add($label)
            
            $clickableText = New-ClickableText -Value $basicInfoFields[$i][1]
            $clickableText.Margin = "10,5,0,5"
            [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
            [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
            $basicGrid.Children.Add($clickableText)
        }
        
        $basicContent.Content = $basicGrid
        $basicTab.Content = $basicContent
        
        # ============================================================================
        # 2. SWITCHES TAB
        # ============================================================================
        $switchContent = New-Object System.Windows.Controls.ScrollViewer
        $switchContent.VerticalScrollBarVisibility = "Auto"
        $switchGrid = New-Object System.Windows.Controls.Grid
        $switchGrid.Margin = "10"
        
        # Count actual switches with data
        $actualSwitchCount = 0
        foreach ($switch in $Site.Switches) {
            if (-not [string]::IsNullOrWhiteSpace($switch.ManagementIP) -or 
                -not [string]::IsNullOrWhiteSpace($switch.Name) -or 
                -not [string]::IsNullOrWhiteSpace($switch.AssetTag) -or 
                -not [string]::IsNullOrWhiteSpace($switch.Version) -or 
                -not [string]::IsNullOrWhiteSpace($switch.SerialNumber)) {
                $actualSwitchCount++
            }
        }
        
        if ($Site.Switches.Count -eq 0 -or $actualSwitchCount -eq 0) {
            $noSwitchText = New-Object System.Windows.Controls.TextBlock
            $noSwitchText.Text = "No switches configured for this site."
            $noSwitchText.FontStyle = "Italic"
            $noSwitchText.Foreground = "Gray"
            $noSwitchText.Margin = "10"
            $switchGrid.Children.Add($noSwitchText)
        } else {
            for ($i = 0; $i -lt $Site.Switches.Count; $i++) {
                $switch = $Site.Switches[$i]
                
                # Create switch header
                $switchHeader = New-Object System.Windows.Controls.TextBlock
                $switchHeader.Text = "Switch $($i+1):"
                $switchHeader.FontWeight = "Bold"
                $switchHeader.FontSize = 14
                $switchHeader.Margin = "0,10,0,5"
                
                # Create switch details grid
                $switchDetailGrid = New-Object System.Windows.Controls.Grid
                $switchDetailGrid.Margin = "20,0,0,10"
                
                $col1 = New-Object System.Windows.Controls.ColumnDefinition
                $col1.Width = "Auto"
                $col2 = New-Object System.Windows.Controls.ColumnDefinition
                $col2.Width = "*"
                $switchDetailGrid.ColumnDefinitions.Add($col1)
                $switchDetailGrid.ColumnDefinitions.Add($col2)
                
                $switchDetails = @(
                    @("Management IP:", $(if ($switch.ManagementIP) { $switch.ManagementIP } else { '(Not specified)' })),
                    @("Name:", $(if ($switch.Name) { $switch.Name } else { '(Not specified)' })),
                    @("Asset Tag:", $(if ($switch.AssetTag) { $switch.AssetTag } else { '(Not specified)' })),
                    @("Version:", $(if ($switch.Version) { $switch.Version } else { '(Not specified)' })),
                    @("Serial Number:", $(if ($switch.SerialNumber) { $switch.SerialNumber } else { '(Not specified)' }))
                )
                
                for ($j = 0; $j -lt $switchDetails.Count; $j++) {
                    $row = New-Object System.Windows.Controls.RowDefinition
                    $row.Height = "Auto"
                    $switchDetailGrid.RowDefinitions.Add($row)
                    
                    $label = New-Object System.Windows.Controls.Label
                    $label.Content = $switchDetails[$j][0]
                    $label.FontWeight = "Bold"
                    [System.Windows.Controls.Grid]::SetRow($label, $j)
                    [System.Windows.Controls.Grid]::SetColumn($label, 0)
                    $switchDetailGrid.Children.Add($label)
                    
                    $clickableText = New-ClickableText -Value $switchDetails[$j][1]
                    $clickableText.Margin = "10,5,0,5"
                    [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
                    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                    $switchDetailGrid.Children.Add($clickableText)
                }
                
                # Add to main grid
                $headerRow = New-Object System.Windows.Controls.RowDefinition
                $headerRow.Height = "Auto"
                $detailRow = New-Object System.Windows.Controls.RowDefinition
                $detailRow.Height = "Auto"
                $switchGrid.RowDefinitions.Add($headerRow)
                $switchGrid.RowDefinitions.Add($detailRow)
                
                [System.Windows.Controls.Grid]::SetRow($switchHeader, $switchGrid.RowDefinitions.Count - 2)
                [System.Windows.Controls.Grid]::SetRow($switchDetailGrid, $switchGrid.RowDefinitions.Count - 1)
                $switchGrid.Children.Add($switchHeader)
                $switchGrid.Children.Add($switchDetailGrid)
            }
        }
        
        $switchContent.Content = $switchGrid
        $switchTab.Content = $switchContent
        
        # ============================================================================
        # 3. ACCESS POINTS TAB
        # ============================================================================
        $apContent = New-Object System.Windows.Controls.ScrollViewer
        $apContent.VerticalScrollBarVisibility = "Auto"
        $apGrid = New-Object System.Windows.Controls.Grid
        $apGrid.Margin = "10"
        
        # Count actual APs with data
        $actualAPCount = 0
        foreach ($ap in $Site.AccessPoints) {
            if (-not [string]::IsNullOrWhiteSpace($ap.ManagementIP) -or 
                -not [string]::IsNullOrWhiteSpace($ap.Name) -or 
                -not [string]::IsNullOrWhiteSpace($ap.AssetTag) -or 
                -not [string]::IsNullOrWhiteSpace($ap.Version) -or 
                -not [string]::IsNullOrWhiteSpace($ap.SerialNumber)) {
                $actualAPCount++
            }
        }
        
        if ($Site.AccessPoints.Count -eq 0 -or $actualAPCount -eq 0) {
            $noAPText = New-Object System.Windows.Controls.TextBlock
            $noAPText.Text = "No access points configured for this site."
            $noAPText.FontStyle = "Italic"
            $noAPText.Foreground = "Gray"
            $noAPText.Margin = "10"
            $apGrid.Children.Add($noAPText)
        } else {
            for ($i = 0; $i -lt $Site.AccessPoints.Count; $i++) {
                $ap = $Site.AccessPoints[$i]
                
                # Create AP header
                $apHeader = New-Object System.Windows.Controls.TextBlock
                $apHeader.Text = "Access Point $($i+1):"
                $apHeader.FontWeight = "Bold"
                $apHeader.FontSize = 14
                $apHeader.Margin = "0,10,0,5"
                
                # Create AP details grid
                $apDetailGrid = New-Object System.Windows.Controls.Grid
                $apDetailGrid.Margin = "20,0,0,10"
                
                $col1 = New-Object System.Windows.Controls.ColumnDefinition
                $col1.Width = "Auto"
                $col2 = New-Object System.Windows.Controls.ColumnDefinition
                $col2.Width = "*"
                $apDetailGrid.ColumnDefinitions.Add($col1)
                $apDetailGrid.ColumnDefinitions.Add($col2)
                
                $apDetails = @(
                    @("Management IP:", $(if ($ap.ManagementIP) { $ap.ManagementIP } else { '(Not specified)' })),
                    @("Name:", $(if ($ap.Name) { $ap.Name } else { '(Not specified)' })),
                    @("Asset Tag:", $(if ($ap.AssetTag) { $ap.AssetTag } else { '(Not specified)' })),
                    @("Version:", $(if ($ap.Version) { $ap.Version } else { '(Not specified)' })),
                    @("Serial Number:", $(if ($ap.SerialNumber) { $ap.SerialNumber } else { '(Not specified)' }))
                )
                
                for ($j = 0; $j -lt $apDetails.Count; $j++) {
                    $row = New-Object System.Windows.Controls.RowDefinition
                    $row.Height = "Auto"
                    $apDetailGrid.RowDefinitions.Add($row)
                    
                    $label = New-Object System.Windows.Controls.Label
                    $label.Content = $apDetails[$j][0]
                    $label.FontWeight = "Bold"
                    [System.Windows.Controls.Grid]::SetRow($label, $j)
                    [System.Windows.Controls.Grid]::SetColumn($label, 0)
                    $apDetailGrid.Children.Add($label)
                    
                    $clickableText = New-ClickableText -Value $apDetails[$j][1]
                    $clickableText.Margin = "10,5,0,5"
                    [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
                    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                    $apDetailGrid.Children.Add($clickableText)
                }
                
                # Add to main grid
                $headerRow = New-Object System.Windows.Controls.RowDefinition
                $headerRow.Height = "Auto"
                $detailRow = New-Object System.Windows.Controls.RowDefinition
                $detailRow.Height = "Auto"
                $apGrid.RowDefinitions.Add($headerRow)
                $apGrid.RowDefinitions.Add($detailRow)
                
                [System.Windows.Controls.Grid]::SetRow($apHeader, $apGrid.RowDefinitions.Count - 2)
                [System.Windows.Controls.Grid]::SetRow($apDetailGrid, $apGrid.RowDefinitions.Count - 1)
                $apGrid.Children.Add($apHeader)
                $apGrid.Children.Add($apDetailGrid)
            }
        }
        
        $apContent.Content = $apGrid
        $apTab.Content = $apContent
        
        # ============================================================================
        # 4. FIREWALL TAB
        # ============================================================================
        $firewallContent = New-Object System.Windows.Controls.ScrollViewer
        $firewallContent.VerticalScrollBarVisibility = "Auto"
        $firewallGrid = New-Object System.Windows.Controls.Grid
        $firewallGrid.Margin = "10"
        
        # Create columns
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $firewallGrid.ColumnDefinitions.Add($col1)
        $firewallGrid.ColumnDefinitions.Add($col2)
        
        # Check if firewall has any information
        $hasFirewallInfo = (-not [string]::IsNullOrWhiteSpace($Site.FirewallIP)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.FirewallName)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.FirewallVersion)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.FirewallSN))
        
        if (-not $hasFirewallInfo) {
            $noFirewallText = New-Object System.Windows.Controls.TextBlock
            $noFirewallText.Text = "No firewall information configured for this site."
            $noFirewallText.FontStyle = "Italic"
            $noFirewallText.Foreground = "Gray"
            $noFirewallText.Margin = "10"
            $firewallGrid.Children.Add($noFirewallText)
        } else {
            $firewallFields = @(
                @("Management IP:", $(if ($Site.FirewallIP) { $Site.FirewallIP } else { '(Not specified)' })),
                @("Name:", $(if ($Site.FirewallName) { $Site.FirewallName } else { '(Not specified)' })),
                @("Version:", $(if ($Site.FirewallVersion) { $Site.FirewallVersion } else { '(Not specified)' })),
                @("Serial Number:", $(if ($Site.FirewallSN) { $Site.FirewallSN } else { '(Not specified)' }))
            )
            
            for ($i = 0; $i -lt $firewallFields.Count; $i++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = "Auto"
                $firewallGrid.RowDefinitions.Add($row)
                
                $label = New-Object System.Windows.Controls.Label
                $label.Content = $firewallFields[$i][0]
                $label.FontWeight = "Bold"
                [System.Windows.Controls.Grid]::SetRow($label, $i)
                [System.Windows.Controls.Grid]::SetColumn($label, 0)
                $firewallGrid.Children.Add($label)
                
                $clickableText = New-ClickableText -Value $firewallFields[$i][1]
                $clickableText.Margin = "10,5,0,5"
                [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
                [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                $firewallGrid.Children.Add($clickableText)
            }
        }
        
        $firewallContent.Content = $firewallGrid
        $firewallTab.Content = $firewallContent
        
        # ============================================================================
        # 5. PRIMARY CIRCUIT TAB
        # ============================================================================
        $primaryContent = New-Object System.Windows.Controls.ScrollViewer
        $primaryContent.VerticalScrollBarVisibility = "Auto"
        $primaryGrid = New-Object System.Windows.Controls.Grid
        $primaryGrid.Margin = "10"
        
        # Create columns
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $primaryGrid.ColumnDefinitions.Add($col1)
        $primaryGrid.ColumnDefinitions.Add($col2)
        
        # Check if primary circuit has any information
        $hasPrimaryInfo = (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.Vendor)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.CircuitType)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.CircuitID)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DownloadSpeed)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.UploadSpeed)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.IPAddress)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.SubnetMask)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DefaultGateway)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DNS1)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DNS2)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.RouterModel)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.RouterName)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.RouterSN))
        
        if (-not $hasPrimaryInfo) {
            $noPrimaryText = New-Object System.Windows.Controls.TextBlock
            $noPrimaryText.Text = "No primary circuit information configured for this site."
            $noPrimaryText.FontStyle = "Italic"
            $noPrimaryText.Foreground = "Gray"
            $noPrimaryText.Margin = "10"
            $primaryGrid.Children.Add($noPrimaryText)
        } else {
            $primaryFields = @(
                @("Vendor:", $(if ($Site.PrimaryCircuit.Vendor) { $Site.PrimaryCircuit.Vendor } else { '(Not specified)' })),
                @("Circuit Type:", $(if ($Site.PrimaryCircuit.CircuitType) { $Site.PrimaryCircuit.CircuitType } else { '(Not specified)' })),
                @("Circuit ID:", $(if ($Site.PrimaryCircuit.CircuitID) { $Site.PrimaryCircuit.CircuitID } else { '(Not specified)' })),
                @("Download Speed:", $(if ($Site.PrimaryCircuit.DownloadSpeed) { $Site.PrimaryCircuit.DownloadSpeed } else { '(Not specified)' })),
                @("Upload Speed:", $(if ($Site.PrimaryCircuit.UploadSpeed) { $Site.PrimaryCircuit.UploadSpeed } else { '(Not specified)' })),
                @("IP Address:", $(if ($Site.PrimaryCircuit.IPAddress) { $Site.PrimaryCircuit.IPAddress } else { '(Not specified)' })),
                @("Subnet Mask:", $(if ($Site.PrimaryCircuit.SubnetMask) { $Site.PrimaryCircuit.SubnetMask } else { '(Not specified)' })),
                @("Default Gateway:", $(if ($Site.PrimaryCircuit.DefaultGateway) { $Site.PrimaryCircuit.DefaultGateway } else { '(Not specified)' })),
                @("DNS 1:", $(if ($Site.PrimaryCircuit.DNS1) { $Site.PrimaryCircuit.DNS1 } else { '(Not specified)' })),
                @("DNS 2:", $(if ($Site.PrimaryCircuit.DNS2) { $Site.PrimaryCircuit.DNS2 } else { '(Not specified)' })),
                @("Router Model:", $(if ($Site.PrimaryCircuit.RouterModel) { $Site.PrimaryCircuit.RouterModel } else { '(Not specified)' })),
                @("Router Name:", $(if ($Site.PrimaryCircuit.RouterName) { $Site.PrimaryCircuit.RouterName } else { '(Not specified)' })),
                @("Router Serial Number:", $(if ($Site.PrimaryCircuit.RouterSN) { $Site.PrimaryCircuit.RouterSN } else { '(Not specified)' }))
            )
            
            # Add GPON fields if applicable
            if ($Site.PrimaryCircuit.CircuitType -eq "GPON Fiber") {
                $primaryFields += @(
                    @("PPPoE Username:", $(if ($Site.PrimaryCircuit.PPPoEUsername) { $Site.PrimaryCircuit.PPPoEUsername } else { '(Not specified)' })),
                    @("PPPoE Password:", $(if ($Site.PrimaryCircuit.PPPoEPassword) { $Site.PrimaryCircuit.PPPoEPassword } else { '(Not specified)' }))
                )
            }
            
            # Add modem fields if applicable
            if ($Site.PrimaryCircuit.HasModem) {
                $primaryFields += @(
                    @("Modem Model:", $(if ($Site.PrimaryCircuit.ModemModel) { $Site.PrimaryCircuit.ModemModel } else { '(Not specified)' })),
                    @("Modem Name:", $(if ($Site.PrimaryCircuit.ModemName) { $Site.PrimaryCircuit.ModemName } else { '(Not specified)' })),
                    @("Modem Serial Number:", $(if ($Site.PrimaryCircuit.ModemSN) { $Site.PrimaryCircuit.ModemSN } else { '(Not specified)' }))
                )
            }
            
            for ($i = 0; $i -lt $primaryFields.Count; $i++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = "Auto"
                $primaryGrid.RowDefinitions.Add($row)
                
                $label = New-Object System.Windows.Controls.Label
                $label.Content = $primaryFields[$i][0]
                $label.FontWeight = "Bold"
                [System.Windows.Controls.Grid]::SetRow($label, $i)
                [System.Windows.Controls.Grid]::SetColumn($label, 0)
                $primaryGrid.Children.Add($label)
                
                $clickableText = New-ClickableText -Value $primaryFields[$i][1]
                $clickableText.Margin = "10,5,0,5"
                [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
                [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                $primaryGrid.Children.Add($clickableText)
            }
        }
        
        $primaryContent.Content = $primaryGrid
        $primaryTab.Content = $primaryContent
        
        # ============================================================================
        # 6. BACKUP CIRCUIT TAB
        # ============================================================================
        $backupContent = New-Object System.Windows.Controls.ScrollViewer
        $backupContent.VerticalScrollBarVisibility = "Auto"
        $backupGrid = New-Object System.Windows.Controls.Grid
        $backupGrid.Margin = "10"
        
        if (-not $Site.HasBackupCircuit) {
            $noBackupText = New-Object System.Windows.Controls.TextBlock
            $noBackupText.Text = "No backup circuit configured for this site."
            $noBackupText.FontStyle = "Italic"
            $noBackupText.Foreground = "Gray"
            $noBackupText.Margin = "10"
            $backupGrid.Children.Add($noBackupText)
        } else {
            # Create columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = "Auto"
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = "*"
            $backupGrid.ColumnDefinitions.Add($col1)
            $backupGrid.ColumnDefinitions.Add($col2)
            
            $backupFields = @(
                @("Vendor:", $(if ($Site.BackupCircuit.Vendor) { $Site.BackupCircuit.Vendor } else { '(Not specified)' })),
                @("Circuit Type:", $(if ($Site.BackupCircuit.CircuitType) { $Site.BackupCircuit.CircuitType } else { '(Not specified)' })),
                @("Circuit ID:", $(if ($Site.BackupCircuit.CircuitID) { $Site.BackupCircuit.CircuitID } else { '(Not specified)' })),
                @("Download Speed:", $(if ($Site.BackupCircuit.DownloadSpeed) { $Site.BackupCircuit.DownloadSpeed } else { '(Not specified)' })),
                @("Upload Speed:", $(if ($Site.BackupCircuit.UploadSpeed) { $Site.BackupCircuit.UploadSpeed } else { '(Not specified)' })),
                @("IP Address:", $(if ($Site.BackupCircuit.IPAddress) { $Site.BackupCircuit.IPAddress } else { '(Not specified)' })),
                @("Subnet Mask:", $(if ($Site.BackupCircuit.SubnetMask) { $Site.BackupCircuit.SubnetMask } else { '(Not specified)' })),
                @("Default Gateway:", $(if ($Site.BackupCircuit.DefaultGateway) { $Site.BackupCircuit.DefaultGateway } else { '(Not specified)' })),
                @("DNS 1:", $(if ($Site.BackupCircuit.DNS1) { $Site.BackupCircuit.DNS1 } else { '(Not specified)' })),
                @("DNS 2:", $(if ($Site.BackupCircuit.DNS2) { $Site.BackupCircuit.DNS2 } else { '(Not specified)' })),
                @("Router Model:", $(if ($Site.BackupCircuit.RouterModel) { $Site.BackupCircuit.RouterModel } else { '(Not specified)' })),
                @("Router Name:", $(if ($Site.BackupCircuit.RouterName) { $Site.BackupCircuit.RouterName } else { '(Not specified)' })),
                @("Router Serial Number:", $(if ($Site.BackupCircuit.RouterSN) { $Site.BackupCircuit.RouterSN } else { '(Not specified)' }))
            )
            
            # Add GPON fields if applicable
            if ($Site.BackupCircuit.CircuitType -eq "GPON Fiber") {
                $backupFields += @(
                    @("PPPoE Username:", $(if ($Site.BackupCircuit.PPPoEUsername) { $Site.BackupCircuit.PPPoEUsername } else { '(Not specified)' })),
                    @("PPPoE Password:", $(if ($Site.BackupCircuit.PPPoEPassword) { $Site.BackupCircuit.PPPoEPassword } else { '(Not specified)' }))
                )
            }
            
            # Add modem fields if applicable
            if ($Site.BackupCircuit.HasModem) {
                $backupFields += @(
                    @("Modem Model:", $(if ($Site.BackupCircuit.ModemModel) { $Site.BackupCircuit.ModemModel } else { '(Not specified)' })),
                    @("Modem Name:", $(if ($Site.BackupCircuit.ModemName) { $Site.BackupCircuit.ModemName } else { '(Not specified)' })),
                    @("Modem Serial Number:", $(if ($Site.BackupCircuit.ModemSN) { $Site.BackupCircuit.ModemSN } else { '(Not specified)' }))
                )
            }
            
            for ($i = 0; $i -lt $backupFields.Count; $i++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = "Auto"
                $backupGrid.RowDefinitions.Add($row)
                
                $label = New-Object System.Windows.Controls.Label
                $label.Content = $backupFields[$i][0]
                $label.FontWeight = "Bold"
                [System.Windows.Controls.Grid]::SetRow($label, $i)
                [System.Windows.Controls.Grid]::SetColumn($label, 0)
                $backupGrid.Children.Add($label)
                
                $clickableText = New-ClickableText -Value $backupFields[$i][1]
                $clickableText.Margin = "10,5,0,5"
                [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
                [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                $backupGrid.Children.Add($clickableText)
            }
        }
        
        $backupContent.Content = $backupGrid
        $backupTab.Content = $backupContent
        
        # ============================================================================
        # 7. VLANs TAB
        # ============================================================================
        $vlanContent = New-Object System.Windows.Controls.ScrollViewer
        $vlanContent.VerticalScrollBarVisibility = "Auto"
        $vlanGrid = New-Object System.Windows.Controls.Grid
        $vlanGrid.Margin = "10"
        
        # Create columns
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $vlanGrid.ColumnDefinitions.Add($col1)
        $vlanGrid.ColumnDefinitions.Add($col2)
        
        # Check if VLANs have any information
        $hasVLANInfo = (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN100_Servers)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN101_NetworkDevices)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN102_UserDevices)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN103_UserDevices2)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN104_VOIP)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN105_WiFiCorp)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN106_WiFiBYOD)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN107_WiFiGuest)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN108_Spare)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN109_DMZ)) -or
                    (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN110_CCTV))
        
        if (-not $hasVLANInfo) {
            $noVLANText = New-Object System.Windows.Controls.TextBlock
            $noVLANText.Text = "No VLAN information configured for this site."
            $noVLANText.FontStyle = "Italic"
            $noVLANText.Foreground = "Gray"
            $noVLANText.Margin = "10"
            $vlanGrid.Children.Add($noVLANText)
        } else {
            $vlanFields = @(
                @("VLAN 100 - Servers:", $Site.VLANs.VLAN100_Servers),
                @("VLAN 101 - Network:", $Site.VLANs.VLAN101_NetworkDevices),
                @("VLAN 102 - User 1:", $Site.VLANs.VLAN102_UserDevices),
                @("VLAN 103 - User 2:", $Site.VLANs.VLAN103_UserDevices2),
                @("VLAN 104 - VOIP:", $Site.VLANs.VLAN104_VOIP),
                @("VLAN 105 - Wi-Fi Corp:", $Site.VLANs.VLAN105_WiFiCorp),
                @("VLAN 106 - Wi-Fi BYOD:", $Site.VLANs.VLAN106_WiFiBYOD),
                @("VLAN 107 - Wi-Fi Guest:", $Site.VLANs.VLAN107_WiFiGuest),
                @("VLAN 108 - Spare:", $Site.VLANs.VLAN108_Spare),
                @("VLAN 109 - DMZ:", $Site.VLANs.VLAN109_DMZ),
                @("VLAN 110 - CCTV:", $Site.VLANs.VLAN110_CCTV)
            )
            
            for ($i = 0; $i -lt $vlanFields.Count; $i++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = "Auto"
                $vlanGrid.RowDefinitions.Add($row)
                
                $label = New-Object System.Windows.Controls.Label
                $label.Content = $vlanFields[$i][0]
                $label.FontWeight = "Bold"
                [System.Windows.Controls.Grid]::SetRow($label, $i)
                [System.Windows.Controls.Grid]::SetColumn($label, 0)
                $vlanGrid.Children.Add($label)
                
                $clickableText = New-ClickableText -Value $vlanFields[$i][1]
                $clickableText.Margin = "10,5,0,5"
                [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
                [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                $vlanGrid.Children.Add($clickableText)
            }
        }
        
        $vlanContent.Content = $vlanGrid
        $vlanTab.Content = $vlanContent
        
# ============================================================================
        # 8. CCTV TAB
        # ============================================================================
        $cctvContent = New-Object System.Windows.Controls.ScrollViewer
        $cctvContent.VerticalScrollBarVisibility = "Auto"
        $cctvGrid = New-Object System.Windows.Controls.Grid
        $cctvGrid.Margin = "10"
        
        # Count actual CCTV with data
        $actualCCTVCount = 0
        foreach ($cctv in $Site.CCTVDevices) {
            if (-not [string]::IsNullOrWhiteSpace($cctv.ManagementIP) -or 
                -not [string]::IsNullOrWhiteSpace($cctv.Name) -or 
                -not [string]::IsNullOrWhiteSpace($cctv.SerialNumber)) {
                $actualCCTVCount++
            }
        }
        
        if ($Site.CCTVDevices.Count -eq 0 -or $actualCCTVCount -eq 0) {
            $noCCTVText = New-Object System.Windows.Controls.TextBlock
            $noCCTVText.Text = "No CCTV cameras configured for this site."
            $noCCTVText.FontStyle = "Italic"
            $noCCTVText.Foreground = "Gray"
            $noCCTVText.Margin = "10"
            $cctvGrid.Children.Add($noCCTVText)
        } else {
            for ($i = 0; $i -lt $Site.CCTVDevices.Count; $i++) {
                $cctv = $Site.CCTVDevices[$i]
                
                # Create CCTV header
                $cctvHeader = New-Object System.Windows.Controls.TextBlock
                $cctvHeader.Text = "Camera $($i+1):"
                $cctvHeader.FontWeight = "Bold"
                $cctvHeader.FontSize = 14
                $cctvHeader.Margin = "0,10,0,5"
                
                # Create CCTV details grid
                $cctvDetailGrid = New-Object System.Windows.Controls.Grid
                $cctvDetailGrid.Margin = "20,0,0,10"
                
                $col1 = New-Object System.Windows.Controls.ColumnDefinition
                $col1.Width = "Auto"
                $col2 = New-Object System.Windows.Controls.ColumnDefinition
                $col2.Width = "*"
                $cctvDetailGrid.ColumnDefinitions.Add($col1)
                $cctvDetailGrid.ColumnDefinitions.Add($col2)
                
                $cctvDetails = @(
                    @("Management IP:", $(if ($cctv.ManagementIP) { $cctv.ManagementIP } else { '(Not specified)' })),
                    @("Name:", $(if ($cctv.Name) { $cctv.Name } else { '(Not specified)' })),
                    @("Serial Number:", $(if ($cctv.SerialNumber) { $cctv.SerialNumber } else { '(Not specified)' }))
                )
                
                for ($j = 0; $j -lt $cctvDetails.Count; $j++) {
                    $row = New-Object System.Windows.Controls.RowDefinition
                    $row.Height = "Auto"
                    $cctvDetailGrid.RowDefinitions.Add($row)
                    
                    $label = New-Object System.Windows.Controls.Label
                    $label.Content = $cctvDetails[$j][0]
                    $label.FontWeight = "Bold"
                    [System.Windows.Controls.Grid]::SetRow($label, $j)
                    [System.Windows.Controls.Grid]::SetColumn($label, 0)
                    $cctvDetailGrid.Children.Add($label)
                    
                    $clickableText = New-ClickableText -Value $cctvDetails[$j][1]
                    $clickableText.Margin = "10,5,0,5"
                    [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
                    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                    $cctvDetailGrid.Children.Add($clickableText)
                }
                
                # Add to main grid
                $headerRow = New-Object System.Windows.Controls.RowDefinition
                $headerRow.Height = "Auto"
                $detailRow = New-Object System.Windows.Controls.RowDefinition
                $detailRow.Height = "Auto"
                $cctvGrid.RowDefinitions.Add($headerRow)
                $cctvGrid.RowDefinitions.Add($detailRow)
                
                [System.Windows.Controls.Grid]::SetRow($cctvHeader, $cctvGrid.RowDefinitions.Count - 2)
                [System.Windows.Controls.Grid]::SetRow($cctvDetailGrid, $cctvGrid.RowDefinitions.Count - 1)
                $cctvGrid.Children.Add($cctvHeader)
                $cctvGrid.Children.Add($cctvDetailGrid)
            }
        }
        
        $cctvContent.Content = $cctvGrid
        $cctvTab.Content = $cctvContent
        
        # ============================================================================
        # 9. UPS TAB
        # ============================================================================
        $upsContent = New-Object System.Windows.Controls.ScrollViewer
        $upsContent.VerticalScrollBarVisibility = "Auto"
        $upsGrid = New-Object System.Windows.Controls.Grid
        $upsGrid.Margin = "10"
        
        # Count actual UPS with data
        $actualUPSCount = 0
        foreach ($ups in $Site.UPSDevices) {
            if (-not [string]::IsNullOrWhiteSpace($ups.ManagementIP) -or 
                -not [string]::IsNullOrWhiteSpace($ups.Name)) {
                $actualUPSCount++
            }
        }
        
        if ($Site.UPSDevices.Count -eq 0 -or $actualUPSCount -eq 0) {
            $noUPSText = New-Object System.Windows.Controls.TextBlock
            $noUPSText.Text = "No UPS devices configured for this site."
            $noUPSText.FontStyle = "Italic"
            $noUPSText.Foreground = "Gray"
            $noUPSText.Margin = "10"
            $upsGrid.Children.Add($noUPSText)
        } else {
            for ($i = 0; $i -lt $Site.UPSDevices.Count; $i++) {
                $ups = $Site.UPSDevices[$i]
                
                # Create UPS header
                $upsHeader = New-Object System.Windows.Controls.TextBlock
                $upsHeader.Text = "UPS $($i+1):"
                $upsHeader.FontWeight = "Bold"
                $upsHeader.FontSize = 14
                $upsHeader.Margin = "0,10,0,5"
                
                # Create UPS details grid
                $upsDetailGrid = New-Object System.Windows.Controls.Grid
                $upsDetailGrid.Margin = "20,0,0,10"
                
                $col1 = New-Object System.Windows.Controls.ColumnDefinition
                $col1.Width = "Auto"
                $col2 = New-Object System.Windows.Controls.ColumnDefinition
                $col2.Width = "*"
                $upsDetailGrid.ColumnDefinitions.Add($col1)
                $upsDetailGrid.ColumnDefinitions.Add($col2)
                
                $upsDetails = @(
                    @("Management IP:", $(if ($ups.ManagementIP) { $ups.ManagementIP } else { '(Not specified)' })),
                    @("Name:", $(if ($ups.Name) { $ups.Name } else { '(Not specified)' }))
                )
                
                for ($j = 0; $j -lt $upsDetails.Count; $j++) {
                    $row = New-Object System.Windows.Controls.RowDefinition
                    $row.Height = "Auto"
                    $upsDetailGrid.RowDefinitions.Add($row)
                    
                    $label = New-Object System.Windows.Controls.Label
                    $label.Content = $upsDetails[$j][0]
                    $label.FontWeight = "Bold"
                    [System.Windows.Controls.Grid]::SetRow($label, $j)
                    [System.Windows.Controls.Grid]::SetColumn($label, 0)
                    $upsDetailGrid.Children.Add($label)
                    
                    $clickableText = New-ClickableText -Value $upsDetails[$j][1]
                    $clickableText.Margin = "10,5,0,5"
                    [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
                    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                    $upsDetailGrid.Children.Add($clickableText)
                }
                
                # Add to main grid
                $headerRow = New-Object System.Windows.Controls.RowDefinition
                $headerRow.Height = "Auto"
                $detailRow = New-Object System.Windows.Controls.RowDefinition
                $detailRow.Height = "Auto"
                $upsGrid.RowDefinitions.Add($headerRow)
                $upsGrid.RowDefinitions.Add($detailRow)
                
                [System.Windows.Controls.Grid]::SetRow($upsHeader, $upsGrid.RowDefinitions.Count - 2)
                [System.Windows.Controls.Grid]::SetRow($upsDetailGrid, $upsGrid.RowDefinitions.Count - 1)
                $upsGrid.Children.Add($upsHeader)
                $upsGrid.Children.Add($upsDetailGrid)
            }
        }
        
        $upsContent.Content = $upsGrid
        $upsTab.Content = $upsContent
        
# ============================================================================
        # 10. PRINTER TAB
        # ============================================================================
        $printerContent = New-Object System.Windows.Controls.ScrollViewer
        $printerContent.VerticalScrollBarVisibility = "Auto"
        $printerGrid = New-Object System.Windows.Controls.Grid
        $printerGrid.Margin = "10"
        
        # Count actual Printers with data
        $actualPrinterCount = 0
        foreach ($printer in $Site.PrinterDevices) {
            if (-not [string]::IsNullOrWhiteSpace($printer.ManagementIP) -or 
                -not [string]::IsNullOrWhiteSpace($printer.Name) -or 
                -not [string]::IsNullOrWhiteSpace($printer.Model) -or
                -not [string]::IsNullOrWhiteSpace($printer.SerialNumber)) {
                $actualPrinterCount++
            }
        }
        
        if ($Site.PrinterDevices.Count -eq 0 -or $actualPrinterCount -eq 0) {
            $noPrinterText = New-Object System.Windows.Controls.TextBlock
            $noPrinterText.Text = "No printers configured for this site."
            $noPrinterText.FontStyle = "Italic"
            $noPrinterText.Foreground = "Gray"
            $noPrinterText.Margin = "10"
            $printerGrid.Children.Add($noPrinterText)
        } else {
            for ($i = 0; $i -lt $Site.PrinterDevices.Count; $i++) {
                $printer = $Site.PrinterDevices[$i]
                
                # Create Printer header
                $printerHeader = New-Object System.Windows.Controls.TextBlock
                $printerHeader.Text = "Printer $($i+1):"
                $printerHeader.FontWeight = "Bold"
                $printerHeader.FontSize = 14
                $printerHeader.Margin = "0,10,0,5"
                
                # Create Printer details grid
                $printerDetailGrid = New-Object System.Windows.Controls.Grid
                $printerDetailGrid.Margin = "20,0,0,10"
                
                $col1 = New-Object System.Windows.Controls.ColumnDefinition
                $col1.Width = "Auto"
                $col2 = New-Object System.Windows.Controls.ColumnDefinition
                $col2.Width = "*"
                $printerDetailGrid.ColumnDefinitions.Add($col1)
                $printerDetailGrid.ColumnDefinitions.Add($col2)
                
                $printerDetails = @(
                    @("Management IP:", $(if ($printer.ManagementIP) { $printer.ManagementIP } else { '(Not specified)' })),
                    @("Name:", $(if ($printer.Name) { $printer.Name } else { '(Not specified)' })),
                    @("Model:", $(if ($printer.Model) { $printer.Model } else { '(Not specified)' })),
                    @("Serial Number:", $(if ($printer.SerialNumber) { $printer.SerialNumber } else { '(Not specified)' }))
                )
                
                for ($j = 0; $j -lt $printerDetails.Count; $j++) {
                    $row = New-Object System.Windows.Controls.RowDefinition
                    $row.Height = "Auto"
                    $printerDetailGrid.RowDefinitions.Add($row)
                    
                    $label = New-Object System.Windows.Controls.Label
                    $label.Content = $printerDetails[$j][0]
                    $label.FontWeight = "Bold"
                    [System.Windows.Controls.Grid]::SetRow($label, $j)
                    [System.Windows.Controls.Grid]::SetColumn($label, 0)
                    $printerDetailGrid.Children.Add($label)
                    
                    $clickableText = New-ClickableText -Value $printerDetails[$j][1]
                    $clickableText.Margin = "10,5,0,5"
                    [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
                    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
                    $printerDetailGrid.Children.Add($clickableText)
                }
                
                # Add to main grid
                $headerRow = New-Object System.Windows.Controls.RowDefinition
                $headerRow.Height = "Auto"
                $detailRow = New-Object System.Windows.Controls.RowDefinition
                $detailRow.Height = "Auto"
                $printerGrid.RowDefinitions.Add($headerRow)
                $printerGrid.RowDefinitions.Add($detailRow)
                
                [System.Windows.Controls.Grid]::SetRow($printerHeader, $printerGrid.RowDefinitions.Count - 2)
                [System.Windows.Controls.Grid]::SetRow($printerDetailGrid, $printerGrid.RowDefinitions.Count - 1)
                $printerGrid.Children.Add($printerHeader)
                $printerGrid.Children.Add($printerDetailGrid)
            }
        }
        
        $printerContent.Content = $printerGrid
        $printerTab.Content = $printerContent
        
        # ============================================================================
        # ADD ALL TABS TO TABCONTROL
        # ============================================================================
        $stkSiteDetails.Items.Add($basicTab)
        $stkSiteDetails.Items.Add($switchTab)
        $stkSiteDetails.Items.Add($apTab)
        $stkSiteDetails.Items.Add($firewallTab)
        $stkSiteDetails.Items.Add($primaryTab)
        $stkSiteDetails.Items.Add($backupTab)
        $stkSiteDetails.Items.Add($vlanTab)
        $stkSiteDetails.Items.Add($cctvTab)
        $stkSiteDetails.Items.Add($upsTab)
        $stkSiteDetails.Items.Add($printerTab)
        
    }
    catch {
        [System.Windows.MessageBox]::Show("Error displaying site details: $_", "Display Error", "OK", "Error")
    }
}

function Update-DataGridWithSearch {
    $searchTerm = $txtSearchSites.Text
    
    # Store current selection
    $selectedItems = @()
    if ($dgSites.SelectedItems.Count -gt 0) {
        foreach ($item in $dgSites.SelectedItems) {
            $selectedItems += $item.ID
        }
    }
    
    # Get all data from the data store
    $allData = $siteDataStore.GetAllEntries()
    
    # Filter data if search term exists
    if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
        $searchTerm = $searchTerm.Trim().ToLower()
        $allData = $allData | Where-Object {
            $_.SiteCode.ToLower().Contains($searchTerm) -or
            $_.SiteName.ToLower().Contains($searchTerm) -or
            $_.SiteSubnetCode.ToLower().Contains($searchTerm) -or
            $_.SiteAddress.ToLower().Contains($searchTerm) -or
            $_.MainContactName.ToLower().Contains($searchTerm) -or
            $_.Switch1IP.ToLower().Contains($searchTerm) -or
            $_.Switch1Name.ToLower().Contains($searchTerm) -or
            $_.FirewallIP.ToLower().Contains($searchTerm) -or
            $_.PrimaryVendor.ToLower().Contains($searchTerm)
        }
    }
    
    # Sort by ID numerically and update DataGrid
    $allData = $allData | Sort-Object -Property @{Expression={[int]$_.ID}; Ascending=$true}
    
    # Only update ItemsSource if data actually changed
    if ($dgSites.Items.Count -ne $allData.Count) {
        $dgSites.ItemsSource = @($allData)
        
        # Restore selection if items still exist
        if ($selectedItems.Count -gt 0) {
            $dgSites.SelectedItems.Clear()
            foreach ($item in $dgSites.Items) {
                if ($item.ID -in $selectedItems) {
                    $dgSites.SelectedItems.Add($item)
                }
            }
        }
    }
    
    # Update status bar
    $txtStatusBarSites.Text = "Total Sites: $($allData.Count)"
    $selectedCount = $dgSites.SelectedItems.Count
    if ($selectedCount -gt 0) {
        if ($selectedCount -eq 1) {
            $txtStatusBarSiteSelected.Text = "Selected: $($dgSites.SelectedItems[0].SiteCode) - $($dgSites.SelectedItems[0].SiteName)"
        } else {
            $txtStatusBarSiteSelected.Text = "Selected: $selectedCount sites"
        }
    } else {
        $txtStatusBarSiteSelected.Text = "Selected: None"
}}

function Lookup-Site {
    param([string]$SearchTerm)
    
    # Hide results at start
    $grpSiteLookupResults.Visibility = "Collapsed"
    
    if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
        Show-CustomDialog "Please enter a site code or name to search" "Input Required" "OK" "Warning"
        return
    }
    
    $searchTerm = $SearchTerm.Trim().ToLower()
    $allSites = $siteDataStore.GetAllEntries()
    
    $foundSite = $allSites | Where-Object {
        $_.SiteCode.ToLower().Contains($searchTerm) -or
        $_.SiteName.ToLower().Contains($searchTerm)
    } | Select-Object -First 1
    
    if ($foundSite) {
        Show-SiteDetails -Site $foundSite
        $grpSiteLookupResults.Visibility = "Visible"
    } else {
        Show-CustomDialog "Site '$SearchTerm' not found in the database." "Not Found" "OK" "Information"
    }
}

function Initialize-ControlReferences {
    param (
        $mainWin
    )
    
    # Basic Info controls
    $script:txtSiteCode = $mainWin.FindName("txtSiteCode")
    $script:txtSiteSubnetCode = $mainWin.FindName("txtSiteSubnetCode")
    $script:txtSiteSubnet = $mainWin.FindName("txtSiteSubnet")
    $script:txtSiteName = $mainWin.FindName("txtSiteNameManage")
    $script:txtSiteAddress = $mainWin.FindName("txtSiteAddress")
    $script:txtMainContactName = $mainWin.FindName("txtMainContactName")
    $script:txtMainContactPhone = $mainWin.FindName("txtMainContactPhone")
    $script:txtSecondContactName = $mainWin.FindName("txtSecondContactName")
    $script:txtSecondContactPhone = $mainWin.FindName("txtSecondContactPhone")

    # Network Equipment
    $script:cmbSwitchCount = $mainWin.FindName("cmbSwitchCount")
    $script:stkSwitches = $mainWin.FindName("stkSwitches")
    $script:txtFirewallIP = $mainWin.FindName("txtFirewallIP")
    $script:txtFirewallName = $mainWin.FindName("txtFirewallName")
    $script:txtFirewallVersion = $mainWin.FindName("txtFirewallVersion")
    $script:txtFirewallSN = $mainWin.FindName("txtFirewallSN")

    # Primary Circuit
    $script:txtPrimaryVendor = $mainWin.FindName("txtPrimaryVendor")
    $script:cmbPrimaryCircuitType = $mainWin.FindName("cmbPrimaryCircuitType")
    $script:txtPrimaryCircuitID = $mainWin.FindName("txtPrimaryCircuitID")
    $script:txtPrimaryDownloadSpeed = $mainWin.FindName("txtPrimaryDownloadSpeed")
    $script:txtPrimaryUploadSpeed = $mainWin.FindName("txtPrimaryUploadSpeed")
    $script:txtPrimaryIPAddress = $mainWin.FindName("txtPrimaryIPAddress")
    $script:txtPrimarySubnetMask = $mainWin.FindName("txtPrimarySubnetMask")
    $script:txtPrimaryDefaultGateway = $mainWin.FindName("txtPrimaryDefaultGateway")
    $script:txtPrimaryDNS1 = $mainWin.FindName("txtPrimaryDNS1")
    $script:txtPrimaryDNS2 = $mainWin.FindName("txtPrimaryDNS2")
    $script:txtPrimaryRouterModel = $mainWin.FindName("txtPrimaryRouterModel")
    $script:txtPrimaryRouterName = $mainWin.FindName("txtPrimaryRouterName")
    $script:txtPrimaryRouterSN = $mainWin.FindName("txtPrimaryRouterSN")
    $script:txtPrimaryPPPoEUsername = $mainWin.FindName("txtPrimaryPPPoEUsername")
    $script:txtPrimaryPPPoEPassword = $mainWin.FindName("txtPrimaryPPPoEPassword")
    $script:chkPrimaryHasModem = $mainWin.FindName("chkPrimaryHasModem")
    $script:stkPrimaryModem = $mainWin.FindName("stkPrimaryModem")
    $script:txtPrimaryModemModel = $mainWin.FindName("txtPrimaryModemModel")
    $script:txtPrimaryModemName = $mainWin.FindName("txtPrimaryModemName")
    $script:txtPrimaryModemSN = $mainWin.FindName("txtPrimaryModemSN")

    # Backup Circuit
    $script:chkHasBackupCircuit = $mainWin.FindName("chkHasBackupCircuit")
    $script:grdBackupCircuit = $mainWin.FindName("grdBackupCircuit")
    $script:txtBackupVendor = $mainWin.FindName("txtBackupVendor")
    $script:cmbBackupCircuitType = $mainWin.FindName("cmbBackupCircuitType")
    $script:txtBackupCircuitID = $mainWin.FindName("txtBackupCircuitID")
    $script:txtBackupDownloadSpeed = $mainWin.FindName("txtBackupDownloadSpeed")
    $script:txtBackupUploadSpeed = $mainWin.FindName("txtBackupUploadSpeed")
    $script:txtBackupIPAddress = $mainWin.FindName("txtBackupIPAddress")
    $script:txtBackupSubnetMask = $mainWin.FindName("txtBackupSubnetMask")
    $script:txtBackupDefaultGateway = $mainWin.FindName("txtBackupDefaultGateway")
    $script:txtBackupDNS1 = $mainWin.FindName("txtBackupDNS1")
    $script:txtBackupDNS2 = $mainWin.FindName("txtBackupDNS2")
    $script:txtBackupRouterModel = $mainWin.FindName("txtBackupRouterModel")
    $script:txtBackupRouterName = $mainWin.FindName("txtBackupRouterName")
    $script:txtBackupRouterSN = $mainWin.FindName("txtBackupRouterSN")
    $script:txtBackupPPPoEUsername = $mainWin.FindName("txtBackupPPPoEUsername")
    $script:txtBackupPPPoEPassword = $mainWin.FindName("txtBackupPPPoEPassword")
    $script:chkBackupHasModem = $mainWin.FindName("chkBackupHasModem")
    $script:stkBackupModem = $mainWin.FindName("stkBackupModem")
    $script:txtBackupModemModel = $mainWin.FindName("txtBackupModemModel")
    $script:txtBackupModemName = $mainWin.FindName("txtBackupModemName")
    $script:txtBackupModemSN = $mainWin.FindName("txtBackupModemSN")

    # VLANs
    $script:txtVlan100 = $mainWin.FindName("txtVlan100")
    $script:txtVlan101 = $mainWin.FindName("txtVlan101")
    $script:txtVlan102 = $mainWin.FindName("txtVlan102")
    $script:txtVlan103 = $mainWin.FindName("txtVlan103")
    $script:txtVlan104 = $mainWin.FindName("txtVlan104")
    $script:txtVlan105 = $mainWin.FindName("txtVlan105")
    $script:txtVlan106 = $mainWin.FindName("txtVlan106")
    $script:txtVlan107 = $mainWin.FindName("txtVlan107")
    $script:txtVlan108 = $mainWin.FindName("txtVlan108")
    $script:txtVlan109 = $mainWin.FindName("txtVlan109")
    $script:txtVlan110 = $mainWin.FindName("txtVlan110")

    # Access Points
    $script:cmbAPCount = $mainWin.FindName("cmbAPCount")
    $script:stkAccessPoints = $mainWin.FindName("stkAccessPoints")

    # UPS
    $script:cmbUPSCount = $mainWin.FindName("cmbUPSCount")
    $script:stkUPS = $mainWin.FindName("stkUPS")

    # CCTV
    $script:cmbCCTVCount = $mainWin.FindName("cmbCCTVCount")
    $script:stkCCTV = $mainWin.FindName("stkCCTV")

    # Printer
    $script:cmbPrinterCount = $mainWin.FindName("cmbPrinterCount")
    $script:stkPrinter = $mainWin.FindName("stkPrinter")

    # Buttons and Controls
    $script:btnAddSite = $mainWin.FindName("btnAddSite")
    $script:btnClearForm = $mainWin.FindName("btnClearForm")
    $script:btnEditSite = $mainWin.FindName("btnEditSite")
    $script:btnDeleteSite = $mainWin.FindName("btnDeleteSite")
    $script:dgSites = $mainWin.FindName("dgSites")
    $script:txtSearchSites = $mainWin.FindName("txtSearchSites")
    $script:btnClearSearchSites = $mainWin.FindName("btnClearSearchSites")
    $script:txtBlkSiteStatus = $mainWin.FindName("txtBlkSiteStatus")

    # Lookup Controls
    $script:txtSiteLookup = $mainWin.FindName("txtSiteLookup")
    $script:btnLookupSite = $mainWin.FindName("btnLookupSite")
    $script:grpSiteLookupResults = $mainWin.FindName("grpSiteLookupResults")
    $script:stkSiteDetails = $mainWin.FindName("stkSiteDetails")

    # Import/Export Controls
    $script:txtSiteCsvFilePath = $mainWin.FindName("txtSiteCsvFilePath")
    $script:btnBrowseSiteCsv = $mainWin.FindName("btnBrowseSiteCsv")
    $script:btnImportSiteCsv = $mainWin.FindName("btnImportSiteCsv")
    $script:btnExportSiteCsv = $mainWin.FindName("btnExportSiteCsv")
    $script:txtBlkSiteImportStatus = $mainWin.FindName("txtBlkSiteImportStatus")
    $script:pnlSiteImportProgress = $mainWin.FindName("pnlSiteImportProgress")
    $script:pbSiteImportProgress = $mainWin.FindName("pbSiteImportProgress")
    $script:txtSiteProgressStatus = $mainWin.FindName("txtSiteProgressStatus")
    $script:txtSiteProgressDetails = $mainWin.FindName("txtSiteProgressDetails")

    # Tab Controls
    $script:MainTabControl = $mainWin.FindName("MainTabControl")
    $script:SiteManagementTabControl = $mainWin.FindName("SiteManagementTabControl")

    # Status Bar
    $script:SiteStatusBar = $mainWin.FindName("SiteStatusBar")
    $script:txtStatusBarSites = $mainWin.FindName("txtStatusBarSites")
    $script:txtStatusBarSiteSelected = $mainWin.FindName("txtStatusBarSiteSelected")

    # IP Network Identifier Controls
    $script:txtIpSubnet = $mainWin.FindName("txtIpSubnet")
    $script:txtVlanId = $mainWin.FindName("txtVlanId")
    $script:txtVlanName = $mainWin.FindName("txtVlanName")
    $script:txtSiteName = $mainWin.FindName("txtSiteName")
    $script:btnAddEntry = $mainWin.FindName("btnAddEntry")
    $script:txtIpLookup = $mainWin.FindName("txtIpLookup")
    $script:btnLookup = $mainWin.FindName("btnLookup")
    $script:txtBlkMatchedSubnet = $mainWin.FindName("txtBlkMatchedSubnet")
    $script:txtBlkVlanId = $mainWin.FindName("txtBlkVlanId")
    $script:txtBlkVlanName = $mainWin.FindName("txtBlkVlanName")
    $script:txtBlkSiteName = $mainWin.FindName("txtBlkSiteName")
    $script:grpLookupResults = $mainWin.FindName("grpLookupResults")
    $script:dgSubnets = $mainWin.FindName("dgSubnets")
    $script:btnDeleteEntry = $mainWin.FindName("btnDeleteEntry")
    $script:txtCsvFilePath = $mainWin.FindName("txtCsvFilePath")
    $script:btnBrowseCsv = $mainWin.FindName("btnBrowseCsv")
    $script:btnImportCsv = $mainWin.FindName("btnImportCsv")
    $script:btnExportCsv = $mainWin.FindName("btnExportCsv")
    $script:txtBlkImportStatus = $mainWin.FindName("txtBlkImportStatus")
    $script:txtBlkSearchedIp = $mainWin.FindName("txtBlkSearchedIp")
    $script:txtSearch = $mainWin.FindName("txtSearch")
    $script:btnClearSearch = $mainWin.FindName("btnClearSearch")
    $script:txtStatusBarSubnets = $mainWin.FindName("txtStatusBarSubnets")
    $script:txtStatusBarSelected = $mainWin.FindName("txtStatusBarSelected")
    $script:pbImportProgress = $mainWin.FindName("pbImportProgress")
    $script:txtProgressStatus = $mainWin.FindName("txtProgressStatus")
    $script:txtProgressDetails = $mainWin.FindName("txtProgressDetails")
    $script:pnlImportProgress = $mainWin.FindName("pnlImportProgress")
    $script:MainStatusBar = $mainWin.FindName("MainStatusBar")
}

function Initialize-EventHandlers {
    param (
        $mainWin
    )

    # Device count handlers
    $cmbSwitchCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'Switch' $cmbSwitchCount })
    $cmbAPCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'AccessPoint' $cmbAPCount })
    $cmbUPSCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'UPS' $cmbUPSCount })
    $cmbCCTVCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'CCTV' $cmbCCTVCount })
    $cmbPrinterCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'Printer' $cmbPrinterCount })

    # Backup circuit checkbox
    $chkHasBackupCircuit.Add_Checked({
        if ($grdBackupCircuit) {
            $grdBackupCircuit.Visibility = "Visible"
        }
    })

    $chkHasBackupCircuit.Add_Unchecked({
        if ($grdBackupCircuit) {
            $grdBackupCircuit.Visibility = "Collapsed"
        }
    })

    # Primary modem checkbox
    $chkPrimaryHasModem.Add_Checked({
        if ($stkPrimaryModem) {
            $stkPrimaryModem.Visibility = "Visible"
        }
    })

    $chkPrimaryHasModem.Add_Unchecked({
        if ($stkPrimaryModem) {
            $stkPrimaryModem.Visibility = "Collapsed"
        }
    })

    # Backup modem checkbox
    $chkBackupHasModem.Add_Checked({
        if ($stkBackupModem) {
            $stkBackupModem.Visibility = "Visible"
        }
    })

    $chkBackupHasModem.Add_Unchecked({
        if ($stkBackupModem) {
            $stkBackupModem.Visibility = "Collapsed"
        }
    })

    # Site Subnet auto-population
    $txtSiteSubnet.Add_TextChanged({
        $vlanControls = @{
            VLAN100 = $txtVlan100
            VLAN101 = $txtVlan101
            VLAN102 = $txtVlan102
            VLAN103 = $txtVlan103
            VLAN104 = $txtVlan104
            VLAN105 = $txtVlan105
            VLAN106 = $txtVlan106
            VLAN107 = $txtVlan107
            VLAN108 = $txtVlan108
            VLAN109 = $txtVlan109
            VLAN110 = $txtVlan110
        }
        Update-VLANsAndIPsFromSubnet -SubnetInput $txtSiteSubnet.Text -VLANControls $vlanControls -DeviceManager $script:DeviceManager -FirewallIPControl $txtFirewallIP -SiteSubnetCodeControl $txtSiteSubnetCode
    })

    # Site Code auto-population
    $txtSiteCode.Add_TextChanged({
        Update-DeviceNamesFromSiteCode -SiteCode $txtSiteCode.Text -DeviceManager $script:DeviceManager -FirewallNameControl $txtFirewallName
    })

    # Circuit type handlers
    $cmbPrimaryCircuitType.Add_SelectionChanged({
        $stkPrimaryGPONElement = $mainWin.FindName("stkPrimaryGPON")
        if ($stkPrimaryGPONElement) {
            if ($cmbPrimaryCircuitType.SelectedItem -and $cmbPrimaryCircuitType.SelectedItem.Content -eq "GPON Fiber") {
                $stkPrimaryGPONElement.Visibility = "Visible"
            } else {
                $stkPrimaryGPONElement.Visibility = "Collapsed"
            }
        }
    })

    $cmbBackupCircuitType.Add_SelectionChanged({
        $stkBackupGPONElement = $mainWin.FindName("stkBackupGPON")
        if ($stkBackupGPONElement) {
            if ($cmbBackupCircuitType.SelectedItem -and $cmbBackupCircuitType.SelectedItem.Content -eq "GPON Fiber") {
                $stkBackupGPONElement.Visibility = "Visible"
            } else {
                $stkBackupGPONElement.Visibility = "Collapsed"
            }
        }
    })
}

function Initialize-ButtonEventHandlers {
    param (
        $mainWin
    )

    # Add Site button
    $btnAddSite.Add_Click({
        $null = Add-Site
    })

    # Clear form button
    $btnClearForm.Add_Click({
        Clear-SiteForm
        $txtBlkSiteStatus.Text = "Form cleared"
        $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Blue
    })

    # Edit site button
    $btnEditSite.Add_Click({
        $selectedItems = @($dgSites.SelectedItems)
        if ($selectedItems.Count -eq 1) {
            # Get the selected site data
            $selectedSite = $selectedItems[0]
            
            # Find the full site entry from the data store
            $allSites = $siteDataStore.GetAllEntries()
            $siteToEdit = $allSites | Where-Object { $_.ID -eq $selectedSite.ID }
            
            if ($siteToEdit) {
                # Show the edit window
                $editResult = Show-EditSiteWindow -SiteToEdit $siteToEdit
                
                if ($editResult) {
                    # Refresh the data grid to show changes
                    Update-DataGridWithSearch
                }
            } else {
                Show-ValidationError "Selected site not found in database." "Site Not Found"
            }
        } elseif ($selectedItems.Count -eq 0) {
            Show-ValidationError "Please select a site to edit." "Selection Required"
        } else {
            Show-ValidationError "Please select only one site to edit at a time." "Multiple Selection"
        }
    })

    # Delete site button
    $btnDeleteSite.Add_Click({
        $selectedItems = @($dgSites.SelectedItems)
        if ($selectedItems.Count -gt 0) {
            $confirm = Show-CustomDialog "Are you sure you want to delete $($selectedItems.Count) selected sites?" "Confirm Deletion" "YesNo" "Warning"
           
            if ($confirm -eq "Yes") {
                $idsToDelete = @()
                foreach ($item in $selectedItems) {
                    $idsToDelete += $item.ID
                }
                if ($siteDataStore.DeleteEntries($idsToDelete)) {
                    Update-DataGridWithSearch
                    Show-ValidationError "Successfully deleted $($selectedItems.Count) sites." "Success"
                } else {
                    Show-ValidationError "Error deleting sites." "Delete Error"
                }
            }
        } else {
            Show-ValidationError "Please select one or more sites to delete." "Selection Required"
        }
    })

    # Lookup site button
    $btnLookupSite.Add_Click({
        $searchTerm = $txtSiteLookup.Text.Trim()
        Lookup-Site -SearchTerm $searchTerm
    })

    # Clear search button
    $btnClearSearchSites.Add_Click({
        $txtSearchSites.Text = ""
        Update-DataGridWithSearch
    })

    # Search functionality
    $txtSearchSites.Add_TextChanged({
        $script:SearchTimer.Stop()
        $script:SearchTimer.Start()
    })

    # Enter key support for lookup
    $txtSiteLookup.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $searchTerm = $txtSiteLookup.Text.Trim()
            Lookup-Site -SearchTerm $searchTerm
        }
    })
}

function Initialize-ImportExportHandlers {
    # Browse for Excel file
    $btnBrowseSiteCsv.Add_Click({
        try {
            $openDialog = New-Object Microsoft.Win32.OpenFileDialog
            $openDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
            $openDialog.DefaultExt = "xlsx"
            
            if ($openDialog.ShowDialog() -eq $true) {
                $txtSiteCsvFilePath.Text = $openDialog.FileName
                $txtBlkSiteImportStatus.Text = "File selected: $($openDialog.FileName)"
                $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Blue
            }
        } catch {
            $txtBlkSiteImportStatus.Text = "Error selecting file: $_"
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
        }
    })

    # Import from Excel
    $btnImportSiteCsv.Add_Click({
        try {
            $filePath = $txtSiteCsvFilePath.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($filePath)) {
                Show-CustomDialog "Please select an Excel file first." "No File Selected" "OK" "Warning"
                return
            }
            
            $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Visible
            $pbSiteImportProgress.Value = 0
            $txtSiteProgressStatus.Text = "Initializing Excel application..."
            $txtSiteProgressDetails.Text = "Please wait while Excel is starting up..."
            [System.Windows.Forms.Application]::DoEvents()
            
            $result = Import-SitesFromExcel -ExcelFilePath $filePath
            
            $txtBlkSiteImportStatus.Text = $result
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
            
            $importLines = $result -split "`n"
            
            $totalLine = $importLines | Where-Object { $_ -like "Total sites processed: *" } | Select-Object -First 1
            $updatedLine = $importLines | Where-Object { $_ -like " Updated existing: * sites" -or $_ -like " Updated: * sites" } | Select-Object -First 1  
            $noChangesLine = $importLines | Where-Object { $_ -like " No changes needed: * sites" -or $_ -like " No changes: * sites" } | Select-Object -First 1
            $newLine = $importLines | Where-Object { $_ -like " Successfully imported: * sites" } | Select-Object -First 1
            $subnetWarningsLine = $importLines | Where-Object { $_ -like " Subnet warnings: * sites" } | Select-Object -First 1
            
            $popupBody = ""
            if ($totalLine) { $popupBody += $totalLine.Trim() }
            if ($newLine) { $popupBody += "`n" + $newLine.Trim() }
            if ($updatedLine) { $popupBody += "`n" + $updatedLine.Trim() }
            if ($noChangesLine) { $popupBody += "`n" + $noChangesLine.Trim() }
            if ($subnetWarningsLine) { $popupBody += "`n" + $subnetWarningsLine.Trim() }
            
            $errorLines = $importLines | Where-Object { $_ -like "❌*" }
            if ($errorLines.Count -gt 0) {
                $popupBody += "`nValidation errors: $($errorLines.Count) sites"
            }
            
            Show-CustomDialog $popupBody "Import completed successfully!" "OK" "Information"
            
            Update-DataGridWithSearch
            
        } catch {
            $errorMsg = "Import failed: $_"
            $txtBlkSiteImportStatus.Text = $errorMsg
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
            Show-CustomDialog $errorMsg "Import Error" "OK" "Error"
        } finally {
            if ($pnlSiteImportProgress) {
                $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
            }
        }
    })

    # Export to Excel
    $btnExportSiteCsv.Add_Click({
        try {
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            $saveDialog.DefaultExt = "xlsx"
            $saveDialog.FileName = "sites_export_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
            
            if ($saveDialog.ShowDialog() -eq $true) {
                $result = Export-SitesToExcel -FilePath $saveDialog.FileName
                
                $fileInfo = Get-Item $saveDialog.FileName
                $fileSizeBytes = $fileInfo.Length
                
                if ($fileSizeBytes -lt 1KB) {
                    $fileSize = "$fileSizeBytes bytes"
                } elseif ($fileSizeBytes -lt 1MB) {
                    $fileSizeKB = [math]::Round($fileSizeBytes / 1KB, 1)
                    $fileSize = "$fileSizeKB KB"
                } else {
                    $fileSizeMB = [math]::Round($fileSizeBytes / 1MB, 1)
                    $fileSize = "$fileSizeMB MB"
                }
                
                $allSites = $siteDataStore.GetAllEntries()
                $totalSites = $allSites.Count
                $siteCodesList = ($allSites.SiteCode | Sort-Object) -join ", "
                
                $mainResult = @"
Excel export completed successfully!
==========================================

Total sites exported: $totalSites

Site Exported :
$siteCodesList

EXPORT DETAILS:
File size: $fileSize
Location: $($saveDialog.FileName)
"@

                $popupResult = @"
Total sites exported: $totalSites
File size: $fileSize
Location: $([System.IO.Path]::GetFileName($saveDialog.FileName))
"@

                $txtBlkSiteImportStatus.Text = $mainResult
                $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                
                Show-CustomDialog $popupResult "Export completed successfully!" "OK" "Information"
                
            } else {
                $txtBlkSiteImportStatus.Text = "Export cancelled"
                $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Orange
            }
            
        } catch {
            $errorMsg = "Export failed: $_"
            $txtBlkSiteImportStatus.Text = $errorMsg
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
            Show-CustomDialog $errorMsg "Export Error" "OK" "Error"
        }
    })
}

function Initialize-DataGridHandlers {
    # DataGrid selection changed
    $dgSites.Add_SelectionChanged({
        $selectedItems = @($dgSites.SelectedItems)
        if ($selectedItems.Count -gt 0) {
            if ($selectedItems.Count -eq 1) {
                $txtStatusBarSiteSelected.Text = "Selected: $($selectedItems[0].SiteCode) - $($selectedItems[0].SiteName)"
            } else {
                $txtStatusBarSiteSelected.Text = "Selected: $($selectedItems.Count) sites"
            }
        } else {
            $txtStatusBarSiteSelected.Text = "Selected: None"
        }
    })

    # Enable Delete key for selected sites
    $dgSites.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
            $selectedItems = @($dgSites.SelectedItems)
            if ($selectedItems.Count -gt 0) {
                $confirm = Show-CustomDialog "Are you sure you want to delete $($selectedItems.Count) selected sites?" "Confirm Deletion" "YesNo" "Warning"
                if ($confirm -eq "Yes") {
                    $idsToDelete = @()
                    foreach ($item in $selectedItems) {
                        $idsToDelete += $item.ID
                    }
                    if ($siteDataStore.DeleteEntries($idsToDelete)) {
                        Update-DataGridWithSearch
                        Show-ValidationError "Successfully deleted $($selectedItems.Count) sites." "Success"
                    } else {
                        Show-ValidationError "Error deleting sites." "Delete Error"
                    }
                }
                $e.Handled = $true
            }
        }
    })

    # Double-click handler
    $dgSites.Add_MouseDoubleClick({
        param($sender, $e)
        try {
            $clickedItem = $dgSites.SelectedItem
            if ($clickedItem) {
                $txtBlkSiteStatus.Text = "Opening edit window for site: $($clickedItem.SiteCode)..."
                $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Blue
                
                $allSites = $siteDataStore.GetAllEntries()
                $siteToEdit = $allSites | Where-Object { $_.ID -eq $clickedItem.ID }
                
                if ($siteToEdit) {
                    $editResult = Show-EditSiteWindow -SiteToEdit $siteToEdit
                    if ($editResult) {
                        Update-DataGridWithSearch
                        $txtBlkSiteStatus.Text = "Site '$($siteToEdit.SiteCode)' updated successfully!"
                        $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Green
                    } else {
                        $txtBlkSiteStatus.Text = "Edit cancelled for site: $($siteToEdit.SiteCode)"
                        $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Orange
                    }
                } else {
                    $txtBlkSiteStatus.Text = "Error: Selected site not found in database"
                    $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Red
                }
            } else {
                $txtBlkSiteStatus.Text = "Double-click on a site row to edit it"
                $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Gray
            }
        } catch {
            $txtBlkSiteStatus.Text = "Error opening edit window: $_"
            $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Red
        }
    })
}

function Initialize-WindowLoadedHandler {
    $mainWin.Add_Loaded({
        try {
            # Initialize managers after window is fully loaded
            $script:DeviceManager = [DevicePanelManager]::new($mainWin)
            $script:FieldManager = [FieldMappingManager]::new($mainWin)
            
            # Set initial visibility states
            if ($grpSiteLookupResults) { $grpSiteLookupResults.Visibility = "Collapsed" }
            if ($grdBackupCircuit) { $grdBackupCircuit.Visibility = "Collapsed" }
            if ($stkPrimaryModem) { $stkPrimaryModem.Visibility = "Collapsed" }
            if ($stkBackupModem) { $stkBackupModem.Visibility = "Collapsed" }
            if ($pnlSiteImportProgress) { $pnlSiteImportProgress.Visibility = "Collapsed" }
            if ($SiteStatusBar) { $SiteStatusBar.Visibility = "Collapsed" }
            
            # Initialize IP Network components
            if ($grpLookupResults) { $grpLookupResults.Visibility = "Collapsed" }
            if ($pnlImportProgress) { $pnlImportProgress.Visibility = "Collapsed" }
            if ($MainStatusBar) { $MainStatusBar.Visibility = "Collapsed" }
            
            # Initialize the data grids
            Update-DataGridWithSearch
            
            if ($dgSubnets -ne $null) {
                Update-SubnetDataGridWithSearch
            }
            
            if ($btnAddEntry -ne $null) {
                Initialize-IPNetworkEventHandlers
            }
            
            # Initialize phone formatting
            $txtMainContactPhone.Add_LostFocus({ $this.Text = Format-PhoneNumber $this.Text })
            $txtSecondContactPhone.Add_LostFocus({ $this.Text = Format-PhoneNumber $this.Text })
            
        } catch {
            [System.Windows.MessageBox]::Show("Error initializing application: $_", "Initialization Error", "OK", "Error")
        }
    })
}

function Initialize-TabHandlers {
    $MainTabControl.Add_SelectionChanged({
        $selectedTab = $MainTabControl.SelectedItem
        
        if ($selectedTab -ne $null) {
            $txtBlkSiteStatus.Text = ""
            $txtBlkSiteImportStatus.Text = ""
            
            if ($selectedTab.Header -eq "Manage Sites") {
                $SiteStatusBar.Visibility = [System.Windows.Visibility]::Visible
                Update-DataGridWithSearch
            } else {
                $SiteStatusBar.Visibility = [System.Windows.Visibility]::Collapsed
            }
            
            $parentTab = $MainTabControl.SelectedItem
            if ($parentTab -and $parentTab.Header -ne "Site Network Identifier") {
                $txtSiteLookup.Text = ""
                $grpSiteLookupResults.Visibility = "Collapsed"
            }
            
            if ($selectedTab.Header -ne "Import/Export") {
                $txtSiteCsvFilePath.Text = ""
            }
        }
    })
}

# Modules\Site\UI\SiteControls.ps1

function Initialize-AllHandlers {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$mainWin
    )
    
    try {
        Write-Host "Initializing all handlers..."

        # Initialize all the controls
        Initialize-ControlReferences -mainWin $mainWin
        
        # Initialize event handlers
        Initialize-EventHandlers -mainWin $mainWin
        Initialize-ButtonEventHandlers -mainWin $mainWin
        Initialize-ImportExportHandlers
        Initialize-DataGridHandlers
        Initialize-WindowLoadedHandler
        Initialize-TabHandlers
        
        # Initialize data stores
        $script:siteDataStore = [SiteDataStore]::new()
        $script:subnetDataStore = [SubnetDataStore]::new()
        
        Write-Host "All handlers initialized successfully"
        return $true
    }
    catch {
        Write-Host "Error initializing handlers: $_"
        throw "Failed to initialize handlers: $_"
    }
}