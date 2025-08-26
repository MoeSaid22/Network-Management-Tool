function Clear-SiteForm {
    # Clear all mapped fields using centralized field manager
    if ($script:FieldManager) {
        $script:FieldManager.ClearAllMappings()
    }
    
    # Reset devices using centralized manager
    if ($cmbSwitchCount) { $cmbSwitchCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('Switch', 0) }

    if ($cmbAPCount) { $cmbAPCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('AccessPoint', 0) }

    if ($cmbUPSCount) { $cmbUPSCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('UPS', 0) }

    if ($cmbCCTVCount) { $cmbCCTVCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('CCTV', 0) }
    
    # Reset main checkboxes
    if ($chkHasBackupCircuit) { $chkHasBackupCircuit.IsChecked = $false }
    if ($chkPrimaryHasModem) { $chkPrimaryHasModem.IsChecked = $false }
    if ($chkBackupHasModem) { $chkBackupHasModem.IsChecked = $false }
    
    # Hide conditional sections
    if ($grdBackupCircuit) { $grdBackupCircuit.Visibility = "Collapsed" }
    if ($stkPrimaryModem) { $stkPrimaryModem.Visibility = "Collapsed" }
    if ($stkBackupModem) { $stkBackupModem.Visibility = "Collapsed" }
}

function Get-SiteDataFromForm {
    $site = [SiteEntry]::new()
    
    # Get all mapped fields using centralized field manager
    $script:FieldManager.GetAllMappings($site)
    
    # Get device data using centralized device manager
    $site.SwitchCount = if ($cmbSwitchCount.SelectedItem) { [int]$cmbSwitchCount.SelectedItem.Content } else { 1 }
    $site.Switches = $script:DeviceManager.GetDeviceDataFromUI('Switch')

    $site.APCount = if ($cmbAPCount.SelectedItem) { [int]$cmbAPCount.SelectedItem.Content } else { 1 }
    $site.AccessPoints = $script:DeviceManager.GetDeviceDataFromUI('AccessPoint')

    $site.UPSCount = if ($cmbUPSCount.SelectedItem) { [int]$cmbUPSCount.SelectedItem.Content } else { 0 }
    $site.UPSDevices = $script:DeviceManager.GetDeviceDataFromUI('UPS')

    $site.CCTVCount = if ($cmbCCTVCount.SelectedItem) { [int]$cmbCCTVCount.SelectedItem.Content } else { 0 }
    $site.CCTVDevices = $script:DeviceManager.GetDeviceDataFromUI('CCTV')

    $site.PrinterCount = if ($cmbPrinterCount.SelectedItem) { [int]$cmbPrinterCount.SelectedItem.Content } else { 0 }
    $site.PrinterDevices = $script:DeviceManager.GetDeviceDataFromUI('Printer')
        
    # Get main checkboxes
    $site.HasBackupCircuit = $chkHasBackupCircuit.IsChecked
    
    return $site
}

function Add-Site {
   try {
       # Get site data from form
       $site = Get-SiteDataFromForm

        # Use centralized validation
        try {
            Validate-SiteBasicInfo -Site $site -StatusControl $txtBlkSiteStatus
        } catch {
            Show-CustomDialog $_.Exception.Message "Validation Error" "OK" "Error"
            return $false
        }
       # Try to add the site
       try {
           $addResult = $siteDataStore.AddEntry($site)
           
           if ($addResult -eq $true) {
               Show-CustomDialog "Site '$($site.SiteCode)' added successfully!" "Success" "OK" "Information"
               Clear-SiteForm
               Update-DataGridWithSearch
               return $true
           }
       } catch {
           Show-CustomDialog "Error adding site: $($_.Exception.Message)" "Error" "OK" "Error"
           return $false
       }
   } catch {
       Show-CustomDialog "Error in Add-Site: $($_.Exception.Message)" "Error" "OK" "Error"
       return $false
   }
}