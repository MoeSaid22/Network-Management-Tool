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
