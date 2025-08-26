function Import-SubnetsFromCsv {
    param (
        [string]$CsvFilePath
    )

    try {
        # Basic file path validation
        if ([string]::IsNullOrWhiteSpace($CsvFilePath)) {
            throw "File path cannot be empty"
        }
        
        if (-not (Test-Path $CsvFilePath)) {
            throw "CSV file not found"
        }
        
        # Check file extension
        if (-not $CsvFilePath.ToLower().EndsWith('.csv')) {
            throw "File must be a CSV file (.csv extension)"
        }
        
        # Check file size (prevent loading huge files)
        $fileInfo = Get-Item $CsvFilePath
        if ($fileInfo.Length -gt 50MB) {
            throw "File is too large. Maximum size is 50MB"
        }

        # Show progress panel
        if ($pnlImportProgress -ne $null) {
            $pnlImportProgress.Visibility = [System.Windows.Visibility]::Visible
        }
        if ($pbImportProgress -ne $null) { $pbImportProgress.Value = 0 }
        if ($txtProgressStatus -ne $null) { $txtProgressStatus.Text = "Starting import..." }
        if ($txtProgressDetails -ne $null) { $txtProgressDetails.Text = "" }

        # Read CSV data
        $csvData = Import-Csv -Path $CsvFilePath
        $totalLines = $csvData.Count
        $importedCount = 0
        $skippedCount = 0
        $errorMessages = @()

        foreach ($row in $csvData) {
            try {
                # Validate required fields
                if ([string]::IsNullOrWhiteSpace($row.IP_Subnet) -or 
                    [string]::IsNullOrWhiteSpace($row.VLAN_ID) -or 
                    [string]::IsNullOrWhiteSpace($row.VLAN_Name) -or 
                    [string]::IsNullOrWhiteSpace($row.Site_Name)) {
                    throw "Missing required fields"
                }

                # Validate IP format
                if (-not (Test-IPSubnetFormat $row.IP_Subnet)) {
                    throw "Invalid subnet format: $($row.IP_Subnet)"
                }

                # Check for duplicates
                $existingEntries = $subnetDataStore.GetAllEntries()
                if ($existingEntries.IP_Subnet -contains $row.IP_Subnet) {
                    throw "Duplicate subnet: $($row.IP_Subnet)"
                }

                # Sanitize data to prevent CSV injection attacks
                $cleanIPSubnet = Remove-CSVInjection $row.IP_Subnet
                $cleanVlanName = Remove-CSVInjection $row.VLAN_Name
                $cleanSiteName = Remove-CSVInjection $row.Site_Name

                # Create and add entry with sanitized data
                $entry = [SubnetEntry]::new(
                    $cleanIPSubnet,
                    $row.VLAN_ID,
                    $cleanVlanName,
                    $cleanSiteName
                )
                
                if ($subnetDataStore.AddEntry($entry)) {
                    $importedCount++
                }

                # Update progress
                $totalProcessed = $importedCount + $skippedCount
                if ($totalProcessed % 10 -eq 0 -or $totalProcessed -eq $totalLines) {
                    $progress = [math]::Round(($totalProcessed / $totalLines) * 100)
                    if ($pbImportProgress -ne $null) { $pbImportProgress.Value = $progress }
                    if ($txtProgressDetails -ne $null) { $txtProgressDetails.Text = "$importedCount / $totalLines records processed" }
                    [System.Windows.Forms.Application]::DoEvents()
                }

            } catch {
                $skippedCount++
                $errorMessages += "Line $($importedCount + $skippedCount): $_"
            }
        }

        $result = @(
            "Import completed",
            "Successfully imported: $importedCount",
            "Skipped: $skippedCount"
        ) -join "`n"

        if ($errorMessages.Count -gt 0) {
            $result += "`n`nErrors:`n" + ($errorMessages -join "`n")
        }

        return $result
    } catch {
        throw "Import failed: $_"
    } finally {
        if ($pnlImportProgress -ne $null) {
            $pnlImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }
}

function Export-SubnetsToCsv {
    param (
        [string]$FilePath
    )
    
    try {
        # Basic file path validation
        if ([string]::IsNullOrWhiteSpace($FilePath)) {
            throw "File path cannot be empty"
        }
        
        $data = $subnetDataStore.GetAllEntries() | Select-Object IP_Subnet, VLAN_ID, VLAN_Name, Site_Name
        $data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        return "Exported $($data.Count) subnet entries to:`n$FilePath"
    } catch {
        throw "Export failed: $($_.Exception.Message)"
    }
}