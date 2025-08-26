class SubnetEntry {
    [int]$ID
    [string]$IP_Subnet
    [int]$VLAN_ID
    [string]$VLAN_Name
    [string]$Site_Name

    SubnetEntry([string]$subnet, [int]$vlanId, [string]$vlanName, [string]$siteName) {
        if ([string]::IsNullOrWhiteSpace($subnet) -or $subnet -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}$') {
            throw "Invalid subnet format: $subnet"
        }
        $this.IP_Subnet = $subnet
        $this.VLAN_ID = $vlanId
        $this.VLAN_Name = if($vlanName) { $vlanName } else { "" }
        $this.Site_Name = if($siteName) { $siteName } else { "" }
    }
}

class SubnetDataStore {
    hidden [string]$DataFile = "$PSScriptRoot\ip_data.json"
    hidden [System.Collections.Generic.List[SubnetEntry]]$Entries

    SubnetDataStore() {
        $this.Entries = [System.Collections.Generic.List[SubnetEntry]]::new()
        $this.LoadData()
    }

    [void] LoadData() {
        try {
            if (Test-Path $this.DataFile) {
                $jsonData = Get-Content $this.DataFile -Raw | ConvertFrom-Json
                $this.Entries.Clear()
                
                if ($jsonData) {
                    foreach ($item in $jsonData) {
                        if ($item -and $item.IP_Subnet) {
                            $entry = [SubnetEntry]::new(
                                $item.IP_Subnet,
                                $item.VLAN_ID,
                                $item.VLAN_Name,
                                $item.Site_Name
                            )
                            $entry.ID = $item.ID
                            $this.Entries.Add($entry)
                        }
                    }
                }
            }
        } catch {
            $this.Entries.Clear()
        }
    }

    [void] SaveData() {
        try {
            if ($this.Entries.Count -eq 0) {
                if (Test-Path $this.DataFile) {
                    Remove-Item $this.DataFile -Force
                }
            } else {
                $this.Entries | ConvertTo-Json -Depth 3 | Set-Content $this.DataFile -Encoding UTF8
            }
        } catch {
        }
    }

    [SubnetEntry[]] GetAllEntries() {
        if ($this.Entries -eq $null -or $this.Entries.Count -eq 0) {
            return @()
        }
        return $this.Entries.ToArray()
    }

    [bool] AddEntry([SubnetEntry]$entry) {
        try {
            if ($entry -eq $null) { return $false }
            
            # Check for duplicates
            foreach ($existing in $this.Entries) {
                if ($existing.IP_Subnet -eq $entry.IP_Subnet) {
                    return $false
                }
            }
            
            $entry.ID = $this.GetNextAvailableId()
            $this.Entries.Add($entry)
            $this.SaveData()
            return $true
        } catch {
            return $false
        }
    }

    [bool] DeleteEntries([int[]]$ids) {
        try {
            if ($ids -eq $null -or $ids.Count -eq 0) { return $false }
            
            $countBefore = $this.Entries.Count
            $newList = [System.Collections.Generic.List[SubnetEntry]]::new()
            
            foreach ($entry in $this.Entries) {
                if ($entry.ID -notin $ids) {
                    $newList.Add($entry)
                }
            }
            
            $this.Entries = $newList
            $this.SaveData()
            return $this.Entries.Count -lt $countBefore
        } catch {
            return $false
        }
    }

    hidden [int] GetNextAvailableId() {
        if ($this.Entries.Count -eq 0) { return 1 }
        $maxId = 0
        foreach ($entry in $this.Entries) {
            if ($entry.ID -gt $maxId) { $maxId = $entry.ID }
        }
        return $maxId + 1
    }
}
