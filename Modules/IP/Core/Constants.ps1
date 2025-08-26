# Constants.ps1
$COLUMN_IP_SUBNET = "IP Subnet"
$COLUMN_VLAN_ID = "VLAN ID" 
$COLUMN_VLAN_NAME = "VLAN Name"
$COLUMN_SITE_NAME = "Site Name"

$script:IPSearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:IPSearchTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$script:IPSearchTimer.Add_Tick({
    Update-SubnetDataGridWithSearch
    $script:IPSearchTimer.Stop()
})