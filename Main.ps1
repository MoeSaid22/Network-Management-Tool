# Main.ps1
#
# Author: MoeSaid22
# Created: 2025-08-26 20:59:14 UTC

# Set ErrorActionPreference to ensure we catch errors
$ErrorActionPreference = 'Stop'

# Import required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms

# Get the script's directory
$scriptRoot = $PSScriptRoot

Write-Host "Current Date and Time (UTC): $([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Current User's Login: MoeSaid22"

# Define module paths based on actual structure and dependencies
$modulePaths = @(
    # Models/Entry classes must be loaded first
    "Modules\IP\Models\SubnetEntry.ps1"
    "Modules\Site\Models\DeviceInfo.ps1"
    "Modules\Site\Models\SiteEntry.ps1"

    # Shared Modules
    "Modules\Shared\ControlManager.ps1"
    "Modules\Shared\DialogManager.ps1"
    "Modules\Shared\UIControls.ps1"

    # Core Utilities and Validation (needed by other modules)
    "Modules\Site\Core\ValidationUtils.ps1"
    
    # IP Core/Data (after models)
    "Modules\IP\Core\Constants.ps1"
    "Modules\IP\Core\SubnetManager.ps1"
    "Modules\IP\Data\ImportExport.ps1"
    "Modules\IP\UI\IPNetworkControls.ps1"

    # Site Core (after ValidationUtils)
    "Modules\Site\Core\FieldManager.ps1"
    "Modules\Site\Core\DeviceManager.ps1"
    "Modules\Site\Core\SubnetManager.ps1"
    "Modules\Site\Core\WindowInitialization.ps1"

    # Site Data and UI (load these last)
    "Modules\Site\Data\DataStore.ps1"
    "Modules\Site\Data\ImportExport.ps1"
    "Modules\Site\UI\EditWindow.ps1"
    "Modules\Site\UI\FormManager.ps1"
    "Modules\Site\UI\SiteControls.ps1"  # This contains Initialize-AllHandlers
)

# Import all modules
foreach ($modulePath in $modulePaths) {
    $fullPath = Join-Path $scriptRoot $modulePath
    if (Test-Path $fullPath) {
        try {
            . $fullPath
            Write-Host "Successfully loaded: $modulePath"
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Failed to load module: $modulePath`nError: $_",
                "Module Load Error",
                "OK",
                "Error"
            )
            exit 1
        }
    }
    else {
        [System.Windows.MessageBox]::Show(
            "Module not found: $modulePath",
            "Missing Module",
            "OK",
            "Error"
        )
        exit 1
    }
}

# Function to initialize data stores
function Initialize-DataStores {
    param (
        [string]$siteDataPath,
        [string]$ipDataPath
    )
    
    try {
        Write-Host "Initializing data stores..."
        
        # Initialize site data store
        if ($script:siteDataStore) {
            Write-Host "Setting site data path: $siteDataPath"
            $script:siteDataStore.SetDataPath($siteDataPath)
        }
        else {
            Write-Host "Creating new site data store"
            $script:siteDataStore = [SiteDataStore]::new()
            $script:siteDataStore.SetDataPath($siteDataPath)
        }

        # Initialize subnet data store
        if ($script:subnetDataStore) {
            Write-Host "Setting IP data path: $ipDataPath"
            $script:subnetDataStore.SetDataPath($ipDataPath)
        }
        else {
            Write-Host "Creating new subnet data store"
            $script:subnetDataStore = [SubnetDataStore]::new()
            $script:subnetDataStore.SetDataPath($ipDataPath)
        }

        Write-Host "Data stores initialized successfully"
    }
    catch {
        throw "Failed to initialize data stores: $_"
    }
}

# Main application block
try {
    # Define paths
    $xamlPath = Join-Path $scriptRoot "UI\NetworkManagement.xaml"
    $siteDataPath = Join-Path $scriptRoot "Data\site_data.json"
    $ipDataPath = Join-Path $scriptRoot "Data\ip_data.json"

    Write-Host "Starting application initialization..."

    # Initialize main window
    $mainWindow = Initialize-MainWindow -xamlPath $xamlPath
    
    if (-not $mainWindow) {
        throw "Failed to initialize main window"
    }

    # Initialize all handlers and controls
    Initialize-AllHandlers -mainWin $mainWindow

    # Initialize data stores with correct paths
    Initialize-DataStores -siteDataPath $siteDataPath -ipDataPath $ipDataPath

    Write-Host "Application initialization complete"

    # Show the window
    $mainWindow.ShowDialog()
}
catch {
    Write-Host "Fatal Error: $_"
    [System.Windows.MessageBox]::Show(
        "A fatal error has occurred:`n`n$($_.Exception.Message)",
        "Fatal Error",
        "OK",
        "Error"
    )
    exit 1
}
finally {
    Write-Host "Application closing - $([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss')) UTC"
}