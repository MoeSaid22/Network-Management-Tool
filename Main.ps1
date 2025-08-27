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

# Debug functions
$script:DEBUG_MODE = $true

function Write-DebugInfo {
    param([string]$message)
    if ($script:DEBUG_MODE) {
        Write-Host "[DEBUG] $message" -ForegroundColor Magenta
    }
}

function Test-DataStores {
    Write-DebugInfo "Testing data store accessibility..."
    
    if ($null -eq $script:siteDataStore) {
        Write-DebugInfo "siteDataStore is null!"
    } else {
        Write-DebugInfo "siteDataStore exists"
    }
    
    if ($null -eq $script:subnetDataStore) {
        Write-DebugInfo "subnetDataStore is null!"
    } else {
        Write-DebugInfo "subnetDataStore exists"
    }
}

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
        Write-Host "=== Starting Data Store Initialization ==="
        Write-Host "[DEBUG] Beginning data store initialization"
        
        # Check if data files exist
        Write-Host "[DEBUG] Checking data file paths..."
        Write-Host "[DEBUG] Site data path: $siteDataPath"
        Write-Host "[DEBUG] IP data path: $ipDataPath"
        
        if (-not (Test-Path $siteDataPath)) {
            Write-Host "[DEBUG] Site data file does not exist, will be created: $siteDataPath"
            # Create empty JSON file
            '[]' | Out-File -FilePath $siteDataPath -Encoding UTF8
        } else {
            Write-Host "[DEBUG] Site data file exists: $siteDataPath"
        }
        
        if (-not (Test-Path $ipDataPath)) {
            Write-Host "[DEBUG] IP data file does not exist, will be created: $ipDataPath"
            # Create empty JSON file
            '[]' | Out-File -FilePath $ipDataPath -Encoding UTF8
        } else {
            Write-Host "[DEBUG] IP data file exists: $ipDataPath"
        }
        
        Write-Host "Initializing data stores..."
        # Initialize site data store
        Write-Host "Creating new site data store..." -ForegroundColor Yellow
        Write-DebugInfo "Creating site data store"
        $script:siteDataStore = [SiteDataStore]::new()
        Write-Host "Setting site data path: $siteDataPath" -ForegroundColor Yellow
        Write-DebugInfo "Setting site data path"
        $script:siteDataStore.SetDataPath($siteDataPath)

        # Initialize subnet data store
        Write-Host "Creating new subnet data store..." -ForegroundColor Yellow
        Write-DebugInfo "Creating subnet data store"
        $script:subnetDataStore = [SubnetDataStore]::new()
        Write-Host "Setting IP data path: $ipDataPath" -ForegroundColor Yellow
        Write-DebugInfo "Setting IP data path"
        $script:subnetDataStore.SetDataPath($ipDataPath)

        Write-Host "=== Data Store Initialization Completed Successfully ===" -ForegroundColor Green
        Write-DebugInfo "Data store initialization complete"
        
        # Verify data stores
        Test-DataStores
    }
    catch {
        Write-Host "=== Data Store Initialization Failed ===" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host "Stack Trace:" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
        Write-DebugInfo "Data store initialization failed with error: $_"
        throw
    }
}

# Add this function after your other debug functions
function Test-ControlReferences {
    param(
        [System.Windows.Window]$window
    )
    Write-DebugInfo "Testing control references..."
    Write-DebugInfo "Window type: $($window.GetType().FullName)"
    
    try {
        # Check if Content is loaded
        if ($null -eq $window.Content) {
            Write-DebugInfo "WARNING: Window.Content is null!"
        } else {
            Write-DebugInfo "Window.Content is of type: $($window.Content.GetType().FullName)"
        }

        # Try to find all named elements
        $allElements = Get-NamedElements -window $window
        if ($allElements.Count -eq 0) {
            Write-DebugInfo "WARNING: No named elements found in window!"
        } else {
            foreach ($element in $allElements) {
                Write-DebugInfo "Found element: Name='$($element.Name)' Type='$($element.GetType().Name)'"
            }
        }
    }
    catch {
        Write-DebugInfo "Error testing controls: $_"
        Write-DebugInfo "Stack trace: $($_.ScriptStackTrace)"
    }
}

function Get-NamedElements {
    param(
        [System.Windows.DependencyObject]$window
    )
    
    $result = @()
    
    try {
        if ($null -ne $window) {
            # Get the count of children
            $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($window)
            Write-DebugInfo "Child count: $count"
            
            for ($i = 0; $i -lt $count; $i++) {
                $child = [System.Windows.Media.VisualTreeHelper]::GetChild($window, $i)
                if ($null -ne $child.Name) {
                    $result += $child
                }
                # Recursively check children
                $result += Get-NamedElements -window $child
            }
        }
    }
    catch {
        Write-DebugInfo "Error in Get-NamedElements: $_"
    }
    
    return $result
}

# Main initialization sequence
try {
    Write-Host "=== Starting Application Initialization ===" -ForegroundColor Cyan
    Write-DebugInfo "Starting main initialization sequence"
    
    # Define paths
    $xamlPath = Join-Path $scriptRoot "UI\NetworkManagement.xaml"
    $siteDataPath = Join-Path $scriptRoot "Data\site_data.json"
    $ipDataPath = Join-Path $scriptRoot "Data\ip_data.json"

    Write-DebugInfo "Paths configured:"
    Write-DebugInfo "XAML: $xamlPath"
    Write-DebugInfo "Site Data: $siteDataPath"
    Write-DebugInfo "IP Data: $ipDataPath"

    # Initialize data stores FIRST
    Write-Host "`nInitializing data stores..." -ForegroundColor Yellow
    Initialize-DataStores -siteDataPath $siteDataPath -ipDataPath $ipDataPath
    
    # Initialize main window
    Write-Host "`nInitializing main window..." -ForegroundColor Yellow
    Write-DebugInfo "About to initialize main window"
    $mainWindow = Initialize-MainWindow -xamlPath $xamlPath
    Write-DebugInfo "Main window initialized"
    
    if (-not $mainWindow) {
        throw "Failed to initialize main window"
    }

    # Add this line to test controls
    Test-ControlReferences -window $mainWindow

    # Test data stores before handler initialization
    Test-DataStores
    
    # Initialize handlers
    Write-Host "`nInitializing handlers..." -ForegroundColor Yellow
    Write-DebugInfo "About to initialize handlers"
    $result = Initialize-AllHandlers -mainWin $mainWindow
    Write-DebugInfo "Handler initialization completed with result: $result"
    
    if (-not $result) {
        throw "Handler initialization failed"
    }

    # Final data store test
    Test-DataStores
    
    Write-Host "`n=== Application Initialization Completed Successfully ===" -ForegroundColor Green

    # Add debug info to window loaded event
    $mainWindow.Add_Loaded({
        Write-DebugInfo "Window Loaded event triggered"
        Test-DataStores
    })

    Write-DebugInfo "About to show main window"
    # Show the window
    try {
        Write-Host "About to show main window dialog..."
        $result = $mainWindow.ShowDialog()
        Write-Host "Dialog result: $result"
        Write-Host "Dialog closed normally"
    } catch {
        Write-Host "Error showing dialog: $_"
        Write-Host "Exception details: $($_.Exception.GetType().Name) - $($_.Exception.Message)"
    }
}
catch {
    Write-Host "`n=== Application Initialization Failed ===" -ForegroundColor Red
    Write-Host "Error details:" -ForegroundColor Red
    Write-Host "Message: $_" -ForegroundColor Red
    Write-Host "Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    # Debug state on error
    Write-DebugInfo "Error occurred - checking final data store state"
    Test-DataStores
    
    [System.Windows.MessageBox]::Show(
        "Error initializing application: $($_.Exception.Message)",
        "Fatal Error",
        "OK",
        "Error"
    )
    exit 1
}
finally {
    Write-Host "Application closing - $([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss')) UTC"
}