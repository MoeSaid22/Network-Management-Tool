# WindowInitialization.ps1
# Handles initialization of the main window and control references

function Initialize-MainWindow {
    param (
        [string]$xamlPath
    )
    
    try {
        Write-DebugInfo "Loading XAML from: $xamlPath"
        
        if (-not (Test-Path $xamlPath)) {
            throw "XAML file not found: $xamlPath"
        }

        # Load XAML content
        $xamlContent = Get-Content $xamlPath -Raw
        Write-DebugInfo "XAML content length: $($xamlContent.Length)"
        
        # Create a memory stream of the XAML content
        $stream = [System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($xamlContent))
        
        # Create the window using XamlReader
        Write-DebugInfo "Creating window from XAML"
        $window = [System.Windows.Markup.XamlReader]::Load($stream)
        $stream.Close()
        
        if ($null -eq $window) {
            throw "Failed to create window from XAML"
        }

        Write-DebugInfo "Window created successfully"
        Write-DebugInfo "Window type: $($window.GetType().FullName)"
        
        # Store controls in a script-level variable
        $script:controlMap = @{}
        
        # Find all named controls
        Write-DebugInfo "Finding named controls..."
        Find-NamedControls -window $window
        
        # Add the control map to the window for access
        $window | Add-Member -MemberType NoteProperty -Name "ControlMap" -Value $script:controlMap
        
        # Add a helper method to get controls
        $window | Add-Member -MemberType ScriptMethod -Name "GetControl" -Value {
            param([string]$name)
            return $this.ControlMap[$name]
        }
        
        return $window
    }
    catch {
        Write-DebugInfo "Error initializing main window: $_"
        Write-DebugInfo "Stack trace: $($_.ScriptStackTrace)"
        throw
    }
}

function Find-NamedControls {
    param(
        [System.Windows.DependencyObject]$window
    )
    
    try {
        Write-DebugInfo "Searching for named controls in window..."
        $script:foundCount = 0
        
        function Find-Controls {
            param(
                [System.Windows.DependencyObject]$parent
            )
            
            if ($null -eq $parent) { return }
            
            # Check if the control has a name
            if ($parent -is [System.Windows.FrameworkElement]) {
                $name = $parent.Name
                if (-not [string]::IsNullOrEmpty($name)) {
                    if (-not $script:controlMap.ContainsKey($name)) {
                        Write-DebugInfo "Found named control: $name (Type: $($parent.GetType().Name))"
                        $script:controlMap[$name] = $parent
                        $script:foundCount++
                    }
                }
            }
            
            # Process logical children
            $logicalChildren = [System.Windows.LogicalTreeHelper]::GetChildren($parent)
            foreach ($child in $logicalChildren) {
                if ($child -is [System.Windows.DependencyObject]) {
                    Find-Controls -parent $child
                }
            }
            
            # Process visual children
            if ($parent -is [System.Windows.Media.Visual]) {
                try {
                    $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent)
                    for ($i = 0; $i -lt $count; $i++) {
                        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                        Find-Controls -parent $child
                    }
                }
                catch {
                    Write-DebugInfo "Skipping visual children for $($parent.GetType().Name): $_"
                }
            }
        }
        
        # Start the search from the window
        Find-Controls -parent $window
        
        Write-DebugInfo "Found $($script:controlMap.Count) unique named controls"
        Write-DebugInfo "Control names found: $($script:controlMap.Keys -join ', ')"
    }
    catch {
        Write-DebugInfo "Error finding named controls: $_"
        Write-DebugInfo "Stack trace: $($_.ScriptStackTrace)"
    }
}

function Initialize-ControlReferences {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$mainWin
    )
    
    try {
        Write-DebugInfo "Initializing control references"
        
        if ($null -eq $mainWin.ControlMap) {
            throw "Window does not have a ControlMap"
        }
        
        # Store references to controls in script variables
        foreach ($controlName in $mainWin.ControlMap.Keys) {
            $control = $mainWin.ControlMap[$controlName]
            Set-Variable -Name $controlName -Value $control -Scope Script
            Write-DebugInfo "Initialized control reference for $controlName"
        }
        
        Write-DebugInfo "Control references initialized successfully"
        return $true
    }
    catch {
        Write-DebugInfo "Error initializing control references: $_"
        Write-DebugInfo "Stack trace: $($_.ScriptStackTrace)"
        throw
    }
}

function Test-ControlReferences {
    param(
        [System.Windows.Window]$window
    )
    Write-DebugInfo "Testing control references..."
    Write-DebugInfo "Window type: $($window.GetType().FullName)"
    
    if ($null -eq $window.Content) {
        Write-DebugInfo "WARNING: Window.Content is null!"
    }
    else {
        Write-DebugInfo "Window.Content is of type: $($window.Content.GetType().FullName)"
        
        # Get child count using VisualTreeHelper
        try {
            $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($window.Content)
            Write-DebugInfo "Child count: $childCount"
        }
        catch {
            Write-DebugInfo "Error getting child count: $_"
        }
    }
    
    if ($null -eq $window.ControlMap) {
        Write-DebugInfo "WARNING: Window.ControlMap is null!"
        return
    }
    
    Write-DebugInfo "Found $($window.ControlMap.Count) controls in ControlMap"
    foreach ($controlName in $window.ControlMap.Keys) {
        $control = $window.ControlMap[$controlName]
        Write-DebugInfo "Control '$controlName' of type '$($control.GetType().Name)' exists"
        if ($null -eq $control) {
            Write-DebugInfo "WARNING: Control '$controlName' is null!"
        }
    }
}

function Initialize-AllHandlers {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$mainWin
    )
    
    try {
        Write-Host "Initializing all handlers..."
        Write-DebugInfo "Starting handler initialization sequence"

        # Initialize all the controls first
        Write-DebugInfo "About to initialize control references"
        Initialize-ControlReferences -mainWin $mainWin
        Write-DebugInfo "Control references initialized"
        
        # Initialize event handlers
        Write-DebugInfo "About to initialize event handlers"
        Initialize-EventHandlers -mainWin $mainWin
        Write-DebugInfo "Event handlers initialized"
        
        # Initialize button handlers
        Write-DebugInfo "About to initialize button event handlers"
        Initialize-ButtonEventHandlers -mainWin $mainWin
        Write-DebugInfo "Button event handlers initialized"
        
        # Initialize import/export handlers
        Write-DebugInfo "About to initialize import/export handlers"
        Initialize-ImportExportHandlers
        Write-DebugInfo "Import/export handlers initialized"
        
        # Initialize DataGrid handlers
        Write-DebugInfo "About to initialize data grid handlers"
        Initialize-DataGridHandlers
        Write-DebugInfo "Data grid handlers initialized"
        
        # Initialize window loaded handler
        Write-DebugInfo "About to initialize window loaded handler"
        Initialize-WindowLoadedHandler
        Write-DebugInfo "Window loaded handler initialized"
        
        # Initialize tab handlers
        Write-DebugInfo "About to initialize tab handlers"
        Initialize-TabHandlers
        Write-DebugInfo "Tab handlers initialized"
        
        Write-Host "All handlers initialized successfully"
        return $true
    }
    catch {
        Write-Host "Error initializing handlers: $_"
        Write-DebugInfo "Handler initialization failed with error: $_"
        Write-DebugInfo "Stack trace: $($_.ScriptStackTrace)"
        throw
    }
}

# Helper function for other modules to access controls
function Get-Control {
    param (
        [Parameter(Mandatory = $true)]
        [string]$name
    )
    
    try {
        if ($null -eq $script:controlMap) {
            throw "Control map is not initialized"
        }
        
        if (-not $script:controlMap.ContainsKey($name)) {
            throw "Control '$name' not found"
        }
        
        return $script:controlMap[$name]
    }
    catch {
        Write-DebugInfo "Error getting control '$name': $_"
        throw
    }
}