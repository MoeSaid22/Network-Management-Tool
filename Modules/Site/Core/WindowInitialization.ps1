function Initialize-MainWindow {
    param (
        [string]$xamlPath
    )
    
    # Load Assemblies
    Add-Type -AssemblyName WindowsBase, PresentationFramework, PresentationCore, System.Windows.Forms

    # Validate XAML file exists and is readable
    if (-not (Test-Path $xamlPath)) {
        throw "XAML file not found: $xamlPath"
    }

    try {
        $xaml = Get-Content $xamlPath -Raw
        $xml = [xml]$xaml
        
        # Create the window first without loading XAML content
        $reader = New-Object System.Xml.XmlNodeReader $xml
        
        # Add the phone formatter resource BEFORE loading XAML
        $phoneConverter = [PhoneNumberConverter]::new()
        
        # Create a resource dictionary and add our converter
        $resourceDict = New-Object System.Windows.ResourceDictionary
        $resourceDict.Add("PhoneFormatter", $phoneConverter)
        
        # Load the window with resources
        $mainWin = [Windows.Markup.XamlReader]::Load($reader)
        $mainWin.Resources = $resourceDict

        return $mainWin
    }
    catch {
        throw "Error creating window: $_"
    }
}