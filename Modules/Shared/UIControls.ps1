class PhoneNumberConverter : System.Windows.Data.IValueConverter {
    [object] Convert([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return Format-PhoneNumber $value
    }
    
    [object] ConvertBack([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return $value
    }
}

function Format-PhoneNumber {
    param([string]$PhoneNumber)
    
    if ([string]::IsNullOrWhiteSpace($PhoneNumber)) { return "" }
    
    # Remove all non-digits
    $digits = $PhoneNumber -replace '\D', ''
    
    # Format 10 digits: xxxxxxxxxx -> +1 (xxx) xxx-xxxx
    if ($digits.Length -eq 10) {
        return "+1 ($($digits.Substring(0,3))) $($digits.Substring(3,3))-$($digits.Substring(6,4))"
    }
    
    # Return original if not 10 digits
    return $PhoneNumber
}

function Set-ComboBoxValue {
    param(
        [System.Windows.Controls.ComboBox]$ComboBox,
        [object]$Value,  # Can be string, int, or any type
        [switch]$ByContent = $false  # If true, match by Content property, otherwise by value
    )
    
    if ($ComboBox -eq $null) { return }
    
    if ($Value -eq $null -or $Value -eq "") {
        $ComboBox.SelectedIndex = -1
        return
    }
    
    # Convert value to string for comparison
    $searchValue = $Value.ToString().Trim()
    
    for ($i = 0; $i -lt $ComboBox.Items.Count; $i++) {
        $itemValue = ""
        
        if ($ByContent -and $ComboBox.Items[$i].Content) {
            $itemValue = $ComboBox.Items[$i].Content.ToString().Trim()
        } elseif ($ComboBox.Items[$i]) {
            $itemValue = $ComboBox.Items[$i].ToString().Trim()
        }
        
        # Try exact match first, then try numeric comparison for numbers
        if ($itemValue -eq $searchValue) {
            $ComboBox.SelectedIndex = $i
            return
        }
        
        # Try numeric comparison if both values can be converted to numbers
        try {
            $numericSearch = [decimal]$searchValue
            $numericItem = [decimal]$itemValue
            if ($numericSearch -eq $numericItem) {
                $ComboBox.SelectedIndex = $i
                return
            }
        } catch {
            # Not numeric, continue with string comparison
        }
    }
    
    # No match found
    $ComboBox.SelectedIndex = -1
}

function New-ClickableText {
        param(
            [string]$Value
        )
        
        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.VerticalAlignment = 'Center'
        
        # Check for empty, null, or "(Not specified)" values
        if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq "(Not specified)") {
            $textBlock.Text = "(Not specified)"
            $textBlock.Foreground = [System.Windows.Media.Brushes]::Gray
            return $textBlock
        }
        
        # Make it clickable
        $textBlock.Text = $Value
        $textBlock.Cursor = [System.Windows.Input.Cursors]::Hand
        $textBlock.Foreground = [System.Windows.Media.Brushes]::Blue
        $textBlock.TextDecorations = [System.Windows.TextDecorations]::Underline
        $textBlock.ToolTip = "Click to copy: $Value"
        
        # Store original value as a property to avoid closure issues
        $textBlock | Add-Member -MemberType NoteProperty -Name "OriginalText" -Value $Value
        
        # Add click event to copy
        $textBlock.Add_MouseLeftButtonDown({
            param($sender, $e)
            try {
                $valueToCopy = $sender.OriginalText
                [System.Windows.Clipboard]::SetText($valueToCopy)
                
                # Simple feedback - change text briefly
                $sender.Text = "Copied!"
                $sender.Foreground = [System.Windows.Media.Brushes]::Green
                
                # Use DispatcherTimer for proper UI thread handling
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [System.TimeSpan]::FromMilliseconds(800)
                
                # Store reference to the textblock in timer's tag
                $timer.Tag = $sender
                
                $timer.Add_Tick({
                    param($timerSender, $timerArgs)
                    $textBlock = $timerSender.Tag
                    $textBlock.Text = $textBlock.OriginalText
                    $textBlock.Foreground = [System.Windows.Media.Brushes]::Blue
                    $timerSender.Stop()
                })
                $timer.Start()
                
            } catch {
            }
        })
        return $textBlock
}

function Release-ComObject {
    param($ComObject)
    if ($ComObject) {
        try {
            $refCount = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            # Keep releasing until reference count is 0
            while ($refCount -gt 0) {
                $refCount = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            }
        } catch {
            # Silent fail on release
        }
    }
    return $null
}

