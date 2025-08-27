# DialogManager.ps1 - Centralized dialog management
# Handles custom dialogs and message boxes for the Network Management Tool

# Script-scoped variable to store main window reference
$script:mainWin = $null

# Function to initialize the main window reference for dialogs
function Set-DialogMainWindow {
    param(
        [System.Windows.Window]$MainWindow
    )
    
    if ($null -eq $MainWindow) {
        Write-Warning "Attempting to set null main window for dialogs"
        return
    }
    
    $script:mainWin = $MainWindow
    Write-Host "Dialog main window reference set successfully"
}

# Function to get the main window reference for other modules
function Get-DialogMainWindow {
    return $script:mainWin
}

function Show-CustomDialog {
    param(
        [string]$Message,
        [string]$Title,
        [string]$ButtonType = "OK",  # OK, YesNo, YesNoCancel
        [string]$Icon = "Information"  # Information, Warning, Error, Question
    )
    
    try {
        # Create a new window
        $dialog = New-Object System.Windows.Window
        if ($null -eq $dialog) {
            Write-Error "Failed to create dialog window"
            return $null
        }
        
        $dialog.Title = $Title
        $dialog.Width = 400
        $dialog.Height = 200
        $dialog.WindowStartupLocation = "CenterScreen"  # Default to center screen
        $dialog.ResizeMode = "NoResize"
        $dialog.WindowStyle = "SingleBorderWindow"
        
        # Set owner if available, otherwise center on screen
        if ($null -ne $script:mainWin -and $script:mainWin -is [System.Windows.Window]) {
            try {
                $dialog.Owner = $script:mainWin
                $dialog.WindowStartupLocation = "CenterOwner"
            } catch {
                # If setting owner fails, keep center screen
                Write-Warning "Could not set dialog owner, centering on screen: $_"
            }
        }
    
    # Create the content
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "20"
    
    # Add row definitions
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = "*"
    $row2 = New-Object System.Windows.Controls.RowDefinition  
    $row2.Height = "Auto"
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    
    # Message text
    $textBlock = New-Object System.Windows.Controls.TextBlock
    $textBlock.Text = $Message
    $textBlock.TextWrapping = "Wrap"
    $textBlock.VerticalAlignment = "Center"
    $textBlock.HorizontalAlignment = "Center"
    $textBlock.FontSize = 12
    [System.Windows.Controls.Grid]::SetRow($textBlock, 0)
    $grid.Children.Add($textBlock)
    
    # Button panel
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Center"
    $buttonPanel.Margin = "0,20,0,0"
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
    
    $result = $null
    
    if ($ButtonType -eq "OK") {
        $okButton = New-Object System.Windows.Controls.Button
        $okButton.Content = "OK"
        $okButton.Width = 75
        $okButton.Height = 25
        $okButton.IsDefault = $true
        $okButton.Add_Click({
            $script:result = "OK"
            $dialog.DialogResult = $true
            $dialog.Close()
        })
        $buttonPanel.Children.Add($okButton)
    }
    elseif ($ButtonType -eq "YesNo") {
        $yesButton = New-Object System.Windows.Controls.Button
        $yesButton.Content = "Yes"
        $yesButton.Width = 75
        $yesButton.Height = 25
        $yesButton.Margin = "0,0,10,0"
        $yesButton.IsDefault = $true
        $yesButton.Add_Click({
            $script:result = "Yes"
            $dialog.DialogResult = $true
            $dialog.Close()
        })
        
        $noButton = New-Object System.Windows.Controls.Button
        $noButton.Content = "No"
        $noButton.Width = 75
        $noButton.Height = 25
        $noButton.IsCancel = $true
        $noButton.Add_Click({
            $script:result = "No"
            $dialog.DialogResult = $false
            $dialog.Close()
        })
        
        $buttonPanel.Children.Add($yesButton)
        $buttonPanel.Children.Add($noButton)
    }
    
    $grid.Children.Add($buttonPanel)
    $dialog.Content = $grid
    
    # Show dialog and return result with proper error handling
    try {
        if ($null -eq $dialog) {
            Write-Error "Dialog object is null before ShowDialog"
            return $null
        }
        
        $null = $dialog.ShowDialog()
        return $script:result
    } catch {
        Write-Error "Error showing dialog: $_"
        # Try to show a simple message box as fallback
        try {
            [System.Windows.MessageBox]::Show($Message, $Title, "OK", $Icon)
        } catch {
            Write-Error "Could not show fallback message box: $_"
        }
        return $null
    }
} catch {
    Write-Error "Error creating or configuring dialog: $_"
    # Fallback to simple message box
    try {
        [System.Windows.MessageBox]::Show($Message, $Title, "OK", $Icon)
    } catch {
        Write-Error "Could not show fallback message box: $_"
    }
    return $null
}
}

function Show-ValidationError {
        param(
            [string]$Message,
            [string]$Title = "Validation Error"
        )
        
        # Update status text safely
        try {
            if ($txtBlkSiteStatus) {
                $statusType = switch ($Title) {
                    "Success" { "Success" }
                    "Warning" { "Warning" }
                    default { "Error" }
                }
                [StatusManager]::SetStatus($txtBlkSiteStatus, $Message, $statusType)
            }
        } catch {
            # If status text update fails, just continue with dialog
        }
        
        # Show dialog
        Show-CustomDialog $Message $Title "OK" "Information"
    }

