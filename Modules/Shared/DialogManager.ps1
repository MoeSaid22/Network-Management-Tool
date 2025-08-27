function Show-CustomDialog {
    param(
        [string]$Message,
        [string]$Title,
        [string]$ButtonType = "OK",  # OK, YesNo, YesNoCancel
        [string]$Icon = "Information"  # Information, Warning, Error, Question
    )
    
    try {
        # Validate WPF assemblies are loaded
        try {
            if (-not ([System.Management.Automation.PSTypeName]'System.Windows.Window').Type) {
                Write-Warning "WPF assemblies not loaded properly"
                return $null
            }
        }
        catch {
            Write-Warning "Failed to validate WPF assemblies: $_"
            return $null
        }
        
        # Create a new window with enhanced error handling
        $dialog = $null
        try {
            $dialog = New-Object System.Windows.Window
        }
        catch {
            Write-Warning "Failed to create dialog window object: $_"
            return $null
        }
        
        if ($null -eq $dialog) {
            Write-Warning "Failed to create dialog window - object is null"
            return $null
        }

        $dialog.Title = $Title
        $dialog.Width = 400
        $dialog.Height = 200
        $dialog.WindowStartupLocation = "CenterScreen" # Changed from CenterOwner since we may not have an owner
        
        # Only set owner if $mainWin exists and is visible
        if ($null -ne $mainWin -and $mainWin.IsVisible) {
            $dialog.Owner = $mainWin
        }
        
        $dialog.ResizeMode = "NoResize"
        $dialog.WindowStyle = "SingleBorderWindow"
        
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
        
        $script:result = $null
        
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
        
        # Show dialog with comprehensive error handling
        if ($null -ne $dialog) {
            try {
                # Additional validation before showing dialog
                if ($null -eq $dialog.Content) {
                    Write-Warning "Dialog content is null - cannot show dialog"
                    return $null
                }
                
                # Validate dialog is properly initialized
                if ([string]::IsNullOrEmpty($dialog.Title)) {
                    Write-Warning "Dialog title is not set - dialog may not be properly initialized"
                    return $null
                }
                
                Write-Host "[DEBUG] Showing dialog: '$($dialog.Title)'" -ForegroundColor Yellow
                $null = $dialog.ShowDialog()
                Write-Host "[DEBUG] Dialog closed successfully" -ForegroundColor Yellow
            }
            catch {
                Write-Warning "Error showing dialog '$Title': $($_.Exception.Message)"
                Write-Host "[DEBUG] Exception details: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -ForegroundColor Red
                
                # Try to close dialog if it was partially shown
                try {
                    if ($dialog -and $dialog.IsVisible) {
                        $dialog.Close()
                    }
                }
                catch {
                    # Silent fail on cleanup
                }
                return $null
            }
        }
        else {
            Write-Warning "Cannot show dialog - dialog object is null"
            return $null
        }
        
        return $script:result
    }
    catch {
        Write-Warning "Error in Show-CustomDialog: $($_.Exception.Message)"
        Write-Host "[DEBUG] Full exception details: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "[DEBUG] Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
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

