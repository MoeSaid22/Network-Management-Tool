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
            Write-Warning "Failed to create dialog window"
            return
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
        
        # Show dialog with error handling
        if ($null -ne $dialog) {
            try {
                $null = $dialog.ShowDialog()
            }
            catch {
                Write-Warning "Error showing dialog: $_"
            }
        }
        
        return $script:result
    }
    catch {
        Write-Warning "Error in Show-CustomDialog: $_"
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

function Show-ImportModeDialog {
    # Create dialog with proper variable handling
    Add-Type -AssemblyName PresentationFramework
    
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="Import Mode Selection" Height="350" Width="500"
    WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="How should duplicate sites be handled?" 
                   FontSize="14" FontWeight="Bold" Margin="0,0,0,20"/>
        
        <StackPanel Grid.Row="1" Margin="0,0,0,20">
            <RadioButton Name="rbUpdate" Content="Smart Update" 
                        GroupName="ImportMode" IsChecked="False" Margin="0,0,0,10"/>
            <TextBlock Text="Only update fields with new data`n   • Preserve existing data where Excel is empty" 
                      Foreground="Gray" Margin="0,0,0,15"/>
            
            <RadioButton Name="rbSkip" Content="Skip Duplicates (Recommended)" GroupName="ImportMode" 
                    IsChecked="True" Margin="0,0,0,10"/>
            <TextBlock Text="Existing sites will not be modified" 
                      Foreground="Gray" Margin="0,0,0,15"/>
            
            <RadioButton Name="rbReplace" Content="Replace Completely" 
                        GroupName="ImportMode" Margin="0,0,0,10"/>
            <TextBlock Text="WARNING: May lose existing data not in Excel" 
                      Foreground="DarkRed" Margin="0,0,0,15"/>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOK" Content="Continue Import" Width="100" Height="30" 
                    Margin="0,0,10,0" IsDefault="True"/>
            <Button Name="btnCancel" Content="Cancel" Width="75" Height="30" 
                    IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    $dialog.Owner = $mainWin
    
    # Get controls
    $rbSkip = $dialog.FindName("rbSkip")
    $rbUpdate = $dialog.FindName("rbUpdate") 
    $rbReplace = $dialog.FindName("rbReplace")
    $btnOK = $dialog.FindName("btnOK")
    $btnCancel = $dialog.FindName("btnCancel")
    
    # Set up event handlers with direct return values
    $btnOK.Add_Click({
        if ($rbUpdate.IsChecked) {
            $dialog.Tag = "Update"
        } elseif ($rbSkip.IsChecked) {
            $dialog.Tag = "Skip"
        } elseif ($rbReplace.IsChecked) {
            $dialog.Tag = "Replace"
        } else {
            $dialog.Tag = "Skip"  # Default to Skip
        }
        $dialog.DialogResult = $true
        $dialog.Close()
    })
    
    $btnCancel.Add_Click({
        $dialog.Tag = "Cancel"
        $dialog.DialogResult = $false
        $dialog.Close()
    })
    
    # Show dialog and return result
    $null = $dialog.ShowDialog()
    
    $result = $dialog.Tag
    if ([string]::IsNullOrEmpty($result)) {
        return "Cancel"
    }
    
    return $result
}

