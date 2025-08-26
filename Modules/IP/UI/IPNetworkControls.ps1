function Initialize-IPNetworkEventHandlers {
    try {
        # --- Event Handlers ---
        if ($btnAddEntry -ne $null) {
            $btnAddEntry.Add_Click({
                try {
                    # Safely get values with null checks
                    $ipSubnet = ""
                    $vlanId = ""
                    $vlanName = ""
                    $siteName = ""
                    
                    if ($txtIpSubnet -ne $null) { $ipSubnet = $txtIpSubnet.Text.Trim() }
                    if ($txtVlanId -ne $null) { $vlanId = $txtVlanId.Text.Trim() }
                    if ($txtVlanName -ne $null) { $vlanName = $txtVlanName.Text.Trim() }
                    if ($txtSiteName -ne $null) { $siteName = $txtSiteName.Text.Trim() }

                    # Validate all fields are filled
                    if ([string]::IsNullOrEmpty($ipSubnet) -or 
                        [string]::IsNullOrEmpty($vlanId) -or 
                        [string]::IsNullOrEmpty($vlanName) -or 
                        [string]::IsNullOrEmpty($siteName)) {
                        
                        $errorMsg = "Error: All fields must be filled."
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $errorMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                        }
                        Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                        return
                    }

                    # Validate VLAN ID is numeric
                    if (-not (Test-VlanId $vlanId)) {
                        $errorMsg = "Error: VLAN ID must be a valid number between 1 and 4094."
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $errorMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                        }
                        Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                        return
                    }

                    # Validate IP Subnet format (CIDR notation)
                    if (-not (Test-IPSubnetFormat $ipSubnet)) {
                        $errorMsg = "Error: Invalid IP Subnet format. Use CIDR notation (e.g., 10.10.10.0/24)."
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $errorMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                        }
                        Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                        return
                    }

                    # Check for duplicate subnets
                    if ($subnetDataStore -ne $null) {
                        $currentData = @($subnetDataStore.GetAllEntries())
                        if ($currentData | Where-Object { $_.IP_Subnet -eq $ipSubnet }) {
                            $errorMsg = "Error: The subnet '$ipSubnet' already exists in the database."
                            if ($txtBlkImportStatus -ne $null) {
                                $txtBlkImportStatus.Text = $errorMsg
                                $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                            }
                            Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                            return
                        }
                    }

                    # Try to add the entry
                    if (Add-SubnetEntry -IpSubnet $ipSubnet -VlanId ([int]$vlanId) -VlanName $vlanName -SiteName $siteName) {
                        $successMsg = "New subnet added successfully!"
                        
                        # Use safe status update
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $successMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                        }
                        
                        Show-CustomDialog $successMsg "Success" "OK" "Information"
                        
                        # Reset form safely
                        if ($txtIpSubnet -ne $null) { $txtIpSubnet.Text = "" }
                        if ($txtVlanId -ne $null) { $txtVlanId.Text = "" }
                        if ($txtVlanName -ne $null) { $txtVlanName.Text = "" }
                        if ($txtSiteName -ne $null) { $txtSiteName.Text = "" }
                        
                        Update-SubnetDataGridWithSearch
                    }
                    
                } catch {
                    if ($txtBlkImportStatus -ne $null) {
                        $txtBlkImportStatus.Text = "Error adding entry: $_"
                        $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                    }
                    Show-CustomDialog "Error adding entry: $_" "Error" "OK" "Error"
                }
            })
        }

        if ($btnLookup -ne $null) {
            $btnLookup.Add_Click({
                try {
                    $ipToLookup = if ($txtIpLookup -ne $null) { $txtIpLookup.Text.Trim() } else { "" }
                    Lookup-IpAddress -IpAddress $ipToLookup
                } catch {
                    Show-CustomDialog "Error during lookup: $_" "Error" "OK" "Error"
                }
            })
        }

        if ($btnDeleteEntry -ne $null) {
            $btnDeleteEntry.Add_Click({
                try {
                    if ($dgSubnets -eq $null) {
                        return
                    }
                    
                    $selectedItems = @($dgSubnets.SelectedItems)
                    if ($selectedItems.Count -gt 0) {
                        $confirm = Show-CustomDialog "Are you sure you want to delete $($selectedItems.Count) selected entries?" "Confirm Deletion" "YesNo" "Warning"
                        
                        if ($confirm -eq "Yes") {
                            # Get IDs safely
                            $idsToDelete = @()
                            foreach ($item in $selectedItems) {
                                if ($item -ne $null -and $item.ID -ne $null) {
                                    $idsToDelete += $item.ID
                                }
                            }
                            
                            if ($idsToDelete.Count -gt 0) {
                                if ($subnetDataStore.DeleteEntries($idsToDelete)) {
                                    # Clear selection before updating
                                    $dgSubnets.SelectedItems.Clear()
                                    
                                    # Update DataGrid
                                    Update-SubnetDataGridWithSearch
                                    
                                    # Use safe status update
                                    $successMsg = "Successfully deleted $($idsToDelete.Count) entries."
                                    if ($txtBlkImportStatus -ne $null) {
                                        $txtBlkImportStatus.Text = $successMsg
                                        $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                                    }
                                    } else {
                                   $errorMsg = "Error deleting entries."
                                   if ($txtBlkImportStatus -ne $null) {
                                       $txtBlkImportStatus.Text = $errorMsg
                                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                   }
                               }
                           }
                       }
                   } else {
                       $warningMsg = "Please select one or more entries to delete."
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = $warningMsg
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Orange
                       }
                   }
               } catch {
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Error during delete operation: $_"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
                   
                   # Try to recover by refreshing the DataGrid
                   try {
                       Update-SubnetDataGridWithSearch
                   } catch {
                   }
               }
           })
       }

       if ($dgSubnets -ne $null) {
           $dgSubnets.Add_CellEditEnding({
               param($sender, $e)
               
               try {
                   if ($e.EditAction -ne [System.Windows.Controls.DataGridEditAction]::Commit) {
                       return
                   }

                   $editedItem = $e.Row.Item
                   $column = $e.Column.Header
                   $newValue = ""

                   if ($e.EditingElement -is [System.Windows.Controls.TextBox]) {
                       $newValue = $e.EditingElement.Text
                   }

                   # Get all data from the data store
                   $currentData = $subnetDataStore.GetAllEntries()
                   $originalItem = $currentData | Where-Object { $_.ID -eq $editedItem.ID }

                   if ($column -eq $COLUMN_IP_SUBNET -and -not (Test-IPSubnetFormat $newValue)) {
                       Show-CustomDialog "Invalid IP Subnet format! Use CIDR notation (e.g., 10.10.10.0/24)" "Error" "OK" "Error"
                       $e.EditingElement.Text = $originalItem.IP_Subnet
                       $e.Cancel = $true
                       $dgSubnets.CancelEdit()
                       $dgSubnets.CommitEdit()
                       return
                   }
                   
                   if ($column -eq $COLUMN_VLAN_ID -and -not (Test-VlanId $newValue)) {
                       Show-CustomDialog "VLAN ID must be a number between 1 and 4094!" "Error" "OK" "Error"
                       $originalValue = $originalItem.VLAN_ID
                       $e.EditingElement.Text = $originalValue
                       $e.Cancel = $true
                       $dgSubnets.CancelEdit()
                       $dgSubnets.CommitEdit()
                       return
                   }
                   
                   # Check for duplicates only when editing IP Subnet column
                   if ($column -eq $COLUMN_IP_SUBNET) {
                       $duplicateEntry = $currentData | Where-Object { $_.IP_Subnet -eq $newValue -and $_.ID -ne $editedItem.ID }
                       if ($duplicateEntry) {
                           Show-CustomDialog "Error: IP Subnet '$newValue' already exists in another entry." "Duplicate Subnet" "OK" "Error"
                           $originalValue = $originalItem.IP_Subnet
                           $e.EditingElement.Text = $originalValue
                           $e.Cancel = $true
                           $dgSubnets.CancelEdit()
                           $dgSubnets.CommitEdit()
                           return
                       }
                   }
                   
                   # Find the item to update by ID
                   $itemToUpdate = $null
                   $itemIndex = -1
                   
                   for ($i = 0; $i -lt $currentData.Count; $i++) {
                       if ($currentData[$i].ID -eq $editedItem.ID) {
                           $itemToUpdate = $currentData[$i]
                           $itemIndex = $i
                           break
                       }
                   }
                   
                   if ($itemIndex -ge 0) {
                       # Update the specific property based on column
                       switch ($column) {
                           $COLUMN_IP_SUBNET { $currentData[$itemIndex].IP_Subnet = $newValue }
                           $COLUMN_VLAN_ID { $currentData[$itemIndex].VLAN_ID = [int]$newValue }
                           $COLUMN_VLAN_NAME { $currentData[$itemIndex].VLAN_Name = $newValue }
                           $COLUMN_SITE_NAME { $currentData[$itemIndex].Site_Name = $newValue }
                       }
                       
                       # Save the updated data
                       $subnetDataStore.Entries = [System.Collections.Generic.List[SubnetEntry]]::new()
                       $subnetDataStore.Entries.AddRange($currentData)
                       $subnetDataStore.SaveData()
                       
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Entry updated successfully!"
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                       }
                   } else {
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Error: Could not find entry to update."
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                       }
                       $e.Cancel = $true
                   }
               } catch {
                   Show-CustomDialog "Error updating entry: $($_.Exception.Message)" "Error" "OK" "Error"
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Error updating entry: $($_.Exception.Message)"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
                   $e.Cancel = $true
               }
           })

           # DataGrid selection changed event
           $dgSubnets.Add_SelectionChanged({
               try {
                   if ($txtStatusBarSelected -ne $null) {
                       $selectedItems = $dgSubnets.SelectedItems
                       if ($selectedItems.Count -gt 0) {
                           if ($selectedItems.Count -eq 1) {
                               $txtStatusBarSelected.Text = "Selected: $($selectedItems[0].IP_Subnet) (ID: $($selectedItems[0].ID))"
                           } else {
                               $txtStatusBarSelected.Text = "Selected: $($selectedItems.Count) entries"
                           }
                       } else {
                           $txtStatusBarSelected.Text = "Selected: None"
                       }
                   }
               } catch {
               }
           })

           # Enable Delete key to remove selected entries
           $dgSubnets.Add_PreviewKeyDown({
               param($sender, $e)
               
               try {
                   # Check if Delete key was pressed
                   if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                       $selectedItems = @($dgSubnets.SelectedItems)
                       
                       if ($selectedItems.Count -gt 0) {
                           # Call the delete function directly (bypassing the button)
                           if ($btnDeleteEntry -ne $null) {
                               $btnDeleteEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                           }
                           
                           # Mark the event as handled to prevent default behavior
                           $e.Handled = $true
                       }
                   }
               } catch {
               }
           })
       }

       if ($btnBrowseCsv -ne $null) {
           $btnBrowseCsv.Add_Click({
               try {
                   # Disable button during operation
                   $btnBrowseCsv.IsEnabled = $false
                   if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = "Selecting file..." }
                   
                   # Create and configure file dialog
                   $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                   $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
                   $openFileDialog.Title = "Select CSV File to Import"
                   $openFileDialog.CheckFileExists = $false
                   $openFileDialog.CheckPathExists = $false
                   
                   # Show dialog and get result
                   $result = $openFileDialog.ShowDialog()
                   
                   if ($result -eq "OK") {
                       # Just set the path without any validation
                       if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = $openFileDialog.FileName }
                   } else {
                       if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = "" }
                   }
               } catch {
                   Show-CustomDialog "Error selecting file: $($_.Exception.Message)" "Error" "OK" "Error"
                   if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = "" }
               } finally {
                   if ($openFileDialog) {
                       $openFileDialog.Dispose()
                   }
                   $btnBrowseCsv.IsEnabled = $true
               }
           })
       }

       # Optimized Import CSV Button Click Handler
       if ($btnImportCsv -ne $null) {
           $btnImportCsv.Add_Click({
               try {
                   $csvPath = if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text.Trim() } else { "" }
                   if ([string]::IsNullOrEmpty($csvPath)) {
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Please select a CSV file first."
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                       }
                       return
                   }

                   # Disable UI during import
                   $btnImportCsv.IsEnabled = $false
                   if ($btnBrowseCsv -ne $null) { $btnBrowseCsv.IsEnabled = $false }
                   if ($MainTabControl -ne $null) { $MainTabControl.IsEnabled = $false }
                   
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Starting import..."
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Blue
                   }

                   # Show progress panel
                   if ($pnlImportProgress -ne $null) {
                       $pnlImportProgress.Visibility = [System.Windows.Visibility]::Visible
                   }
                   if ($pbImportProgress -ne $null) { $pbImportProgress.Value = 0 }
                   if ($txtProgressStatus -ne $null) { $txtProgressStatus.Text = "Preparing import..." }
                   if ($txtProgressDetails -ne $null) { $txtProgressDetails.Text = "" }
                   
                   $result = Import-SubnetsFromCsv -CsvFilePath $csvPath
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = $result
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                   }
                   Update-SubnetDataGridWithSearch
               } catch {
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Import failed: $($_.Exception.Message)"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
               } finally {
                   # Hide progress panel
                   if ($pnlImportProgress -ne $null) {
                       $pnlImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
                   }
                   
                   # Re-enable UI
                   $btnImportCsv.IsEnabled = $true
                   if ($btnBrowseCsv -ne $null) { $btnBrowseCsv.IsEnabled = $true }
                   if ($MainTabControl -ne $null) { $MainTabControl.IsEnabled = $true }
               }
           })
       }

       if ($btnExportCsv -ne $null) {
           $btnExportCsv.Add_Click({
               try {
                   $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                   $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
                   $saveDialog.FileName = "SubnetExport-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').csv"

                   if ($saveDialog.ShowDialog() -eq "OK") {
                       $btnExportCsv.IsEnabled = $false
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Exporting..."
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Blue
                       }
                       
                       $result = Export-SubnetsToCsv -FilePath $saveDialog.FileName
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = $result
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                       }
                   }
               } catch {
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "ERROR: $($_.Exception.Message)"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
               } finally {
                   $btnExportCsv.IsEnabled = $true
               }
           })
       }

       # --- ENTER KEY SUPPORT FOR LOOKUP TAB ---
       if ($txtIpLookup -ne $null) {
           $txtIpLookup.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       $ipToLookup = $txtIpLookup.Text.Trim()
                       Lookup-IpAddress -IpAddress $ipToLookup
                   }
               } catch {
               }
           })
       }

       # --- ENTER KEY SUPPORT FOR ADD TAB ---
       if ($txtIpSubnet -ne $null) {
           $txtIpSubnet.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       if ($txtVlanId -ne $null) {
           $txtVlanId.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       if ($txtVlanName -ne $null) {
           $txtVlanName.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       if ($txtSiteName -ne $null) {
           $txtSiteName.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       # Search functionality with debouncing
       if ($txtSearch -ne $null) {
           $txtSearch.Add_TextChanged({
               try {
                   $script:IPSearchTimer.Stop()
                   $script:IPSearchTimer.Start()
               } catch {
               }
           })
       }

       if ($btnClearSearch -ne $null) {
           $btnClearSearch.Add_Click({
               try {
                   if ($txtSearch -ne $null) { $txtSearch.Text = "" }
                   Update-SubnetDataGridWithSearch
               } catch {
               }
           })
       }

       if ($MainTabControl -ne $null) {
           $MainTabControl.Add_SelectionChanged({
               try {
                   $selectedTab = $MainTabControl.SelectedItem
                   
                   # Clear import/export status only when leaving Import tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -ne "Import") {
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = ""
                       }
                       if ($txtCsvFilePath -ne $null) {
                           $txtCsvFilePath.Text = ""
                       }
                   }
                   
                   # Clear add entry form AND status only when leaving Manage tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -ne "Manage Subnets") {
                       if ($txtIpSubnet -ne $null) { $txtIpSubnet.Text = "" }
                       if ($txtVlanId -ne $null) { $txtVlanId.Text = "" }
                       if ($txtVlanName -ne $null) { $txtVlanName.Text = "" }
                       if ($txtSiteName -ne $null) { $txtSiteName.Text = "" }
                   }

                   # Clear lookup form AND results only when leaving Lookup tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -ne "Lookup IP") {
                       if ($txtIpLookup -ne $null) { $txtIpLookup.Text = "" }
                       if ($grpLookupResults -ne $null) { $grpLookupResults.Visibility = "Collapsed" }
                       if ($txtBlkSearchedIp -ne $null) { $txtBlkSearchedIp.Text = "" }
                       if ($txtBlkMatchedSubnet -ne $null) { $txtBlkMatchedSubnet.Text = "" }
                       if ($txtBlkVlanId -ne $null) { $txtBlkVlanId.Text = "" }
                       if ($txtBlkVlanName -ne $null) { $txtBlkVlanName.Text = "" }
                       if ($txtBlkSiteName -ne $null) { $txtBlkSiteName.Text = "" }
                   }

                   # Show/hide status bar based on selected tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -eq "Manage Subnets") {
                       if ($MainStatusBar -ne $null) {
                           $MainStatusBar.Visibility = [System.Windows.Visibility]::Visible
                           if ($dgSubnets.ItemsSource -ne $null) {
                               if ($txtStatusBarSubnets -ne $null) {
                                   $txtStatusBarSubnets.Text = "Total Subnets: $($dgSubnets.ItemsSource.Count)"
                               }
                           }
                       }
                   } else {
                       if ($MainStatusBar -ne $null) {
                           $MainStatusBar.Visibility = [System.Windows.Visibility]::Collapsed
                       }
                   }
               } catch {
               }
           })
       }      
   } catch {
   }
}