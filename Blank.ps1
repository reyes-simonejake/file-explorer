# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import the conversion dialog module
. .\ConversionDialog.ps1

Import-Module .\USBStorageModule.psm1

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(1200, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create a splitter for dynamic preview panel
$splitter = New-Object System.Windows.Forms.Splitter
$splitter.Dock = [System.Windows.Forms.DockStyle]::Right
$splitter.Width = 5
$splitter.Visible = $false
$form.Controls.Add($splitter)

# Create Preview Panel with improved visibility control
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Size = New-Object System.Drawing.Size(390, 510)
$previewPanel.Location = New-Object System.Drawing.Point(780, 50)
$previewPanel.BackColor = [System.Drawing.Color]::White
$previewPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                       [System.Windows.Forms.AnchorStyles]::Right -bor `
                       [System.Windows.Forms.AnchorStyles]::Bottom
$previewPanel.Visible = $false
$form.Controls.Add($previewPanel)

# Create Preview Header
$previewHeader = New-Object System.Windows.Forms.Panel
$previewHeader.Height = 30
$previewHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$previewHeader.BackColor = [System.Drawing.Color]::WhiteSmoke
$previewPanel.Controls.Add($previewHeader)

# Create Close Button for Preview Panel
$closePreviewButton = New-Object System.Windows.Forms.Button
$closePreviewButton.Text = "×"
$closePreviewButton.Size = New-Object System.Drawing.Size(30, 30)
$closePreviewButton.Dock = [System.Windows.Forms.DockStyle]::Right
$closePreviewButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$closePreviewButton.Add_Click({
    Hide-PreviewPanel
})
$previewHeader.Controls.Add($closePreviewButton)

# Create Preview Title Label
$previewTitle = New-Object System.Windows.Forms.Label
$previewTitle.Text = "File Preview"
$previewTitle.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$previewTitle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$previewHeader.Controls.Add($previewTitle)

# Create Preview Content Panel
$previewContent = New-Object System.Windows.Forms.Panel
$previewContent.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewPanel.Controls.Add($previewContent)

# Create various preview controls
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.Visible = $false
$previewContent.Controls.Add($pictureBox)

$textPreview = New-Object System.Windows.Forms.RichTextBox
$textPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$textPreview.ReadOnly = $true
$textPreview.Font = New-Object System.Drawing.Font("Consolas", 10)
$textPreview.Visible = $false
$previewContent.Controls.Add($textPreview)

$webBrowser = New-Object System.Windows.Forms.WebBrowser
$webBrowser.Dock = [System.Windows.Forms.DockStyle]::Fill
$webBrowser.Visible = $false
$previewContent.Controls.Add($webBrowser)

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$previewLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$previewContent.Controls.Add($previewLabel)

# Create MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.Controls.Add($menuStrip)

# Initialize navigation history
$global:navigationHistory = New-Object System.Collections.ArrayList
$global:currentIndex = -1

# File Menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "File"
$fileMenu.Text = "File" # Sets the text of the menu item to "File"

$newWindow = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "New Window"
$newWindow.Text = "New Window" # Sets the text of the menu item to "New Window"
$newWindow.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N # Sets the shortcut keys for the menu item to Ctrl+N
$newWindow.Add_Click({
    Start-Process powershell -ArgumentList "-File `"$PSCommandPath`"" # Adds a click event to start a new PowerShell process with the current script
})

$exit = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "Exit"
$exit.Text = "Exit" # Sets the text of the menu item to "Exit"
$exit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4 # Sets the shortcut keys for the menu item to Alt+F4
$exit.Add_Click({ $form.Close() }) # Adds a click event to close the form

$fileMenu.DropDownItems.AddRange(@($newWindow, $exit)) # Adds the "New Window" and "Exit" items to the "File" menu

# Edit Menu
$editMenu = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "Edit"
$editMenu.Text = "Edit" # Sets the text of the menu item to "Edit"

$copy = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "Copy"
$copy.Text = "Copy" # Sets the text of the menu item to "Copy"
$copy.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::C # Sets the shortcut keys for the menu item to Ctrl+C
$copy.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) { # Adds a click event to copy selected items from the list view to the clipboard if any items are selected
        $paths = $listView.SelectedItems | ForEach-Object { $_.Tag } # Retrieves the paths of the selected items
        [System.Windows.Forms.Clipboard]::SetText(($paths -join "`r`n")) # Copies the paths to the clipboard, joined by newlines
    }
})

$paste = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "Paste"
$paste.Text = "Paste" # Sets the text of the menu item to "Paste"
$paste.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::V # Sets the shortcut keys for the menu item to Ctrl+V
$paste.Add_Click({
    if ([System.Windows.Forms.Clipboard]::ContainsText()) { # Adds a click event to paste items from the clipboard to the current directory if the clipboard contains text
        $paths = [System.Windows.Forms.Clipboard]::GetText() -split "`r`n" # Retrieves and splits the paths from the clipboard by newlines
        foreach ($path in $paths) {
            if (Test-Path $path) { # Checks if each path exists
                $destination = Join-Path $global:currentPath (Split-Path $path -Leaf) # Joins the current path with the leaf part of the path
                Copy-Item -Path $path -Destination $destination -Recurse # Copies the item to the destination
            }
        }
        Populate-ListView $global:currentPath # Refreshes the list view
    }
})

$delete = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "Delete"
$delete.Text = "Delete" # Sets the text of the menu item to "Delete"
$delete.ShortcutKeys = [System.Windows.Forms.Keys]::Delete # Sets the shortcut key for the menu item to Delete
$delete.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) { # Adds a click event to delete selected items from the list view if any items are selected
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete the selected items?", # Displays a confirmation message box
            "Confirm Delete", # Sets the title of the message box
            [System.Windows.Forms.MessageBoxButtons]::YesNo, # Adds Yes and No buttons to the message box
            [System.Windows.Forms.MessageBoxIcon]::Warning # Sets the icon of the message box to a warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { # If the user clicks Yes
            $listView.SelectedItems | ForEach-Object {
                Remove-Item $_.Tag -Recurse -Force # Deletes the selected items
            }
            Populate-ListView $global:currentPath # Refreshes the list view
        }
    }
})

$editMenu.DropDownItems.AddRange(@($copy, $paste, $delete)) # Adds the "Copy," "Paste," and "Delete" items to the "Edit" menu


# View Menu
$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "View"
$viewMenu.Text = "View" # Sets the text of the menu item to "View"

$refresh = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item object for "Refresh"
$refresh.Text = "Refresh" # Sets the text of the menu item to "Refresh"
$refresh.ShortcutKeys = [System.Windows.Forms.Keys]::F5 # Sets the shortcut key for the menu item to F5
$refresh.Add_Click({
    Populate-ListView $global:currentPath # Adds a click event to refresh the list view with the current path
    Populate-TreeView # Refreshes the tree view
})

$viewMenu.DropDownItems.Add($refresh) # Adds the "Refresh" item to the "View" menu

# Create Sort Menu Button
$sortMenu = New-Object System.Windows.Forms.ToolStripDropDownButton # Creates a new drop-down button for sorting
$sortMenu.Text = "Sort" # Sets the text of the drop-down button to "Sort"
$sortMenu.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Text # Sets the display style to text

# Create Sort Options Menu Items
$sortByName = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item for sorting by name
$sortByName.Text = "Name" # Sets the text of the menu item to "Name"
$sortByName.Checked = $true # Sets the menu item as checked by default
$global:currentSortColumn = "Name" # Sets the global current sort column to "Name"
$global:sortAscending = $true # Sets the global sort direction to ascending

$sortByType = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item for sorting by type
$sortByType.Text = "Type" # Sets the text of the menu item to "Type"

$sortByDateModified = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item for sorting by date modified
$sortByDateModified.Text = "Date modified" # Sets the text of the menu item to "Date modified"

$sortBySize = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item for sorting by size
$sortBySize.Text = "Size" # Sets the text of the menu item to "Size"

# Create Sort Direction Options
$separator = New-Object System.Windows.Forms.ToolStripSeparator # Creates a separator for the sort options
$ascending = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item for ascending sort direction
$ascending.Text = "Ascending" # Sets the text of the menu item to "Ascending"
$ascending.Checked = $true # Sets the menu item as checked by default

$descending = New-Object System.Windows.Forms.ToolStripMenuItem # Creates a new menu item for descending sort direction
$descending.Text = "Descending" # Sets the text of the menu item to "Descending"

# Add items to Sort menu
$sortMenu.DropDownItems.AddRange(@(
    $sortByName, # Adds the "Name" sort option
    $sortByType, # Adds the "Type" sort option
    $sortByDateModified, # Adds the "Date modified" sort option
    $sortBySize, # Adds the "Size" sort option
    $separator, # Adds the separator
    $ascending, # Adds the "Ascending" sort direction option
    $descending # Adds the "Descending" sort direction option
))

# Add Sort menu to MenuStrip (after View menu)
$menuStrip.Items.Add($sortMenu) # Adds the sort menu to the menu strip

# Function to update sort checks
function Update-SortChecks {
    param($selectedItem) # Takes the selected item as a parameter
    
    # Unchecks all sort options
    $sortByName.Checked = $false
    $sortByType.Checked = $false
    $sortByDateModified.Checked = $false
    $sortBySize.Checked = $false
    
    # Checks the selected sort option
    $selectedItem.Checked = $true
    $global:currentSortColumn = $selectedItem.Text # Sets the global current sort column to the selected item's text
}

# Function to update direction checks
function Update-DirectionChecks {
    param($isAscending) # Takes the sort direction as a parameter
    
    # Updates the checked state of the sort direction options
    $ascending.Checked = $isAscending
    $descending.Checked = -not $isAscending
    $global:sortAscending = $isAscending # Sets the global sort direction
}

# Function to sort ListView items
function Sort-ListView {
    param (
        [string]$column, # Takes the column to sort by as a parameter
        [bool]$ascending # Takes the sort direction as a parameter
    )
    
    $items = @($listView.Items) # Gets the list view items
    
    # Sorts the items based on the specified column and direction
    switch ($column) {
        "Name" {
            $sorted = if ($ascending) {
                $items | Sort-Object { $_.Text } # Sorts by name (ascending)
            } else {
                $items | Sort-Object { $_.Text } -Descending # Sorts by name (descending)
            }
        }
        "Type" {
            $sorted = if ($ascending) {
                $items | Sort-Object { $_.SubItems[1].Text } # Sorts by type (ascending)
            } else {
                $items | Sort-Object { $_.SubItems[1].Text } -Descending # Sorts by type (descending)
            }
        }
        "Date modified" {
            $sorted = if ($ascending) {
                $items | Sort-Object { [DateTime]::Parse($_.SubItems[3].Text) } # Sorts by date modified (ascending)
            } else {
                $items | Sort-Object { [DateTime]::Parse($_.SubItems[3].Text) } -Descending # Sorts by date modified (descending)
            }
        }
        "Size" {
            $sorted = if ($ascending) {
                $items | Sort-Object {
                    if ($_.SubItems[2].Text -eq "") { 0 } # If size is empty, treats it as 0
                    else {
                        $size = $_.SubItems[2].Text
                        switch -Regex ($size) {
                            "(\d+\.?\d*)\s*B" { [double]$matches[1] } # Sorts by size in bytes
                            "(\d+\.?\d*)\s*KB" { [double]$matches[1] * 1KB } # Sorts by size in kilobytes
                            "(\d+\.?\d*)\s*MB" { [double]$matches[1] * 1MB } # Sorts by size in megabytes
                            "(\d+\.?\d*)\s*GB" { [double]$matches[1] * 1GB } # Sorts by size in gigabytes
                            "(\d+\.?\d*)\s*TB" { [double]$matches[1] * 1TB } # Sorts by size in terabytes
                            default { 0 } # Default case
                        }
                    }
                }
            } else {
                $items | Sort-Object {
                    if ($_.SubItems[2].Text -eq "") { 0 } # If size is empty, treats it as 0
                    else {
                        $size = $_.SubItems[2].Text
                        switch -Regex ($size) {
                            "(\d+\.?\d*)\s*B" { [double]$matches[1] } # Sorts by size in bytes
                            "(\d+\.?\d*)\s*KB" { [double]$matches[1] * 1KB } # Sorts by size in kilobytes
                            "(\d+\.?\d*)\s*MB" { [double]$matches[1] * 1MB } # Sorts by size in megabytes
                            "(\d+\.?\d*)\s*GB" { [double]$matches[1] * 1GB } # Sorts by size in gigabytes
                            "(\d+\.?\d*)\s*TB" { [double]$matches[1] * 1TB } # Sorts by size in terabytes
                            default { 0 } # Default case
                        }
                    }
                } -Descending
            }
        }
    }
    
    $listView.BeginUpdate() # Begins updating the list view
    $listView.Items.Clear() # Clears the current items
    $listView.Items.AddRange($sorted) # Adds the sorted items
    $listView.EndUpdate() # Ends updating the list view
}

# Event handlers for sort options
$sortByName.Add_Click({
    Update-SortChecks $sortByName # Calls the function to update the sort checks for "Name"
    Sort-ListView "Name" $global:sortAscending # Sorts the ListView items by "Name" in the current sort direction
})

$sortByType.Add_Click({
    Update-SortChecks $sortByType # Calls the function to update the sort checks for "Type"
    Sort-ListView "Type" $global:sortAscending # Sorts the ListView items by "Type" in the current sort direction
})

$sortByDateModified.Add_Click({
    Update-SortChecks $sortByDateModified # Calls the function to update the sort checks for "Date modified"
    Sort-ListView "Date modified" $global:sortAscending # Sorts the ListView items by "Date modified" in the current sort direction
})

$sortBySize.Add_Click({
    Update-SortChecks $sortBySize # Calls the function to update the sort checks for "Size"
    Sort-ListView "Size" $global:sortAscending # Sorts the ListView items by "Size" in the current sort direction
})

# Event handlers for sort direction
$ascending.Add_Click({
    Update-DirectionChecks $true # Calls the function to update the direction checks to ascending
    Sort-ListView $global:currentSortColumn $true # Sorts the ListView items by the current sort column in ascending order
})

$descending.Add_Click({
    Update-DirectionChecks $false # Calls the function to update the direction checks to descending
    Sort-ListView $global:currentSortColumn $false # Sorts the ListView items by the current sort column in descending order
})

function Show-FileExplorer {
    param (
        [Parameter(Position=0)]
        [string]$Path = (Get-Location), # Default path is the current location
        
        [Parameter()]
        [ValidateSet("ExtraLargeIcons", "LargeIcons", "MediumIcons", "SmallIcons", "List", "Details")]
        [string]$View = "Details", # Default view is "Details"
        
        [Parameter()]
        [switch]$ShowHidden # Optional switch to show hidden files
    )
    
    # Ensure the path exists
    if (-not (Test-Path $Path)) {
        Write-Error "Path '$Path' does not exist."
        return
    }
    
    # Format parameters for Get-ChildItem
    $params = @{
        Path = $Path
        Force = $ShowHidden
    }
    
    # Get items
    $items = Get-ChildItem @params | Select-Object Mode, LastWriteTime, Length, Name, Extension
    
    # Custom formatting based on view type
    switch ($View) {
        "ExtraLargeIcons" {
            Write-Host "╔════ Extra Large Icons View ════╗"
            $items | ForEach-Object {
                Write-Host ("║ {0,-50} ║" -f $_.Name)
            }
            Write-Host "╚══════════════════════════════════╝"
        }
        "LargeIcons" {
            Write-Host "╔════ Large Icons View ════╗"
            $items | ForEach-Object {
                Write-Host ("║ {0,-40} ║" -f $_.Name)
            }
            Write-Host "╚════════════════════════════╝"
        }
        "MediumIcons" {
            Write-Host "╔════ Medium Icons View ════╗"
            $items | Format-Wide -Column 2 -Property Name
            Write-Host "╚════════════════════════════╝"
        }
        "SmallIcons" {
            Write-Host "╔════ Small Icons View ════╗"
            $items | Format-Wide -Column 3 -Property Name
            Write-Host "╚═══════════════════════════╝"
        }
        "List" {
            $items | Format-Wide -Column 1 -Property Name
        }
        "Details" {
            $items | Format-Table -Property @(
                @{Label="Type"; Expression={$_.Mode}},
                @{Label="Last Modified"; Expression={$_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")}},
                @{Label="Size"; Expression={
                    if ($_.Length -ge 1GB) { "{0:N2} GB" -f ($_.Length / 1GB) }
                    elseif ($_.Length -ge 1MB) { "{0:N2} MB" -f ($_.Length / 1MB) }
                    elseif ($_.Length -ge 1KB) { "{0:N2} KB" -f ($_.Length / 1KB) }
                    else { "{0} B" -f $_.Length }
                }},
                @{Label="Name"; Expression={$_.Name}}
            ) -AutoSize
        }
    }
    
    # Display current path and item count
    Write-Host "`nCurrent Path: $Path"
    Write-Host "Total Items: $($items.Count)"
}

# Add tab completion for the View parameter
Register-ArgumentCompleter -CommandName Show-FileExplorer -ParameterName View -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    @('ExtraLargeIcons', 'LargeIcons', 'MediumIcons', 'SmallIcons', 'List', 'Details') | Where-Object {
        $_ -like "$wordToComplete*"
    }
}

# Create View tab/menu
$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem("View")
# Add view options as dropdown items
$viewOptions = @(
    "Extra large icons",
    "Large icons",
    "Medium-sized icons",
    "Small icons",
    "List",
    "Details"
)

foreach ($option in $viewOptions) {
    $viewMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewMenuItem.Text = $option
    $viewMenuItem.Add_Click({
        $selected = $this.Text
        # Uncheck all items
        $viewMenu.DropDownItems | ForEach-Object { $_.Checked = $false }
        # Check the selected item
        $this.Checked = $true
        
        # Update the view based on selection
        $listView.View = switch ($selected) {
            "Extra large icons" { [System.Windows.Forms.View]::LargeIcon }
            "Large icons" { [System.Windows.Forms.View]::LargeIcon }
            "Medium-sized icons" { [System.Windows.Forms.View]::LargeIcon }
            "Small icons" { [System.Windows.Forms.View]::SmallIcon }
            "List" { [System.Windows.Forms.View]::List }
            "Details" { [System.Windows.Forms.View]::Details }
        }
    })
    $viewMenu.DropDownItems.Add($viewMenuItem)
}

# Insert View menu after Sort in the MenuStrip
# Assuming $menuStrip is your MenuStrip control and Sort is the second item
$menuStrip.Items.Insert(2, $viewMenu)

# Create Navigation Buttons
$btnBack = New-Object System.Windows.Forms.ToolStripMenuItem
$btnBack.Text = "←"
$btnBack.Enabled = $false
$btnBack.Add_Click({
    if ($global:currentIndex -gt 0) {
        $global:currentIndex--
        $previousPath = $global:navigationHistory[$global:currentIndex]
        Populate-ListView $previousPath
        Update-NavigationButtons
    }
})

$btnForward = New-Object System.Windows.Forms.ToolStripMenuItem
$btnForward.Text = "→"
$btnForward.Enabled = $false
$btnForward.Add_Click({
    if ($global:currentIndex -lt $global:navigationHistory.Count - 1) {
        $global:currentIndex++
        $nextPath = $global:navigationHistory[$global:currentIndex]
        Populate-ListView $nextPath
        Update-NavigationButtons
    }
})

$btnUp = New-Object System.Windows.Forms.ToolStripMenuItem
$btnUp.Text = "↑"
$btnUp.Add_Click({
    $parentPath = Split-Path $global:currentPath -Parent
    if ($parentPath) {
        Navigate-To $parentPath
    }
})

$btnDown = New-Object System.Windows.Forms.ToolStripMenuItem
$btnDown.Text = "↓"
$btnDown.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        if ($selectedItem -and (Test-Path -Path $selectedItem.Tag -PathType Container)) {
            Navigate-To $selectedItem.Tag
        }
    }
})

# Add all items to MenuStrip in order
$menuStrip.Items.AddRange(@($btnBack, $btnForward, $btnUp, $fileMenu, $sortMenu, $viewMenu, $editMenu))

# Create address bar in MenuStrip
$addressBar = New-Object System.Windows.Forms.ToolStripTextBox
$addressBar.Size = New-Object System.Drawing.Size(400, 25)
$addressBar.Name = "AddressBar"
$addressBar.AutoSize = $false
$addressBar.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

# Add event handler for address bar
$addressBar.Add_KeyPress({
    param($sender, $e)
    if ($e.KeyChar -eq [System.Windows.Forms.Keys]::Enter) {
        $e.Handled = $true
        $path = $addressBar.Text.Trim()
        if (Test-Path -Path $path) {
            Navigate-To $path
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Invalid path: $path",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Create search box in MenuStrip with modified style
$searchBox = New-Object System.Windows.Forms.ToolStripTextBox
$searchBox.Size = New-Object System.Drawing.Size(200, 25)
$searchBox.Name = "SearchBox"
$searchBox.PlaceholderText = "Search..."
$searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$searchBox.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Create search button in MenuStrip with modified style
$searchButton = New-Object System.Windows.Forms.ToolStripButton
$searchButton.Text = "Search"
$searchButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Text
$searchButton.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$searchButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)

# Add controls to MenuStrip
$menuStrip.Items.Add($addressBar)

# Add a small spring after address bar
$springAfterAddress = New-Object System.Windows.Forms.ToolStripStatusLabel
$springAfterAddress.Spring = $true
$springAfterAddress.Width = 20
$menuStrip.Items.Add($springAfterAddress)

# Add search box and button to MenuStrip
$menuStrip.Items.Add($searchBox)
$menuStrip.Items.Add($searchButton)

# Search button click event
$searchButton.Add_Click({
    Search-Files $searchBox.Text
})

# Search box key press event
$searchBox.Add_KeyPress({
    param($sender, $e)
    if ($e.KeyChar -eq [System.Windows.Forms.Keys]::Enter) {
        $e.Handled = $true
        Search-Files $searchBox.Text
    }
})

# Optional: Add hover effect to search button
$searchButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
})
$searchButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
})

# Create Quick Access panel
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Size = New-Object System.Drawing.Size(250, 220)
$quickAccessPanel.Location = New-Object System.Drawing.Point(10, 50)
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoSize = $false
$quickAccessPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                           [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($quickAccessPanel)
c
# Create TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 290)
$treeView.Location = New-Object System.Drawing.Point(10, 270)
$treeView.Scrollable = $true
$treeView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                   [System.Windows.Forms.AnchorStyles]::Left -bor `
                   [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($treeView)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(500, 510)
$listView.Location = New-Object System.Drawing.Point(270, 50)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                   [System.Windows.Forms.AnchorStyles]::Left -bor `
                   [System.Windows.Forms.AnchorStyles]::Right -bor `
                   [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($listView)

# Add columns to ListView
$columns = @(
    @{Name="Name"; Width=250},
    @{Name="Type"; Width=100},
    @{Name="Size"; Width=100}, 
    @{Name="Modified"; Width=150}
)

foreach ($column in $columns) {
    $listView.Columns.Add($column.Name, $column.Width)
}
#Search-function
function Search-Files {
    param ([string]$searchTerm)
    
    if ([string]::IsNullOrWhiteSpace($searchTerm)) {
        Populate-ListView $global:currentPath
        return
    }
    
    $listView.Items.Clear()
    
    try {
        $searchResults = Get-ChildItem -Path $global:currentPath -Recurse -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -like "*$searchTerm*" }
        
        foreach ($item in $searchResults) {
            $listViewItem = New-Object System.Windows.Forms.ListViewItem
            $listViewItem.Text = $item.Name
            
            # Add subitems
            $listViewItem.SubItems.Add($item.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
            
            if ($item.PSIsContainer) {
                $listViewItem.SubItems.Add("Folder")
                $listViewItem.SubItems.Add("")
                $listViewItem.ImageIndex = 0  # Assuming 0 is your folder icon index
            } else {
                $listViewItem.SubItems.Add("File")
                $listViewItem.SubItems.Add([string]::Format("{0:N2} KB", $item.Length / 1KB))
                $listViewItem.ImageIndex = 1  # Assuming 1 is your file icon index
            }
            
            # Store the full path in the Tag property
            $listViewItem.Tag = $item.FullName
            
            # Add the item to the ListView
            $listView.Items.Add($listViewItem)
        }
        
        # Update status bar with search results count
        $statusMessage = "Found $($searchResults.Count) items matching '$searchTerm'"
        Update-StatusBar -path $statusMessage
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error performing search: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to format file size
function Format-FileSize {
    param ([long]$size)
    if ($size -lt 1KB) { return "$size B" }
    elseif ($size -lt 1MB) { return "{0:N2} KB" -f ($size/1KB) }
    elseif ($size -lt 1GB) { return "{0:N2} MB" -f ($size/1MB) }
    elseif ($size -lt 1TB) { return "{0:N2} GB" -f ($size/1GB) }
    else { return "{0:N2} TB" -f ($size/1TB) }
}

# Function to update navigation buttons
function Update-NavigationButtons {
    $btnBack.Enabled = $global:currentIndex -gt 0
    $btnForward.Enabled = $global:currentIndex -lt ($global:navigationHistory.Count - 1)
    $btnUp.Enabled = (Split-Path $global:currentPath -Parent) -ne $null
}

# Function to handle navigation
function Navigate-To {
    param ([string]$path)
    
    if (Test-Path $path) {
        $global:currentIndex++
        if ($global:currentIndex -lt $global:navigationHistory.Count) {
            $global:navigationHistory.RemoveRange($global:currentIndex, $global:navigationHistory.Count - $global:currentIndex)
        }
        [void]$global:navigationHistory.Add($path)
        Populate-ListView $path
        Update-NavigationButtons
    }
}

# Function to show preview panel
function Show-PreviewPanel {
    if (-not $previewPanel.Visible) {
        $previewPanel.Visible = $true
        $splitter.Visible = $true
        $listView.Width -= ($previewPanel.Width + $splitter.Width)
    }
}

# Function to hide preview panel
function Hide-PreviewPanel {
    if ($previewPanel.Visible) {
        $listView.Width += ($previewPanel.Width + $splitter.Width)
        $previewPanel.Visible = $false
        $splitter.Visible = $false
    }
}

# Function to clear preview
function Clear-Preview {
    if ($pictureBox.Image -ne $null) { $pictureBox.Image.Dispose() }  # Dispose of the image if it is currently loaded in the PictureBox
    $pictureBox.Image = $null
    $pictureBox.Visible = $false
    $textPreview.Clear()
    $textPreview.Visible = $false
    $webBrowser.Visible = $false
    $previewLabel.Visible = $true
    $previewLabel.Text = "Select a file to preview"
}

# Function to preview file with enhanced format support
function Show-FilePreview {
    param ([string]$filePath)
    
    Clear-Preview
    Show-PreviewPanel
    
    if (-not (Test-Path $filePath)) {
        return
    }
    
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    $previewTitle.Text = $fileName
    $previewLabel.Visible = $false
    
    switch -Regex ($extension) {
        # Image files
        '\.(jpg|jpeg|png|gif|bmp|ico|tiff)$' {
            try {
                $image = [System.Drawing.Image]::FromFile($filePath)
                $pictureBox.Image = $image
                $pictureBox.Visible = $true
                $previewLabel.Text = "Size: $($image.Width)x$($image.Height)"
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error loading image"
                $previewLabel.Visible = $true
            }
        }
        
        # Text files
        '\.(txt|log|ps1|cmd|bat|csv|json|xml|html|css|js|md|yml|yaml|ini|conf|cfg|reg)$' {
            try {
                $content = Get-Content -Path $filePath -Raw -ErrorAction Stop
                $textPreview.Text = $content
                $textPreview.Visible = $true
                
                # Syntax highlighting based on extension
                switch -Regex ($extension) {
                    '\.(ps1|cmd|bat)$' {
                        # PowerShell/Batch highlighting (basic)
                        $keywords = @('function', 'param', 'if', 'else', 'while', 'foreach', 'return', 'try', 'catch')
                        foreach ($keyword in $keywords) {
                            $textPreview.SelectionColor = [System.Drawing.Color]::Blue
                        }
                    }
                    '\.(json|xml|html|css)$' {
                        # Web format highlighting (basic)
                        $webBrowser.Navigate($filePath)
                        $webBrowser.Visible = $true
                        $textPreview.Visible = $false
                    }
                }
            }
            catch {
                $previewLabel.Text = "Error loading text file"
                $previewLabel.Visible = $true
            }
        }
        
        # Office documents and PDFs
        '\.(doc|docx|xls|xlsx|ppt|pptx|pdf)$' {
            $fileInfo = Get-Item $filePath
            $previewLabel.Text = @"
File Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $fileInfo.Length)
Created: $($fileInfo.CreationTime)
Modified: $($fileInfo.LastWriteTime)
"@
            $previewLabel.Visible = $true
        }
        
        # Audio files
        '\.(mp3|wav|wma|m4a|aac)$' {
            $previewLabel.Text = @"
Audio File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Video files
        '\.(mp4|avi|mkv|wmv|mov)$' {
            $previewLabel.Text = @"
Video File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Archive files
        '\.(zip|rar|7z|tar|gz)$' {
            try {
                $archive = Get-Item $filePath
                $previewLabel.Text = @"
Archive File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $archive.Length)
Created: $($archive.CreationTime)
Modified: $($archive.LastWriteTime)
"@
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error reading archive"
                $previewLabel.Visible = $true
            }
        }
        
        default {
            $previewLabel.Text = "Preview not available for this file type"
            $previewLabel.Visible = $true
        }
    }
}

# Create StatusStrip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$form.Controls.Add($statusStrip)
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom

# Create status bar labels
$statusItemCount = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusItemCount.Text = "0 items"
$statusItemCount.Spring = $true

$statusTotalSize = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusTotalSize.Text = "Total size: 0 bytes"

$statusSelectedItems = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusSelectedItems.Text = "0 selected"

# Add labels to StatusStrip
$statusStrip.Items.AddRange(@($statusItemCount, $statusTotalSize, $statusSelectedItems))

# Function to update status bar
function Update-StatusBar {
    param (
        [string]$path
    )

    try {
        # Get all items in the current directory
        $items = Get-ChildItem -Path $path -ErrorAction Stop

        # Calculate total number of items
        $totalItems = $items.Count
        $totalFiles = ($items | Where-Object { -not $_.PSIsContainer }).Count
        $totalFolders = ($items | Where-Object { $_.PSIsContainer }).Count

        # Calculate total size of files
        $totalSize = ($items | Where-Object { -not $_.PSIsContainer } | Measure-Object -Property Length -Sum).Sum

        # Update status labels
        $statusItemCount.Text = "$totalItems items ($totalFiles files, $totalFolders folders)"
        $statusTotalSize.Text = "Total size: $(Format-FileSize $totalSize)"
    }
    catch {
        $statusItemCount.Text = "0 items"
        $statusTotalSize.Text = "Total size: 0 bytes"
    }

    # Update selected items
    $selectedCount = $listView.SelectedItems.Count
    if ($selectedCount -gt 0) {
        $selectedSize = ($listView.SelectedItems | ForEach-Object { 
            $path = $_.Tag
            if (Test-Path -Path $path -PathType Leaf) {
                (Get-Item $path).Length 
            } else { 0 }
        } | Measure-Object -Sum).Sum

        $statusSelectedItems.Text = "$selectedCount selected ($(Format-FileSize $selectedSize))"
    }
    else {
        $statusSelectedItems.Text = "0 selected"
    }
}

# Function to populate the ListView
function Populate-ListView {
    param ([string]$path)
    
    $global:currentPath = $path

    $addressBar.Text = $path

    $listView.Items.Clear()
    
    try {
        $items = Get-ChildItem -Path $path -ErrorAction Stop
        
        foreach ($item in $items) {
            $listViewItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
            
            if ($item.PSIsContainer) {
                $type = "Folder"
                $size = ""
            } else {
                $type = if ($item.Extension) { $item.Extension.TrimStart(".").ToUpper() } else { "File" }
                $size = Format-FileSize $item.Length
            }
            
            $listViewItem.SubItems.Add($type)
            $listViewItem.SubItems.Add($size)
            $listViewItem.SubItems.Add($item.LastWriteTime.ToString("g"))
            $listViewItem.Tag = $item.FullName
            
            $listView.Items.Add($listViewItem)
        }
        
        Update-NavigationButtons
        Update-StatusBar -path $path

    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error accessing path: $path`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to create Quick Access buttons
function Add-QuickAccessButton($text, $path) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Width = 240
    $button.Height = 30
    $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.Tag = $path
    
    $button.Add_Click({
        $buttonPath = $this.Tag
        if (Test-Path -Path $buttonPath) {
            Navigate-To -path $buttonPath
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Path does not exist: $buttonPath",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    
    $quickAccessPanel.Controls.Add($button)
    return $button
}

# Add Quick Access buttons
$quickAccessButtons = @{
    "Desktop" = [Environment]::GetFolderPath("Desktop")
    "Downloads" = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    "Documents" = [Environment]::GetFolderPath("MyDocuments")
    "Music" = [Environment]::GetFolderPath("MyMusic")
    "Pictures" = [Environment]::GetFolderPath("MyPictures")
    "Videos" = [Environment]::GetFolderPath("MyVideos")
}

foreach ($button in $quickAccessButtons.GetEnumerator()) {
    Add-QuickAccessButton $button.Key $button.Value
}




# Function to populate the TreeView
function Populate-TreeView {
    $treeView.Nodes.Clear()
    
    $thisPC = $treeView.Nodes.Add("This PC")
    
    Get-PSDrive -PSProvider FileSystem | ForEach-Object {
        $driveNode = $thisPC.Nodes.Add($_.Root)
        $driveNode.Tag = $_.Root
        try {
            Get-ChildItem -Path $_.Root -Directory -ErrorAction Stop | ForEach-Object {
                $subNode = $driveNode.Nodes.Add($_.Name)
                $subNode.Tag = $_.FullName
            }
        } catch {}
    }
    
    $thisPC.Expand()
}

# Event handler for TreeView node click
$treeView.add_AfterSelect({
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode.Tag) {
        Navigate-To -path $selectedNode.Tag
    }
})

# Event handler for ListView double-click
$listView.add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $itemPath = $selectedItem.Tag
        
        if (Test-Path -Path $itemPath -PathType Container) {
            Navigate-To -path $itemPath
        } else {
            Start-Process $itemPath
        }
    }
})

# Create Preview Panel
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Size = New-Object System.Drawing.Size(390, 510)
$previewPanel.Location = New-Object System.Drawing.Point(780, 50)
$previewPanel.BackColor = [System.Drawing.Color]::White
$previewPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                       [System.Windows.Forms.AnchorStyles]::Right -bor `
                       [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($previewPanel)

# Create Preview Controls
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Size = New-Object System.Drawing.Size(380, 380)
$pictureBox.Location = New-Object System.Drawing.Point(5, 5)
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.Visible = $false
$previewPanel.Controls.Add($pictureBox)

$textPreview = New-Object System.Windows.Forms.RichTextBox
$textPreview.Size = New-Object System.Drawing.Size(380, 480)
$textPreview.Location = New-Object System.Drawing.Point(5, 5)
$textPreview.ReadOnly = $true
$textPreview.Font = New-Object System.Drawing.Font("Consolas", 10)
$textPreview.Visible = $false
$previewPanel.Controls.Add($textPreview)

$mediaPlayer = New-Object System.Windows.Forms.Panel
$mediaPlayer.Size = New-Object System.Drawing.Size(380, 380)
$mediaPlayer.Location = New-Object System.Drawing.Point(5, 5)
$mediaPlayer.Visible = $false
$previewPanel.Controls.Add($mediaPlayer)

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Size = New-Object System.Drawing.Size(380, 40)
$previewLabel.Location = New-Object System.Drawing.Point(5, 5)
$previewLabel.Text = "Select a file to preview"
$previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$previewLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$previewPanel.Controls.Add($previewLabel)

# Function to clear preview
function Clear-Preview {
    $pictureBox.Image = $null
    $pictureBox.Visible = $false
    $textPreview.Clear()
    $textPreview.Visible = $false
    $mediaPlayer.Visible = $false
    $previewLabel.Visible = $true
    $previewLabel.Text = "Select a file to preview"
}

function Show-FilePreview {
    param ([string]$filePath)
    
    Clear-Preview
    Show-PreviewPanel
    
    if (-not (Test-Path $filePath)) {
        return
    }
    
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    $previewTitle.Text = $fileName
    $previewLabel.Visible = $false
    
    # Add folder preview at the beginning
    if (Test-Path -Path $filePath -PathType Container) {
        try {
            $folder = Get-Item $filePath
            $items = Get-ChildItem $filePath
            $fileCount = ($items | Where-Object { -not $_.PSIsContainer }).Count
            $folderCount = ($items | Where-Object { $_.PSIsContainer }).Count
            
            $previewLabel.Text = @"
Folder: $($folder.Name)
Created: $($folder.CreationTime)
Modified: $($folder.LastWriteTime)
Contains: $fileCount files, $folderCount folders
"@
            $previewLabel.Visible = $true
            return
        }
        catch {
            $previewLabel.Text = "Error reading folder"
            $previewLabel.Visible = $true
            return
        }
    }
    
    switch -Regex ($extension) {
        # Image files
        '\.(jpg|jpeg|png|gif|bmp|ico|tiff)$' {
            try {
                $image = [System.Drawing.Image]::FromFile($filePath)
                $pictureBox.Image = $image
                $pictureBox.Visible = $true
                $previewLabel.Text = "Size: $($image.Width)x$($image.Height)"
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error loading image"
                $previewLabel.Visible = $true
            }
        }
        
        # Text files
        '\.(txt|log|ps1|cmd|bat|csv|json|xml|html|css|js|md|yml|yaml|ini|conf|cfg|reg)$' {
            try {
                $content = Get-Content -Path $filePath -Raw -ErrorAction Stop
                $textPreview.Text = $content
                $textPreview.Visible = $true
                
                # Syntax highlighting based on extension
                switch -Regex ($extension) {
                    '\.(ps1|cmd|bat)$' {
                        # PowerShell/Batch highlighting (basic)
                        $keywords = @('function', 'param', 'if', 'else', 'while', 'foreach', 'return', 'try', 'catch')
                        foreach ($keyword in $keywords) {
                            $textPreview.SelectionColor = [System.Drawing.Color]::Blue
                        }
                    }
                    '\.(json|xml|html|css)$' {
                        # Web format highlighting (basic)
                        $webBrowser.Navigate($filePath)
                        $webBrowser.Visible = $true
                        $textPreview.Visible = $false
                    }
                }
            }
            catch {
                $previewLabel.Text = "Error loading text file"
                $previewLabel.Visible = $true
            }
        }
        
        # Office documents and PDFs
        '\.(doc|docx|xls|xlsx|ppt|pptx|pdf)$' {
            $fileInfo = Get-Item $filePath
            $previewLabel.Text = @"
File Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $fileInfo.Length)
Created: $($fileInfo.CreationTime)
Modified: $($fileInfo.LastWriteTime)
"@
            $previewLabel.Visible = $true
        }
        
        # Audio files
        '\.(mp3|wav|wma|m4a|aac)$' {
            $previewLabel.Text = @"
Audio File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Video files
        '\.(mp4|avi|mkv|wmv|mov)$' {
            $previewLabel.Text = @"
Video File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Archive files
        '\.(zip|rar|7z|tar|gz)$' {
            try {
                $archive = Get-Item $filePath
                $previewLabel.Text = @"
Archive File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $archive.Length)
Created: $($archive.CreationTime)
Modified: $($archive.LastWriteTime)
"@
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error reading archive"
                $previewLabel.Visible = $true
            }
        }
        
        default {
            $previewLabel.Text = "Preview not available for this file type"
            $previewLabel.Visible = $true
        }
    }
}

# Event handler for ListView selection changed
$listView.add_SelectedIndexChanged({
    Update-NavigationButtons
    Update-StatusBar -path $global:currentPath
})


# Event handler for ListView click
$listView.add_MouseClick({
    param($sender, $e)
    
    $item = $listView.GetItemAt($e.X, $e.Y)
    if ($item -ne $null) {
        $itemPath = $item.Tag
        Show-FilePreview -filePath $itemPath
    }
})

# Event handler for ListView selection changed
$listView.add_SelectedIndexChanged({
    Update-NavigationButtons
})

# Event handler for ListView click
$listView.add_MouseClick({
    param($sender, $e)
    
    $item = $listView.GetItemAt($e.X, $e.Y)
    if ($item -ne $null) {
        $itemPath = $item.Tag
        Show-FilePreview -filePath $itemPath
    }
})

# Event handler for ListView double-click
$listView.add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $itemPath = $selectedItem.Tag
        
        if (Test-Path -Path $itemPath -PathType Container) {
            Navigate-To -path $itemPath
            # Keep preview visible when navigating folders
        } else {
            Start-Process $itemPath
        }
    }
})

# Initialize navigation with desktop path
$desktopPath = [Environment]::GetFolderPath("Desktop")
[void]$global:navigationHistory.Add($desktopPath)
$global:currentIndex = 0

# Initial setup
Populate-TreeView
Populate-ListView $desktopPath

# =====================================================================================
# =====================================================================================
# Create a ContextMenuStrip
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Add menu items
$menuItemOpen = $contextMenu.Items.Add("Open")
$menuItemCopy = $contextMenu.Items.Add("Copy")
$menuItemCut = $contextMenu.Items.Add("Cut")
$menuItemPaste = $contextMenu.Items.Add("Paste")
$menuItemDelete = $contextMenu.Items.Add("Delete")
$menuItemProperties = $contextMenu.Items.Add("Properties")
$menuItemRefresh = $contextMenu.Items.Add("Refresh")

# Add Convert menu item
$menuItemConvert = $contextMenu.Items.Add("Convert")
$menuItemConvert.Add_Click({
    if ($listView.SelectedItems.Count -eq 1) {
        $sourcePath = $listView.SelectedItems[0].Tag
        $extension = [System.IO.Path]::GetExtension($sourcePath)
        $formats = Get-SupportedFormats -extension $extension
        
        if ($formats.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No conversion options available for this file type.",
                "Convert File",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }
        
        $selectedFormat = Show-FormatSelectionDialog -filePath $sourcePath
        if ($selectedFormat) {
            $result = Convert-File -sourcePath $sourcePath -targetFormat $selectedFormat.Extension
            if ($result) {
                [System.Windows.Forms.MessageBox]::Show(
                    "File converted successfully!",
                    "Convert File",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Populate-ListView $global:currentPath
            }
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a single file to convert.",
            "Convert File",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})

# Add separator
$contextMenu.Items.Add("-")

# Create a sub-menu for New options
$newSubMenu = New-Object System.Windows.Forms.ToolStripMenuItem("New")
$contextMenu.Items.Add($newSubMenu)

# Add Folder option
$newSubMenu.DropDownItems.Add("Folder").Add_Click({
    Create-NewFolder
})

# Add separator
$newSubMenu.DropDownItems.Add("-")

# Add file format options
$newSubMenu.DropDownItems.Add("Text File").Add_Click({
    Create-NewFile -extension ".txt"
})
$newSubMenu.DropDownItems.Add("Rich Text File").Add_Click({
    Create-NewFile -extension ".rtf"
})
$newSubMenu.DropDownItems.Add("Word Document").Add_Click({
    Create-NewFile -extension ".docx"
})
$newSubMenu.DropDownItems.Add("Excel Workbook").Add_Click({
    Create-NewFile -extension ".xlsx"
})
$newSubMenu.DropDownItems.Add("PowerPoint Presentation").Add_Click({
    Create-NewFile -extension ".pptx"
})
$newSubMenu.DropDownItems.Add("PDF Document").Add_Click({
    Create-NewFile -extension ".pdf"
})
$newSubMenu.DropDownItems.Add("CSV File").Add_Click({
    Create-NewFile -extension ".csv"
})
$newSubMenu.DropDownItems.Add("Markdown File").Add_Click({
    Create-NewFile -extension ".md"
})
$newSubMenu.DropDownItems.Add("XML File").Add_Click({
    Create-NewFile -extension ".xml"
})
$newSubMenu.DropDownItems.Add("JSON File").Add_Click({
    Create-NewFile -extension ".json"
})

# Add these new functions after your existing functions:

# Function to show an input dialog using Windows Forms
function Show-InputDialog {
    param(
        [string]$prompt,
        [string]$title,
        [string]$defaultValue
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(350,150)
    $form.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = $prompt
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,50)
    $textBox.Size = New-Object System.Drawing.Size(310,20)
    $textBox.Text = $defaultValue
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,80)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(170,80)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }
    return $null
}

# Function to create a new folder
function Create-NewFolder {
    $folderName = Show-InputDialog -prompt "Enter the name for the new folder:" -title "New Folder" -defaultValue "New Folder"
    
    if (![string]::IsNullOrWhiteSpace($folderName)) {
        $fullPath = Join-Path -Path $global:currentPath -ChildPath $folderName
        try {
            # Create the new folder
            New-Item -Path $fullPath -ItemType Directory -Force
            [System.Windows.Forms.MessageBox]::Show("Created: $fullPath", "Folder Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Populate-ListView -path $global:currentPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating folder: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Enhanced function to create a new file
function Create-NewFile {
    param (
        [string]$extension
    )

    $fileName = Show-InputDialog -prompt "Enter the name for the new file (without extension):" -title "New File" -defaultValue "New File"
    
    if (![string]::IsNullOrWhiteSpace($fileName)) {
        $fullPath = Join-Path -Path $global:currentPath -ChildPath ($fileName + $extension)
        try {
            # Create the new file
            $null = New-Item -Path $fullPath -ItemType File -Force
            
            # Add default content based on file type
            switch ($extension) {
                ".md" { 
                    Set-Content -Path $fullPath -Value "# New Markdown Document`n`nCreated on $(Get-Date -Format 'yyyy-MM-dd')"
                }
                ".xml" {
                    Set-Content -Path $fullPath -Value "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`n<root>`n</root>"
                }
                ".json" {
                    Set-Content -Path $fullPath -Value "{`n    `"created`": `"$(Get-Date -Format 'yyyy-MM-dd')`"`n}"
                }
                ".html" {
                    Set-Content -Path $fullPath -Value "<!DOCTYPE html>`n<html>`n<head>`n    <title>New Document</title>`n</head>`n<body>`n    <h1>New Document</h1>`n</body>`n</html>"
                }
            }

            [System.Windows.Forms.MessageBox]::Show("Created: $fullPath", "File Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Populate-ListView -path $global:currentPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Event handler for Open
$menuItemOpen.Add_Click({
    foreach ($selectedItem in $listView.SelectedItems) {
        $itemPath = $selectedItem.Tag
        if (Test-Path -Path $itemPath) {
            Start-Process $itemPath
        }
    }
})

# Event handler for Copy
$menuItemCopy.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $global:clipboardPaths = @()
        foreach ($selectedItem in $listView.SelectedItems) {
            $itemPath = $selectedItem.Tag
            if (Test-Path -Path $itemPath) {
                $global:clipboardPaths += $itemPath
            }
        }
        $global:isCut = $false
        [System.Windows.Forms.MessageBox]::Show(
            "Copied $($global:clipboardPaths.Count) items to clipboard",
            "Copy",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# Event handler for Cut
$menuItemCut.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $global:clipboardPaths = @()
        foreach ($selectedItem in $listView.SelectedItems) {
            $itemPath = $selectedItem.Tag
            if (Test-Path -Path $itemPath) {
                $global:clipboardPaths += $itemPath
            }
        }
        $global:isCut = $true
        [System.Windows.Forms.MessageBox]::Show(
            "Cut $($global:clipboardPaths.Count) items to clipboard",
            "Cut",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# Event handler for Paste
$menuItemPaste.Add_Click({
    if ($global:clipboardPaths -and (Test-Path -Path $global:currentPath)) {
        $totalItems = $global:clipboardPaths.Count
        $successCount = 0
        $errorCount = 0
        
        foreach ($sourcePath in $global:clipboardPaths) {
            $destinationPath = Join-Path -Path $global:currentPath -ChildPath (Split-Path $sourcePath -Leaf)
            try {
                if ($global:isCut) {
                    Move-Item -Path $sourcePath -Destination $destinationPath -Force
                } else {
                    Copy-Item -Path $sourcePath -Destination $destinationPath -Force -Recurse
                }
                $successCount++
            } catch {
                $errorCount++
            }
        }
        
        # Clear clipboard after cut operation
        if ($global:isCut -and $successCount -eq $totalItems) {
            $global:clipboardPaths = $null
        }
        
        # Show results
        $message = "Operation completed:`n" +
                  "Successfully processed: $successCount`n" +
                  "Errors: $errorCount"
        [System.Windows.Forms.MessageBox]::Show($message, "Paste Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        Populate-ListView -path $global:currentPath
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "No items to paste or invalid destination.",
            "Paste Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# Event handler for Delete
# MODIFIED: Event handler for Delete - Updated to handle all file types including images
$menuItemDelete.Add_Click({
    # Modify the deletion code to handle locked files
    foreach ($selectedItem in $listView.SelectedItems) {
        $itemPath = $selectedItem.Tag
        try {
            # Force close any open file handles
            if ($pictureBox.Image -ne $null) {
                $pictureBox.Image.Dispose()
                $pictureBox.Image = $null
            }
            
            # Additional step to release file handles
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            # Use more aggressive file deletion with additional options
            Remove-Item -Path $itemPath -Force -ErrorAction Stop
            
            # Remove the item from the ListView
            $listView.Items.Remove($selectedItem)
            
            $successCount++
        } catch {
            try {
                # Try alternative deletion method
                [System.IO.File]::Delete($itemPath)
                $listView.Items.Remove($selectedItem)
                $successCount++
            } catch {
                $errorCount++
                Write-Warning "Could not delete $itemPath : $($_.Exception.Message)"
            }
        }
    }

    # Update the ListView and status bar
    Populate-ListView $global:currentPath
})

# Event handler for Properties
$menuItemProperties.Add_Click({
    if ($listView.SelectedItems.Count -eq 1) {
        # Single item properties
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            $properties = Get-Item -Path $itemPath
            
            # Determine the type
            $type = if ($properties.PSIsContainer) { 'Folder' } else { 'File' }
            
            # Determine the size
            $size = if ([System.IO.File]::Exists($itemPath)) {
                Format-FileSize $properties.Length
            } else {
                'N/A'
            }
            
            # Create the properties info string
            $propertiesInfo = "Name: $($properties.Name)`n" +
                              "Path: $($properties.FullName)`n" +
                              "Type: $type`n" +
                              "Size: $size`n" +
                              "Modified: $($properties.LastWriteTime.ToString('g'))"
            
            [System.Windows.Forms.MessageBox]::Show($propertiesInfo, "Properties", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    else {
        # Multiple items properties
        $totalItems = $listView.SelectedItems.Count
        $totalFiles = 0
        $totalFolders = 0
        $totalSize = 0
        
        foreach ($selectedItem in $listView.SelectedItems) {
            $itemPath = $selectedItem.Tag
            $item = Get-Item -Path $itemPath
            
            if ($item.PSIsContainer) {
                $totalFolders++
            } else {
                $totalFiles++
                $totalSize += $item.Length
            }
        }
        
        $propertiesInfo = "Selected Items: $totalItems`n" +
                         "Files: $totalFiles`n" +
                         "Folders: $totalFolders`n" +
                         "Total Size: $(Format-FileSize $totalSize)"
        
        [System.Windows.Forms.MessageBox]::Show($propertiesInfo, "Multiple Items Properties", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Event handler for Refresh
$menuItemRefresh.Add_Click({
    Populate-ListView -path $global:currentPath
})

# Associate the context menu with the ListView
$listView.ContextMenuStrip = $contextMenu

# Event handler for right-click on ListView
$listView.add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hitTest = $listView.HitTest($_.Location)
        
        # Only clear selection if clicking empty space
        if ($hitTest.Item -eq $null) {
            $listView.SelectedItems.Clear()
        }
        # If clicking an unselected item, select it while preserving other selections
        elseif (!$hitTest.Item.Selected) {
            if (!([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control)) {
                $listView.SelectedItems.Clear()
            }
            $hitTest.Item.Selected = $true
        }
    }
})

# Associate the context menu with the TreeView
$treeView.ContextMenuStrip = $contextMenu

# Event handler for right-click on TreeView
$treeView.add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hitTest = $treeView.HitTest($_.Location)
        if ($hitTest.Node -ne $null) {
            $treeView.SelectedNode = $hitTest.Node
            # Optionally, you can add context menu actions for folders here
        } else {
            # If right-clicked on empty space, clear selection
            $treeView.SelectedNode = $null
        }
    }
})

# Initial TreeView population
Populate-TreeView

# Initial ListView population (using Desktop path from environment variable)
$desktopPath = [Environment]::GetFolderPath("Desktop")
Populate-ListView -path $desktopPath


# =====================================================================================
# =====================================================================================
# Show the form
[void]$form.ShowDialog()