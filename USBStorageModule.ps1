# USBStorageModule.psm1

# Function to get drive information in detail
function Get-DriveDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    
    try {
        $volume = Get-Volume -DriveLetter $DriveLetter.TrimEnd(':')
        $disk = Get-Disk | Where-Object Number -eq $volume.DiskNumber
        
        return @{
            DriveLetter = $DriveLetter
            VolumeName = $volume.FileSystemLabel
            FileSystem = $volume.FileSystem
            Size = [math]::Round($volume.Size / 1GB, 2)
            FreeSpace = [math]::Round($volume.SizeRemaining / 1GB, 2)
            HealthStatus = $volume.HealthStatus
            BusType = $disk.BusType
            MediaType = $disk.MediaType
        }
    } catch {
        Write-Warning "Error getting details for drive $DriveLetter : $_"
        return $null
    }
}

# Function to detect USB storage devices
function Get-USBStorageDevices {
    try {
        $usbDrives = Get-WmiObject Win32_DiskDrive | 
            Where-Object { $_.InterfaceType -eq "USB" } | 
            ForEach-Object {
                $drive = $_
                Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($drive.DeviceID)'} WHERE AssocClass = Win32_DiskDriveToDiskPartition" | 
                ForEach-Object {
                    Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($_.DeviceID)'} WHERE AssocClass = Win32_LogicalDiskToPartition" | 
                    ForEach-Object {
                        Get-DriveDetails -DriveLetter "$($_.DeviceID):"
                    }
                }
            }
        return $usbDrives
    } catch {
        Write-Warning "Error detecting USB devices: $_"
        return $null
    }
}

# Function to detect connected phones
function Get-ConnectedPhones {
    try {
        $phones = Get-WmiObject Win32_USBHub | 
            Where-Object { $_.Name -match "Phone|Android|iPhone|Mobile" }
        return $phones
    } catch {
        Write-Warning "Error detecting phones: $_"
        return $null
    }
}

# Function to get file listing from a drive
function Get-DriveContents {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [int]$Depth = 1
    )
    
    try {
        $items = Get-ChildItem -Path $Path -Depth $Depth -ErrorAction SilentlyContinue | 
            Select-Object Name, Length, LastWriteTime, @{
                Name='Type'
                Expression={
                    if ($_.PSIsContainer) { 'Folder' } else { 'File' }
                }
            }, @{
                Name='SizeGB'
                Expression={
                    if (!$_.PSIsContainer) {
                        [math]::Round($_.Length / 1GB, 4)
                    } else { $null }
                }
            }
        return $items
    } catch {
        Write-Warning "Error accessing path $Path : $_"
        return $null
    }
}

# Function to monitor for new USB devices
function Start-USBMonitor {
    param (
        [scriptblock]$OnDeviceConnected,
        [scriptblock]$OnDeviceRemoved
    )
    
    $knownDevices = @{}
    
    while ($true) {
        $currentDevices = Get-USBStorageDevices
        
        # Check for new devices
        foreach ($device in $currentDevices) {
            if ($device -and !$knownDevices.ContainsKey($device.DriveLetter)) {
                $knownDevices[$device.DriveLetter] = $device
                if ($OnDeviceConnected) {
                    & $OnDeviceConnected $device
                }
            }
        }
        
        # Check for removed devices
        $removedDevices = @()
        foreach ($knownDevice in $knownDevices.Keys) {
            if (!($currentDevices | Where-Object { $_.DriveLetter -eq $knownDevice })) {
                $removedDevices += $knownDevice
                if ($OnDeviceRemoved) {
                    & $OnDeviceRemoved $knownDevices[$knownDevice]
                }
            }
        }
        
        # Remove disconnected devices from known devices
        foreach ($removedDevice in $removedDevices) {
            $knownDevices.Remove($removedDevice)
        }
        
        Start-Sleep -Seconds 2
    }
}

# Function to integrate with existing file explorer
function Add-USBToExplorer {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.TreeView]$ExplorerTreeView,
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ListView]$FileListView
    )
    
    # Create USB Devices root node
    $usbRoot = New-Object System.Windows.Controls.TreeViewItem
    $usbRoot.Header = "USB Devices"
    $ExplorerTreeView.Items.Add($usbRoot)
    
    # Start USB monitoring
    $monitorJob = Start-Job -ScriptBlock {
        param($OnConnected, $OnRemoved)
        Start-USBMonitor -OnDeviceConnected $OnConnected -OnDeviceRemoved $OnRemoved
    } -ArgumentList {
        param($device)
        
        # Add new device to tree
        $deviceNode = New-Object System.Windows.Controls.TreeViewItem
        $deviceNode.Header = "$($device.VolumeName) ($($device.DriveLetter))"
        $deviceNode.Tag = $device
        $usbRoot.Items.Add($deviceNode)
        
        # Populate initial content
        $contents = Get-DriveContents -Path $device.DriveLetter
        foreach ($item in $contents) {
            $itemNode = New-Object System.Windows.Controls.TreeViewItem
            $itemNode.Header = $item.Name
            $itemNode.Tag = $item
            $deviceNode.Items.Add($itemNode)
        }
    }, {
        param($device)
        
        # Remove device from tree
        $deviceNode = $usbRoot.Items | Where-Object { $_.Tag.DriveLetter -eq $device.DriveLetter }
        if ($deviceNode) {
            $usbRoot.Items.Remove($deviceNode)
        }
    }
    
    # Handle tree view selection
    $ExplorerTreeView.Add_SelectedItemChanged({
        $selected = $ExplorerTreeView.SelectedItem
        if ($selected -and $selected.Tag) {
            $path = if ($selected.Tag.GetType().Name -eq 'Hashtable') {
                $selected.Tag.DriveLetter
            } else {
                Join-Path $selected.Parent.Tag.DriveLetter $selected.Tag.Name
            }
            
            $FileListView.Items.Clear()
            $contents = Get-DriveContents -Path $path
            foreach ($item in $contents) {
                $FileListView.Items.Add($item)
            }
        }
    })
    
    return $monitorJob
}

# Export functions
Export-ModuleMember -Function Get-DriveDetails, Get-USBStorageDevices, Get-ConnectedPhones, 
    Get-DriveContents, Start-USBMonitor, Add-USBToExplorer