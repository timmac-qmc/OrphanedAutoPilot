if(-not (Get-Module Microsoft.Graph -ListAvailable)){
    Install-Module Microsoft.Graph -AllowClobber -Scope CurrentUser -Force -SkipPublisherCheck
}

Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.DeviceManagement.Enrollment

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-AutoPilotDevices
{
    param(
        [Parameter(ParameterSetName="ShowAll")] [bool] $ShowAll = $false
    )
    if($ShowAll -eq $true)
    {    
        return Get-MgDeviceManagementWindowsAutopilotDeviceIdentity
    }
    else {
        return Get-MgDeviceManagementWindowsAutopilotDeviceIdentity | Where-Object { $_.ManagedDeviceId -eq "00000000-0000-0000-0000-000000000000"  } 
    }
}

function Remove-AutoPilotDevice
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ParameterSetName="id")] [guid] $AutoPilotId
    )
    Process {
        return Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $AutoPilotId -PassThru
    }
}

function ConvertTo-DataTable
{
    <#
    .Synopsis
        Creates a DataTable from an object
    .Description
        Creates a DataTable from an object, containing all properties (except built-in properties from a database)
    .Example
        Get-ChildItem| Select Name, LastWriteTime | ConvertTo-DataTable
    .Link
        Select-DataTable
    .Link
        Import-DataTable
    .Link
        Export-Datatable
    #> 
    [OutputType([Data.DataTable])]
    param(
    # The input objects
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)]
    [PSObject[]]
    $InputObject
    ) 
 
    begin { 
        
        $outputDataTable = new-object Data.datatable   
          
        $knownColumns = @{}
        
        
    } 

    process {         
               
        foreach ($In in $InputObject) { 
            $DataRow = $outputDataTable.NewRow()   
            $isDataRow = $in.psobject.TypeNames -like "*.DataRow*" -as [bool]

            $simpleTypes = ('System.Boolean', 'System.Byte[]', 'System.Byte', 'System.Char', 'System.Datetime', 'System.Decimal', 'System.Double', 'System.Guid', 'System.Int16', 'System.Int32', 'System.Int64', 'System.Single', 'System.UInt16', 'System.UInt32', 'System.UInt64')

            $SimpletypeLookup = @{}
            foreach ($s in $simpleTypes) {
                $SimpletypeLookup[$s] = $s
            }            
            
            
            foreach($property in $In.PsObject.properties) {   
                if ($isDataRow -and 
                    'RowError', 'RowState', 'Table', 'ItemArray', 'HasErrors' -contains $property.Name) {
                    continue     
                }
                $propName = $property.Name
                $propValue = $property.Value
                $IsSimpleType = $SimpletypeLookup.ContainsKey($property.TypeNameOfValue)

                if (-not $outputDataTable.Columns.Contains($propName)) {   
                    $outputDataTable.Columns.Add((
                        New-Object Data.DataColumn -Property @{
                            ColumnName = $propName
                            DataType = if ($issimpleType) {
                                $property.TypeNameOfValue
                            } else {
                                'System.Object'
                            }
                        }
                    ))
                }                   
                
                $DataRow.Item($propName) = if ($isSimpleType -and $propValue) {
                    $propValue
                } elseif ($propValue) {
                    [PSObject]$propValue
                } else {
                    [DBNull]::Value
                }
                
            }   
            $outputDataTable.Rows.Add($DataRow)   
        } 
        
    }  
      
    end 
    { 
        ,$outputDataTable

    } 
 
}

function Show-Devices {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Data
    )

    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Ophaned AutoPilot Devices"
    $form.Size = New-Object System.Drawing.Size(800, 450)
    $form.StartPosition = "CenterScreen"
    #$form.AutoSize = $true;
    #$form.AutoSizeMode = 'GrowAndShrink'

    # Create the filter TextBox
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 10)
    $textBox.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textBox)

    # Create the filter Button
    $filterButton = New-Object System.Windows.Forms.Button
    $filterButton.Location = New-Object System.Drawing.Point(310, 10)
    $filterButton.Size = New-Object System.Drawing.Size(75, 20)
    $filterButton.Text = "Filter"
    $form.Controls.Add($filterButton)

    # Create the refresh Button
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(390, 10)
    $refreshButton.Size = New-Object System.Drawing.Size(75, 20)
    $refreshButton.Text = "Refresh"
    $form.Controls.Add($refreshButton)

    # Create the azure Button
    $azureButton = New-Object System.Windows.Forms.Button
    $azureButton.Location = New-Object System.Drawing.Point(470, 10)
    $azureButton.Size = New-Object System.Drawing.Size(100, 20)
    $azureButton.Text = "Show in Azure"
    $azureButton.Enabled = $false
    $form.Controls.Add($azureButton)

    # Create the delete Button
    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Location = New-Object System.Drawing.Point(575, 10)
    $deleteButton.Size = New-Object System.Drawing.Size(100, 20)
    $deleteButton.Text = "Delete Device"
    $deleteButton.Enabled = $false
    $form.Controls.Add($deleteButton)

    # Create checkbox
    $showAllCb = New-Object System.Windows.Forms.Checkbox
    $showAllCb.Location = New-Object System.Drawing.Point(685, 10)
    $showAllCb.Size = New-Object System.Drawing.Size(100, 20)
    $showAllCb.Text = "Show All"
    $form.Controls.Add($showAllCb)

    # Create the DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(10, 40)
    $dataGridView.Size = New-Object System.Drawing.Size(760, 360)
    $dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                           [System.Windows.Forms.AnchorStyles]::Left -bor `
                           [System.Windows.Forms.AnchorStyles]::Right
    $dataGridView.ScrollBars = 'Both'
    $dataGridView.AutoSizeColumnsMode = "AllCells"
    $dataGridView.ReadOnly = $true
    $dataGridView.SelectionMode = 'FullRowSelect'
    $dataGridView.DataBindings.DefaultDataSourceUpdateMode = 0 
    $form.Controls.Add($dataGridView)

    $dataGridView.DataSource = $data | ConvertTo-Datatable
    $dataGridView.Refresh()

    # Handle the Resize event to adjust DataGridView size
    $form.Add_Shown({$form.Activate()})
    $form.add_ResizeEnd({
        #$dataGridView.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 20), ($form.ClientSize.Height - 50))
    })

    $filterButton.Add_Click({
        $filter = $textBox.Text
        if($filter -ne "")
        {
            $filteredData = $Data | Where-Object {
                $match = $false
                foreach ($column in $columns) {
                    if ($_.PSObject.Properties[$column].Value -like "*$filter*") {
                        $match = $true
                    }
                }
                $match
            }
        }
        else {
            $filteredData = $Data
        }
        $dataGridView.DataSource = $filteredData | ConvertTo-Datatable
        $dataGridView.Refresh()
    })

    $refreshButton.Add_Click({
        $dataGridView.DataSource = Get-AutoPilotDevices -ShowAll $showAllCb.Checked | Select-Object | ConvertTo-Datatable
        $dataGridView.Refresh()
    })

    $azureButton.Add_Click({
        try {
            $device = $Data[$dataGridView.CurrentCell.RowIndex]
            $azureDevice = Get-MgDevice -Search "deviceId:$($device.AzureActiveDirectoryDeviceId)" -ConsistencyLevel eventual
            if($azureDevice -ne $null)
            {
                Start-Process "https://portal.azure.com/#view/Microsoft_AAD_Devices/DeviceDetailsMenuBlade/~/Properties/objectId/$($azureDevice.Id)"
            }
        } catch {
            [System.Windows.MessageBox]::Show("An error occured fetching Azure device info" + $_,"Azure Device","Ok","Error")
        }
    })

    $deleteButton.Add_Click({
        $deleteButton.Enabled = $false

        try {
            $device = $Data[$dataGridView.CurrentCell.RowIndex]
            $confirmDialog =  [System.Windows.MessageBox]::Show("Confirm deletion of $($device.Id) ($($device.SerialNumber))","Delete Device","YesNo","Error")

            if($confirmDialog -eq 'Yes')
            {
                if((Remove-AutoPilotDevice -AutoPilotId $device.Id) -eq $true)
                {
                    [System.Windows.MessageBox]::Show("Delete request submitted, it may take up to 30 minutes to process...","Device Deleted","Ok","Information")
                    $dataGridView.DataSource = Get-AutoPilotDevices -ShowAll $showAllCb.Checked | Select-Object | ConvertTo-Datatable
                    $dataGridView.Refresh()
                }
                else {
                    [System.Windows.MessageBox]::Show("Failed to delete device, does it still exist?","Delete Failed","Ok","Error")
                }
            }
        } catch {
            [System.Windows.MessageBox]::Show("An error occured deleting the device: " + $_,"Delete Device","Ok","Error")
        }
        finally {
            $deleteButton.Enabled = $true
        }
    })


    $showAllCb.Add_CheckStateChanged({ 
        $dataGridView.DataSource = Get-AutoPilotDevices -ShowAll $showAllCb.Checked | Select-Object | ConvertTo-Datatable
        $dataGridView.Refresh()
    })

    $dataGridView.add_selectionChanged({
        if($dataGridView.CurrentCell.RowIndex -ge 0)
        {
            $azureButton.Enabled = $true
            $deleteButton.Enabled = $true
        }
        else {
            $azureButton.Enabled = $false
            $deleteButton.Enabled = $false
        }
    })

    # Show the form
    [void] $form.ShowDialog()
}

function Get-MissingAutoPilotDevices {
    Connect-MgGraph -NoWelcome
    Show-Devices -Data (Get-AutoPilotDevices | Select-Object)
}
