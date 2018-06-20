$Directory = New-Item C:\AzureInventory -type directory -ErrorAction Continue
$Directory = 'C:\AzureInventory'
$xlsx = $Directory + '\AzureInventory.xlsx'

Login-AzureRMAccount
#Functions
Function Release-Ref ($ref)
 {
 ([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
 [System.__ComObject]$ref) -gt 0)
 [System.GC]::Collect()
 [System.GC]::WaitForPendingFinalizers()
 }

Function ConvertCSV-ToExcel {
    [CmdletBinding(
        SupportsShouldProcess = $True,
        ConfirmImpact = ‘low’,
        DefaultParameterSetName = ‘file’
    )]
    Param (
        [Parameter(
            ValueFromPipeline = $True,
            Position = 0,
            Mandatory = $True,
            HelpMessage = ”Name of CSV/s to import”)]
        [ValidateNotNullOrEmpty()]
        [array]$csvFiles,
        [Parameter(
            ValueFromPipeline = $False,
            Position = 1,
            Mandatory = $True,
            HelpMessage = ”Name of excel file output”)]
        [ValidateNotNullOrEmpty()]
        [string]$output
    )
    Begin {
        #Configure regular expression to match full path of each file
        [regex]$regex = “^\w\:\\”
        #Find the number of CSVs being imported
        $count = ($csvFiles.count - 1)
        #Create Excel Com Object
        $excel = new-object -com excel.application
        #Disable alerts 
        $excel.DisplayAlerts = $False
        #Show Excel application
        $excel.Visible = $False
        #Add workbook
        $workbook = $excel.workbooks.Add()
        #Remove other worksheets
        #$workbook.worksheets.Item(2).delete()
        #After the first worksheet is removed,the next one takes its place
        #$workbook.worksheets.Item(2).delete()
        #Define initial worksheet number
        $i = 1
    }
    Process {
        ForEach ($csvfile in $csvFiles) {
            #$csvfile = $csvfiles[0]
            #$csvfile = $csvfiles[1]
            #If more than one file, create another worksheet for each file
            If ($i -gt 1) {
                $workbook.worksheets.Add() | Out-Null
            }
            # Use the first worksheet in the workbook (also the newest created worksheet is always 1)
            # Add name of CSV as worksheet name
            $delimiter = "," #Specify the delimiter used in the file
            # Create a new Excel workbook with one empty sheet
            $excel = New-Object -ComObject excel.application 
            #$workbook = $excel.Workbooks.Add(1)
            $worksheet = $workbook.worksheets.Item(1)
            $worksheet.Name = “$((Get-ChildItem $csvfile).basename)”
            # Build the QueryTables.Add command and reformat the data
            $TxtConnector = ("TEXT;" + $csvfile)
            $Connector = $worksheet.QueryTables.add($TxtConnector, $worksheet.Range("A1"))
            $query = $worksheet.QueryTables.item($Connector.name)
            $query.TextFileOtherDelimiter = $delimiter
            $query.TextFileParseType = 1
            $query.TextFileColumnDataTypes = , 1 * $worksheet.Cells.Columns.Count
            $query.AdjustColumnWidth = 1
            # Execute & delete the import query
            $query.Refresh()
            $query.Delete()
            $i++
        }
    }

    End {
        #Save spreadsheet
        $workbook.saveas($xlsx)
        Write-Host -Fore Green “File saved to $xlsx”
        #Close Excel
        $excel.quit()
        $a = Release-Ref($range)
    }
}
# Get Subscriptions
$subscriptions = Get-AzureRmSubscription 
#Resource Groups
$resGroupsCol = @()
foreach ($subscription in $subscriptions) {
    Write-Host 'Now processing '$subscription.Name''
    Select-AzureRmSubscription -Subscription $subscription.Name
    $resGroups = Get-AzureRmResourceGroup | Select-Object ResourceGroupName, Location 
    foreach ($resGroup in $resGroups) {
        $resGroupsObject = [pscustomobject][Ordered]@{
            Subscription          = $subscription.Name
            ResourceGroupName     = $resGroup.ResourceGroupName
            ResourceGroupLocation = $resGroup.Location
        }
        $resGroupsCol += $resGroupsObject
    }
}
$resGroupsPath = $Directory + "\ResourceGroups.csv"
$resGroupsCol | Export-Csv $resGroupsPath -NoTypeInformation
# Virtual Networks
$vnetsCol = @()
foreach ($subscription in $subscriptions) {
    Write-Host 'Now processing '$subscription.Name''
    Select-AzureRmSubscription -Subscription $subscription.Name
    $vnetworks = Get-AzureRmVirtualNetwork
    foreach ($vnetwork in $vnetworks) {
        Write-Host 'Now processing VNET:'$vnetwork.Name''
        $subnets = $vnetwork.Subnets
        #$subnet = $subnets[0]
        foreach ($subnet in $subnets) {
            $NetworkSecurityGroup = Get-AzureRmNetworkSecurityGroup | Where-Object {$_.Id -eq $subnet.NetworkSecurityGroup.Id}
            $RouteTable = Get-AzureRmRouteTable | Where-Object {$_.Id -eq $subnet.RouteTable.Id}
            $vnetworksObject = [pscustomobject][Ordered]@{
                Subscription               = $subscription.Name
                VNETName                   = $vnetwork.Name
                VNETResourceGroup          = $vnetwork.ResourceGroupName
                VNETAddressPrefixes        = $vnetwork.AddressSpace.AddressPrefixes -join "**"
                VNETDNS                    = $vnetwork.DhcpOptions.DnsServers -join "**"
                SubnetName                 = $subnet.Name
                SubnetAddressprefix        = $subnet.AddressPrefix
                SubnetNetWorkSecuritygroup = $NetworkSecurityGroup.Name
                SubnetRouteTable           = $RouteTable.Name
            }
            $vnetsCol += $vnetworksObject
        }
      
    }
    
}
$vnetsPath = $Directory + "\VNETs.csv"
$vnetsCol | Export-Csv $vnetsPath -NoTypeInformation 
# Virtual Machines
$virtualmachinesCol = @()
foreach ($subscription in $subscriptions) {
    Write-Host 'Now processing '$subscription.Name''
    Select-AzureRmSubscription -Subscription $subscription.Name
    $virtualmachines = get-azurermvm -Status
    foreach ($virtualmachine in $virtualmachines) {
        $vnics = Get-AzureRmNetworkInterface |Where-Object {$_.Id -eq $VirtualMachine.NetworkProfile.NetworkInterfaces.Id} 
        $virtualmachinesObject = [pscustomobject][Ordered]@{
            Subscription  = $subscription.Name        
            Name          = $virtualmachine.Name
            ResourceGroup = $virtualmachine.ResourceGroupName
            Size          = $virtualmachine.HardwareProfile.VmSize
            OSDisk        = $virtualmachine.StorageProfile.OsDisk.Name
            DataDisk      = $virtualmachine.StorageProfile.DataDisks.Name -join "**"
            PowerState    = $virtualmachine.PowerState
            Vnic          = $vnics.Name
            VnicIP        = $Vnics.IpConfigurations.PrivateIpAddress
        }
        $virtualmachinesCol += $virtualmachinesObject
    }
}
$virtualmachinesPath = $Directory + "\VMs.csv"
$virtualmachinesCol | Export-Csv $virtualmachinesPath -NoTypeInformation 


ConvertCSV-ToExcel -CSVfiles @($resGroupsPath, $vnetsPath, $virtualmachinesPath) -output $xlsx 


