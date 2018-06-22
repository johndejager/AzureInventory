$ExcelFile = "C:\AzureInventory\AzureInventory-$(get-date -f yyyy-MM-dd-hh-mm).xlsx"

#Login-AzureRMAccount
# Requires module - https://github.com/dfinke/ImportExcel
# Install-Module ImportExcel

Import-Module ImportExcel

# Get Subscriptions
$subscriptions = Get-AzureRmSubscription

$resGroupsCol = @()
$vnetsCol = @()
$virtualmachinesCol = @()
$disksCol = @()
$sqlInfoCol = @()

foreach ($subscription in $subscriptions) {
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
    $vnetworks = Get-AzureRmVirtualNetwork
    foreach ($vnetwork in $vnetworks) {
        $subnets = $vnetwork.Subnets
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
    $virtualmachines = Get-AzureRMVM -Status
    foreach ($virtualmachine in $virtualmachines) {
        $vnics = Get-AzureRmNetworkInterface |Where-Object {$_.Id -eq $VirtualMachine.NetworkProfile.NetworkInterfaces.Id} 
        $vmSize = Get-AzureRmVMSize -Location $virtualmachine.Location | Where-Object {$_.Name -eq $virtualmachine.HardwareProfile.VmSize}
        $virtualmachinesObject = [pscustomobject][Ordered]@{
            Subscription     = $subscription.Name        
            Name             = $virtualmachine.Name
            ResourceGroup    = $virtualmachine.ResourceGroupName
            Size             = $virtualmachine.HardwareProfile.VmSize
            NumberOfCores    = $vmSize.NumberOfCores
            MemoryInMB       = $vmSize.MemoryInMB
            MaxDataDiskCount = $vmSize.MaxDataDiskCount
            OSDisk           = $virtualmachine.StorageProfile.OsDisk.Name
            DataDisk         = $virtualmachine.StorageProfile.DataDisks.Name -join "**"
            PowerState       = $virtualmachine.PowerState
            Location         = $virtualmachine.Location
            Vnic             = $vnics.Name
            VnicIP           = $Vnics.IpConfigurations.PrivateIpAddress
        }
        $virtualmachinesCol += $virtualmachinesObject
    }
    $disks = Get-AzureRmDisk 
    foreach ($disk in $disks) {
        $disksObject = [pscustomobject][Ordered]@{
            Subscription   = $subscription.Name        
            Name           = $disk.Name
            ResourceGroup  = $disk.ResourceGroupName
            Size           = $disk.DiskSizeGB
            Sku            = $disk.Sku.Name
            Tier           = $disk.Sku.Tier
            VirtualMachine = ($disk.ManagedBy -split '/')[-1]
        }
        $disksCol += $disksObject
    }
    $sqlServers = Get-AzureRmSqlServer
    foreach ($sqlServer in $sqlServers) {
        $sqlDatabases = Get-AzureRmSqlDatabase -ServerName $sqlServer.ServerName -ResourceGroupName $sqlServer.ResourceGroupName | where-object {$_.DatabaseName -ne "master"}
        foreach ($sqlDatabase in $sqlDatabases) {
            $sqlInfoObject = [pscustomobject][Ordered]@{
                Subscription      = $subscription.Name        
                DatabaseName      = $sqlDatabase.DatabaseName
                DatabaseStatus    = $sqlDatabase.Status
                DatabaseCollation = $sqlDatabase.CollationName
                DatabaseEdition   = $sqlDatabase.Edition
                PricingTier       = $sqlDatabase.CurrentServiceObjectiveName
                DTUs              = $sqlDatabase.Capacity
                ZoneRedundant     = $sqlDatabase.ZoneRedundant
                ServerName        = $sqlServer.ServerName
                ResourceGroup     = $sqlServer.ResourceGroupName
                Location          = $sqlServer.Location
                ServerVersion     = $sqlServer.ServerVersion
                FQDN              = $sqlServer.FullyQualifiedDomainName
                SQLAdmin          = $sqlServer.SqlAdministratorLogin
            }
            $sqlInfoCol += $sqlInfoObject
        }
    }
}
# Export to Excel
$sqlInfoCol | Export-Excel $ExcelFile -WorkSheetname 'SQLDatabase' -AutoSize -AutoFilter
$resGroupsCol | Export-Excel $ExcelFile -WorkSheetname 'ResourceGroups' -AutoSize -AutoFilter
$vnetsCol | Export-Excel $ExcelFile -WorkSheetname 'VNETs' -AutoSize -AutoFilter
$disksCol | Export-Excel $ExcelFile -WorkSheetname 'Disks' -AutoSize -AutoFilter 

