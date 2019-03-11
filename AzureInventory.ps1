#----------------------------------------------------------------------------------
#---------------------LOGIN TO AZURE AND SELECT THE SUBSCRIPTION-------------------
#----------------------------------------------------------------------------------
param (
    [string] $ConnectionName = 'AzureRunAsConnection'
)
try {
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName
    "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint
    Connect-AzureAD `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint
}
catch {
    if (!$servicePrincipalConnection) {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    }
    else {
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

$excelFile = "$env:TEMP\AzureInventory-" + "$(get-date -f yyyy-MM-dd-hh-mm)" + ".xlsx"

# Get subscriptions
$subscriptions = Get-AzureRmSubscription

# Define arrays
$resGroupsCol = @()
$vnetsCol = @()
$virtualmachinesCol = @()
$disksCol = @()
$sqlInfoCol = @()
$nsgInfoCol = @()
$roleAssignmentInfoCol = @()
$groupMemberInfoCol = @()
$storageAccountInfoCol = @()
$policyAssignmentsInfoCol = @()
$scalesetCol = @()
$vmsnapshotsCol = @()
$azureloadbalancersCol = @()
$gatewaysCol = @()
$keyvaultsCol = @()
$recoveryvaultsCol = @()
$backupjobscol = @()
$publicIpsCol = @()
$peeringsCol = @()
$quotaarr = @()

foreach ($subscription in $subscriptions) {
    Select-AzureRmSubscription -Subscription $subscription.Name
    $resGroups = Get-AzureRmResourceGroup
    foreach ($resGroup in $resGroups) {
        $resGroupsObject = [pscustomobject][Ordered]@{
            Subscription          = $subscription.Name
            ResourceGroupName     = $resGroup.ResourceGroupName
            ResourceGroupLocation = $resGroup.Location
            CreatedBy             = $resGroup.Tags.CreatedBy
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
    $peeringnetworks = $vnetworks | Where-Object {$_.VirtualNetworkPeerings -ne $null}
    foreach ($peeringnetwork in $peeringnetworks) {
        $peering = Get-AzureRmVirtualNetworkPeering -VirtualNetworkName $peeringnetwork.Name -ResourceGroupName $peeringnetwork.ResourceGroupName
        $peeringsObject = [pscustomobject][Ordered]@{
            Subscription                    = $subscription.Name
            Name                            = $peering.Name
            ResourceGroup                   = $peering.VirtualNetworkName
            Status                          = $peering.PeeringState
            RemoteVirtualNetwork            = ($peering.RemoteVirtualNetwork.Id -split '/')[-1]
            AllowVirtualNetworkAccess       = $peering.AllowVirtualNetworkAccess
            AllowForwardedTraffic           = $peering.AllowForwardedTraffic
            AllowGatewayTransit             = $peering.AllowGatewayTransit
            UseRemoteGateways               = $peering.UseRemoteGateways
            RemoteGateways                  = $peering.RemoteGateways
            RemoteVirtualNetwokAddressSpace = $peering.RemoteVirtualNetworkAddressSpace
        }
        $peeringsCol += $peeringsObject
    }
    $publicIps = Get-AzureRmPublicIpAddress
    foreach ($publicIp in $publicIps) {
        $vnics = Get-AzureRmNetworkInterface | Where-Object {$_.IpConfigurations.PublicIPAddress.Id -eq $publicIp.Id }
        $publicIpsObject = [pscustomobject][Ordered]@{
            Subscription             = $subscription.Name
            Name                     = $publicIp.Name
            ResourceGroup            = $publicIp.ResourceGroupName
            Location                 = $publicIp.Location
            SKU                      = $publicIp.Sku.Name
            IpAddress                = $publicIp.IpAddress
            PublicIPAllocationMethod = $publicIp.PublicIPAllocationMethod
            AccociatedTo             = ($vnics.VirtualMachine.Id -split '/')[-1]
        }
        $publicIpsCol += $publicIpsObject
    }

    $virtualmachines = Get-AzureRMVM -Status    
    foreach ($virtualmachine in $virtualmachines) {
        $vnics = Get-AzureRmNetworkInterface | Where-Object {$_.Id -eq $virtualMachine.NetworkProfile.NetworkInterfaces.Id}
        $vmSize = Get-AzureRmVMSize -Location $virtualmachine.Location | Where-Object {$_.Name -eq $virtualmachine.HardwareProfile.VmSize}
        $AVset = Get-AzureRmAvailabilitySet -ResourceGroupName $virtualmachine.ResourceGroupName
        $pip = Get-AzureRmPublicIpAddress | Where-Object {$_.Id -eq $vnics.IpConfigurations.PublicIPAddress.Id }
        $virtualmachinesObject = [pscustomobject][Ordered]@{
            Subscription      = $subscription.Name
            Name              = $virtualmachine.Name
            ResourceGroup     = $virtualmachine.ResourceGroupName
            Size              = $virtualmachine.HardwareProfile.VmSize
            AvailabilitySet   = $AVset.Name
            NumberOfCores     = $vmSize.NumberOfCores
            MemoryInMB        = $vmSize.MemoryInMB
            MaxDataDiskCount  = $vmSize.MaxDataDiskCount
            OperatingSystem   = $virtualmachine.StorageProfile.OsDisk.OsType
            OSDisk            = $virtualmachine.StorageProfile.OsDisk.Name
            DataDisk          = $virtualmachine.StorageProfile.DataDisks.Name -join "**"
            PowerState        = $virtualmachine.PowerState
            Location          = $virtualmachine.Location
            Extensions        = ($virtualmachine.Extensions.Id -split '/')[-1]
            Vnic              = $vnics.Name
            PublicIP          = $pip.IpAddress
            VnicIP            = $Vnics.IpConfigurations.PrivateIpAddress
            DnsServers        = $Vnics.DnsSettings.DnsServers
            AppliedDnsServers = $Vnics.DnsSettings.AppliedDnsServers
        }
        $virtualmachinesCol += $virtualmachinesObject
    }
    $scaleSets = Get-AzureRmVmss -InstanceView
    foreach ($scaleSet in $scaleSets) {
        $scalesetVms = Get-AzureRmVmssVM -ResourceGroupName $scaleSet.ResourceGroupName -VMScaleSetName $scaleSet.Name
        foreach ($scalesetVm in $scalesetVms) {
            $scalesetObject = [pscustomobject][Ordered]@{
                Subscription  = $subscription.Name
                VMName        = $scalesetVm.Name
                VMSku         = $scalesetVm.Sku.Name
                ObjectId      = $scalesetVm.InstanceId
                ScaleSetName  = $scaleSet.Name
                ResourceGroup = $scaleSet.ResourceGroupName
                Location      = $scaleSet.Location
                SKU           = $scaleSet.Sku.Name
            }
            $scalesetCol += $scalesetObject
        }
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
            $tdeStatus = Get-AzureRmSqlDatabaseTransparentDataEncryption -ServerName $sqlServer.ServerName -DatabaseName $sqlDatabase.DatabaseName -ResourceGroupName $sqlDatabase.ResourceGroupName
            $sqlInfoObject = [pscustomobject][Ordered]@{
                Subscription      = $subscription.Name
                DatabaseName      = $sqlDatabase.DatabaseName
                DatabaseStatus    = $sqlDatabase.Status
                DatabaseCollation = $sqlDatabase.CollationName
                DatabaseEdition   = $sqlDatabase.Edition
                PricingTier       = $sqlDatabase.CurrentServiceObjectiveName
                DTUs              = $sqlDatabase.Capacity
                ZoneRedundant     = $sqlDatabase.ZoneRedundant
                TDEStatus         = $tdeStatus.State
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
    $nsgs = Get-AzureRmNetworkSecurityGroup
    foreach ($nsg in $nsgs) {
        $securityRules = $nsg.SecurityRules
        foreach ($securityRule in $securityRules) {
            $nsgInfoObject = [pscustomobject][Ordered]@{
                Subscription                         = $subscription.Name
                NSGName                              = $nsg.Name
                NSGResourceGroupName                 = $nsg.ResourceGroupName
                NSGLocation                          = $nsg.Location
                RuleName                             = $securityRule.Name
                RuleDescription                      = $securityRule.Description
                Protocol                             = $securityRule.Protocol
                SourcePortRange                      = $securityRule.SourcePortRange
                DestionationPortRange                = $securityRule.DestinationPortRange
                SourceAddressPrefix                  = $securityRule.SourceAddressPrefix
                DestionationAddressPrefix            = $securityRule.DestinationAddressPrefix
                Access                               = $securityRule.Access
                Priority                             = $securityRule.Priority
                Direction                            = $securityRule.Direction
                SourceApplicationSecurityGroups      = $securityRule.SoureApplicationSecurityGroups
                DestinationApplicationSecurityGroups = $securityRule.DestionationApplicationSecurityGroups
            }
            $nsgInfoCol += $nsgInfoObject
        }
    }
    $roleDefinitions = Get-AzureRmRoleDefinition | Where-Object {$_.Name -eq 'Owner' -or $_.Name -eq 'Contributor'}
    foreach ($roleDefinition in $roleDefinitions) {
        $roleAssignments = Get-AzureRmRoleAssignment
        $roleInfo = Get-AzureADUser | Where-Object {$_.DisplayName -eq $roleAssignment.DisplayName}
        foreach ($roleAssignment in $roleAssignments) {
            $roleInfo = Get-AzureADUser | Where-Object {$_.DisplayName -eq $roleAssignment.DisplayName}
            $roleAssignmentInfoObject = [pscustomobject][Ordered]@{
                Subscription = $subscription.Name
                Name         = $roleDefinition.Name
                Custom       = $roleDefinition.IsCustom
                DisplayName  = $roleAssignment.DisplayName
                Email        = $roleInfo.Mail
                ObjectType   = $roleAssignment.ObjectType
                Scope        = $roleAssignment.Scope
            }
            $roleAssignmentInfoCol += $roleAssignmentInfoObject
        }
    }
    $storageAccounts = Get-AzureRmStorageAccount
    foreach ($storageAccount in $storageAccounts) {
        $storageAccountInfoObject = [pscustomobject][Ordered]@{
            Subscription           = $subscription.Name
            Name                   = $storageAccount.StorageAccountName
            ResourceGroup          = $storageAccount.ResourceGroupName
            Location               = $storageAccount.Location
            SKU                    = $storageAccount.Sku.Tier
            Encryption             = $storageAccount.Encryption.KeySource
            CustomDomain           = $storageAccount.CustomDomain
            CreationTime           = $storageAccount.CreationTime
            EnableHttpsTrafficOnly = $storageAccount.EnableHttpsTrafficOnly
        }
        $storageAccountInfoCol += $storageAccountInfoObject
    }
    $policyAssignments = Get-AzureRmPolicyAssignment  
    foreach ($policyAssignment in $policyAssignments) {
        $policyAssignmentsInfoObject = [pscustomobject][Ordered]@{
            Subscription       = $subscription.Name
            Name               = $policyAssignment.Properties.displayName
            Description        = $policyAssignment.Properties.Description
            ResourceGroup      = $policyAssignment.ResourceGroupName
            PolicyDefinitionId = ($policyAssignment.Properties.policyDefinitionId -split '/')[-1]
            Scope              = $policyAssignment.Properties.scope
            Exclusions         = $policyAssignment.Properties.notScopes
        }
        $policyAssignmentsInfoCol += $policyAssignmentsInfoObject
    }
    $vmsnapshots = Get-AzureRMsnapshot
    foreach ($vmsnapshot in $vmsnapshots) {
        $vmsnapshotsObject = [pscustomobject][Ordered]@{
            Subscription = $subscription.Name
            Name         = $vmsnapshot.Name
            DiskSizeGB   = $vmsnapshot.DiskSizeGB
            TimeCreated  = $vmsnapshot.TimeCreated
            Location     = $vmsnapshot.Location
        }
        $vmsnapshotsCol += $vmsnapshotsObject
    }
    $azureloadbalancers = Get-AzureRmLoadBalancer
    foreach ($azureloadbalancer in $azureloadbalancers) {
        $azureloadbalancersObject = [pscustomobject][Ordered]@{
            Subscription  = $subscription.Name
            Name          = $azureloadbalancer.Name
            ResourceGroup = $azureloadbalancer.ResourceGroupName
            Location      = $azureloadbalancer.Location
            Sku           = $azureloadbalancer.sku.Name
        }
        $azureloadbalancersCol += $azureloadbalancersObject
    }
    $vnetworks = Get-AzureRmVirtualNetwork
    foreach ($vnetwork in $vnetworks) {
        $subnets = $vnetwork.Subnets
        foreach ($subnet in $subnets | Where-Object {$_.Name -eq 'GatewaySubnet'}) {
            $gateways = Get-AzureRmVirtualNetworkGateway -ResourceGroupName $vnetwork.ResourceGroupName
            foreach ($gateway in $gateways) {
                $gatewaysObject = [pscustomobject][Ordered]@{
                    Subscription  = $subscription.Name
                    Name          = $gateway.Name
                    ResourceGroup = $gateway.ResourceGroupName
                    Location      = $gateway.Location
                    Sku           = $gateway.sku.Name
                    Capacity      = $gateway.sku.Capacity
                    GatewayType   = $gateway.GatewayType
                    EnableBgp     = $gateway.EnableBgp
                    ActiveActive  = $gateway.ActiveActive
                }
                $gatewaysCol += $gatewaysObject
            }
        }
    }
    $keyvaults = Get-AzureRmKeyVault
    foreach ($keyvault in $keyvaults) {
        $keyvaultsObject = [pscustomobject][Ordered]@{
            Subscription  = $subscription.Name
            VaultName     = $keyvault.VaultName
            ResourceGroup = $keyvault.ResourceGroupName
            Location      = $keyvault.Location
        }
        $keyvaultsCol += $keyvaultsObject
    }
    $recoveryvaults = Get-AzureRmRecoveryServicesVault
    foreach ($recoveryvault in $recoveryvaults) {
        $recoveryvaultsObject = [pscustomobject][Ordered]@{
            Subscription  = $subscription.Name
            Name          = $recoveryvault.Name
            ResourceGroup = $recoveryvault.ResourceGroupName
            Location      = $recoveryvault.Location
        }
        $recoveryvaultsCol += $recoveryvaultsObject
    }
    $recoveryvaults = Get-AzureRmRecoveryServicesVault
    foreach ($recoveryvault in $recoveryvaults) {
        Set-AzureRmRecoveryServicesVaultContext -Vault $recoveryvault
        $jobs = Get-AzureRmRecoveryServicesBackupJob -From (Get-Date).AddDays(-1).ToUniversalTime()
        foreach ($job in $jobs) {
            $backupjobsObject = [pscustomobject][Ordered]@{
                Subscription  = $subscription.Name
                Name          = $recoveryvault.Name
                ResourceGroup = $recoveryvault.ResourceGroupName
                Location      = $recoveryvault.Location
                WorkloadName  = $job.WorkloadName
                Status        = $job.Status
                StartTime     = $job.StartTime
                EndTime       = $job.EndTime
                JobID         = $job.JobId
            }
            $backupjobsCol += $backupjobsObject
        }
    }
}

# Get Azure AD Groups with members
$adGroups = Get-AzureADGroup
foreach ($adGroup in $adGroups) {
        $groupMembers = Get-AzureADGroupMember -ObjectId $adGroup.ObjectId
        foreach ($groupMember in $groupMembers) {
            $groupMemberInfoObject = [pscustomobject][Ordered]@{
                Name        = $adGroup.DisplayName
                DisplayName = $groupMember.DisplayName
                Email       = $groupMember.Mail
                UserType    = $groupMember.UserType
                ObjectType  = $adGroup.ObjectType
            }
            $groupMemberInfoCol += $groupMemberInfoObject
        }
    }

# Azure Quota Report // created outside the base inventory report loop for easy removal if not required
$azurelocations = "westeurope","northeurope"

$subscriptions = Get-AzureRmSubscription

foreach ($subscription in $subscriptions) {

    foreach ($azurelocation in $azurelocations) {
         
    Select-AzureRmSubscription -Subscription $subscription.Name
	
	#VM and it's quota
	$vmquota = Get-AzureRmVMUsage -Location $azurelocation | Select-Object Name, CurrentValue, Limit

	#$networkquota 
	$networkquota = Get-AzureRmNetworkUsage -Location $azurelocation | Select-Object Name, CurrentValue, Limit 

	#Loop through network and dump to an array
	$vmquota | ForEach-Object {
		
		$vmobj = [pscustomobject][Ordered]@{  
                Subscription = $Subscription.Name
		        ResourceName = $_.Name.LocalizedValue
		        CurrentlyUsed = $_.CurrentValue
		        Limit = $_.Limit
		        Category = "VM"
                Location = $azurelocation
                }
		$quotaarr += $vmobj
	}
    $networkquota | ForEach-Object {
        $networkobj = [pscustomobject][Ordered]@{ 
                   Subscription = $subscription.Name   
		           ResourceName = $_.Name.LocalizedValue
		           CurrentlyUsed = $_.CurrentValue
		           Limit = $_.Limit
		           Category = "Network"
                   Location = $azurelocation
                   }
        $quotaarr += $networkobj
	    }
	}

  	#Storage quota
	$storagequota = Get-AzureRmStorageUsage | Select-Object LocalizedName, CurrentValue, Limit
	$storagequota | ForEach-Object {		
		$storageobj = [pscustomobject][Ordered]@{ 
                    Subscription = $subscription.Name  
		            ResourceName = $_.LocalizedName
		            CurrentlyUsed = $_.CurrentValue
		            Limit = $_.Limit
		            Category = "Storage"
                    Location = "NA"
                }
		$quotaarr += $storageobj
    }
}

#Excel Export
$sqlInfoCol | Export-Excel $ExcelFile -WorkSheetname 'SQLDatabases' -AutoSize -AutoFilter
$resGroupsCol | Export-Excel $ExcelFile -WorkSheetname 'ResourceGroups' -AutoSize -AutoFilter
$virtualmachinesCol | Export-Excel $ExcelFile -WorkSheetname 'VirtualMachines' -AutoSize -AutoFilter
$vnetsCol | Export-Excel $ExcelFile -WorkSheetname 'VNETs' -AutoSize -AutoFilter
$disksCol | Export-Excel $ExcelFile -WorkSheetname 'Disks' -AutoSize -AutoFilter
$nsgInfoCol | Export-Excel $ExcelFile -WorkSheetname 'NSGs' -AutoSize -AutoFilter
$roleAssignmentInfoCol | Export-Excel $ExcelFile -WorkSheetname 'Roles' -AutoSize -AutoFilter
$groupMemberInfoCol | Export-Excel $ExcelFile -WorkSheetname 'AAD Groups' -AutoSize -AutoFilter
$storageAccountInfoCol | Export-Excel $ExcelFile -WorkSheetname 'StorageAccounts' -AutoSize -AutoFilter
$policyAssignmentsInfoCol | Export-Excel $ExcelFile -WorkSheetname 'Policies' -AutoSize -AutoFilter
$scalesetCol | Export-Excel $ExcelFile -WorkSheetname 'Scalesets' -AutoSize -AutoFilter
$vmsnapshotsCol | Export-Excel $ExcelFile -WorkSheetname 'VMSnapshots' -AutoSize -AutoFilter
$gatewaysCol | Export-Excel $ExcelFile -WorkSheetname 'Gateways' -AutoSize -AutoFilter
$recoveryvaultsCol | Export-Excel $ExcelFile -WorkSheetname 'RecoveryVaults' -AutoSize -AutoFilter
$publicIpsCol | Export-Excel $ExcelFile -WorkSheetname 'PublicIps' -AutoSize -AutoFilter
$peeringsCol | Export-Excel $ExcelFile -WorkSheetname 'Peerings' -AutoSize -AutoFilter
$backupjobsCol | Export-Excel $ExcelFile -WorkSheetname 'BackupJobs' -AutoSize -AutoFilter
$quotaarr | Export-Excel -Path $ExcelFile -Worksheetname 'Subscription Quotas' -AutoSize -AutoFilter

# Get Automation Credentials for sending e-mail
$AutoAccCred = 'InSparkCredentials-JohndeJager'
$smtpCredential = Get-AutomationPSCredential -Name $AutoAccCred

# Send Mail with the overview
$date = get-date
$FromAddress = 'john.de.jager@inspark.nl'
$ToAddress = 'receive@johndejager.com'
$smtpServer = 'smtp.office365.com'
$smtpPort = '587'
$subject = "Azure Inventory - $date"

Send-MailMessage -To $ToAddress `
    -From $FromAddress `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -Credential $smtpCredential `
    -Subject $subject `
    -Attachments $excelFile `
    -Body "Here is the overview of Azure resources" `
    -BodyAsHTML `
    -UseSsl
