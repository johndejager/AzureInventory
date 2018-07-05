# Login to Azure PowerShell
Login-AzureRmAccount

# Variables
$pwd = 'your password'
$dnsName = 'Your SPN Name'
$subscriptionID = 'your subscription ID' # Can be found in the Azure Portal or by running Get-AzureRMSubscription
$role = 'Reader'
# Create the self signed cert
$currentDate = Get-Date
$endDate = $currentDate.AddYears(2)
$notAfter = $endDate.AddYears(2)
$homePage = 'https://' + $dnsName
$exportPath = 'C:\Temp\' + $dnsName + '.pfx'
$thumb = (New-SelfSignedCertificate -CertStoreLocation cert:\localmachine\my -DnsName $dnsName -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter).Thumbprint
$pwd = ConvertTo-SecureString -String $pwd -Force -AsPlainText
Get-ChildItem -Path "cert:\localmachine\my\$thumb" | Export-PfxCertificate -FilePath $exportPath -Password $pwd
$PFXCert = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($exportPath, $pwd)

# Create the Azure Active Directory Application
$azureAdApplication = New-AzureRmADApplication -DisplayName $dnsName -HomePage $homePage -IdentifierUris $homePage 

# Create the Service Principal and connect it to the Application
New-AzureRMADServicePrincipal -ApplicationId $azureAdApplication.ApplicationId.Guid   `
    -CertValue $keyValue -EndDate $PFXCert.NotAfter -StartDate $PFXCert.NotBefore 

# Give the Service Principal Reader access to a subscription
$subscription = Set-AzureRmContext -Subscription $subscriptionID
$tenantID = $subscription.TenantID
New-AzureRmRoleAssignment -RoleDefinitionName $role -ServicePrincipalName $azureAdApplication.ApplicationId

#Print the values
Write-Output "`nCopy and Paste below values for Service Connection" -Verbose
Write-Output "***************************************************************************"
Write-Output "Connection Name       :  $dnsName"
Write-Output "Service Principal Id  : '$azureAdApplication.ApplicationId.Guid'"
Write-Output "Tenant Id             : $tenantID"
Write-Output "CertificateThumbprint : $thumb"
Write-Output "Subscription Id       : $subscriptionID "
Write-Output "Service Principal key : <Password that you typed in>"
Write-Output "***************************************************************************"


