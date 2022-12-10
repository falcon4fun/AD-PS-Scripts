# Daily script for extracting full AzureAD configuration. 
# Can be very long to execute on large AADs. 
# On small company AAD execution tooks around 15~ minutes.
# Very suitable for change tracking when runs daily

# Autorization uses Tokens with certificate from App Registration

# Requires (AzureAD or AzureADPreview) 
# and AzureADExporter 
# and MSAL.PS 
# and MGGraph modules. 
# Install them manually for PSGallery

Import-Module AzureADExporter

# Your tenant ID
$tenantid = "00000000-0000-0000-0000-0000000000"
# Your Application ID
$appid = "111111-1111-1111-1111-1111111111"
# Your certificate path
$certificatePath = Get-Item Cert:\CurrentUser\my\<CERTIFICATE THUMBPRINT ID>

$Token = Get-MsalToken -ClientId $appid -TenantId $TenantId -ClientCertificate $certificatePath
Connect-MgGraph -AccessToken $Token.AccessToken

$Path = 'C:\Temp\ExportAD-Logs\' + $((Get-Date).ToString('yyyy-MM-dd'))
New-Item -ItemType Directory -Path $Path | Out-Null
Export-AzureAD -Path $Path -All

Disconnect-MgGraph
