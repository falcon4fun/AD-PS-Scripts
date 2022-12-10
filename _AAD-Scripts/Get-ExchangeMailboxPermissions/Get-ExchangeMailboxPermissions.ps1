# Gets all mailbox with delegated rights to another users
# Authorization uses certificate from App Registration

# Your domain
$org = "your_domain.com"
# Your Application ID
$appid = "111111-1111-1111-1111-1111111111"
# Your certificate path
$certificatePath = Get-Item Cert:\CurrentUser\my\<CERTIFICATE THUMBPRINT ID>
Connect-ExchangeOnline -Organization $org -AppId $appid -Certificate $certificatePath

Get-Mailbox -resultsize unlimited | Get-MailboxPermission |`
    Select-Object Identity, User, Deny, AccessRights, IsInherited |`
    Where-Object -Property User -NE "NT AUTHORITY\SELF" |`
    Export-Csv -Path "C:\Temp\MailboxPermissions.csv" –NoTypeInformation -Encoding UTF8