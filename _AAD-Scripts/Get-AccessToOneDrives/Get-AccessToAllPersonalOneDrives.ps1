# Add admin to access all company personal OneDrives

Connect-SPOService -Url https://domaincom-admin.sharepoint.com

$Admin = "admin@domain.com"
$AllSites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"

$AllSites | ForEach-Object { 
    Set-SPOUser -Site $_.Url -LoginName $Admin -IsSiteCollectionAdmin $True
    Write-host "Added Site collection Admin to" $_.URL
}