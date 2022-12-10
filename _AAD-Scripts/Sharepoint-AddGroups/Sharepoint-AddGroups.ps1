# Lazy script for initially adding required groups to Sharepoint site and populating with users.
# Not beautified too much. Was in hurry

# Creates groups from GroupNameArray. Example Sec_GroupName1_Group
# Add AAD Admin as Owner

# Next step should be assigning those groups to SPO site

Connect-AzureAD -TenantId '<your tenant id>'

$GroupNamesArr = "GroupName1", "GroupName2", "GroupName3"

Foreach ($GroupName in $GroupNamesArr) {
    $NewGroupName = "Sec_" + $GroupName + "_Group"
    $NewDescription = "Sharepoint Security Group for " + $GroupName 
    Write-Host $NewGroupName
    Write-Host $NewDescription

    New-AzureADGroup -DisplayName $NewGroupName -SecurityEnabled $true -Description $NewDescription -MailEnabled $false -MailNickName "NotSet"
    #Break
}

# Because Azure AD Too SLOW
Start-Sleep -s 5

# Add admin to owner
$OwnerObj = Get-AzureADUser -SearchString "<AAD admin Username>"

Foreach ($GroupName in $GroupNamesArr) {
    $GroupNameFixed = "Sec_" + $GroupName
    $GroupObj = Get-AzureADGroup -SearchString $GroupNameFixed
    Add-AzureADGroupOwner -ObjectId $GroupObj.ObjectId -RefObjectId $OwnerObj.ObjectId
}


# Add users to corresponding groups
Function _AddUsersToGroup ( $UsrArr, $Grp ) {
    Write-Host $UsrArr
    Write-Host $Grp

    $GroupObj = Get-AzureADGroup -SearchString $Grp

    ForEach ($Usr in $UsrArr) {
        $UserObj = Get-AzureADUser -SearchString $Usr
        Add-AzureADGroupMember -ObjectId $GroupObj.ObjectId -RefObjectId $UserObj.ObjectId
    }
}


$UserNamesArr = "FirstName1 Lastname1", "FirstName2 Lastname2"
$GroupName = "Sec_GroupName1"
_AddUsersToGroup $UserNamesArr $GroupName

$UserNamesArr = "FirstName3 Lastname4", "FirstName5 Lastname6"
$GroupName = "Sec_GroupName2"
_AddUsersToGroup $UserNamesArr $GroupName

$UserNamesArr = "FirstName1 Lastname1", "FirstName2 Lastname2"
$GroupName = "Sec_GroupName3"
_AddUsersToGroup $UserNamesArr $GroupName

Pause