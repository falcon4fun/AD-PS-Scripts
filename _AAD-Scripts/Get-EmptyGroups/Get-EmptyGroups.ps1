# Gets all AzureAD empty groups

Connect-AzureAD
Get-AzureADGroup -All $true | Where-Object { (Get-AzureADGroupMember -ObjectId $_.ObjectID).count -eq 0 }