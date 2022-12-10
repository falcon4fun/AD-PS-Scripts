# Script extracts all logs from AAD.
# Should be runned daily
# Writes output to one CSV. Should be CSV. 
# Tracks latest extraction from CSV and tryes to extract [latest extraction date-1]. Remove duplicates.
# Splits all day to 5k entries because of API limitation. If >5k entries/hour, tryes to split quarterly (15min). Can be modied to exend even more detaily.
# Runs quite slow. Around 10 minutes for 2 days.

# Some preparation required to Registred Application worked had access to Connect-ExchangeOnline command because App should have additional permissions which can be assigned using Graph API:
# https://office365itpros.com/2022/10/13/exchange-online-powershell-app/

# Your domain
$org = "your_domain.com"
# Your Application ID
$appid = "111111-1111-1111-1111-1111111111"
# Your certificate path
$certificatePath = Get-Item Cert:\CurrentUser\my\<CERTIFICATE THUMBPRINT ID>
$OutputFile = "C:\Temp\UnifiedAuditLog.csv"

$Today = Get-Date -Date (Get-Date -Format “yyyy-MM-dd”)
$Counter = 0
$ArrMinutes = @(0, 45, 30, 15, 0)
$Report = [System.Collections.Generic.List[Object]]::new()

$Date1 = Get-Content -First 2 $OutputFile | Select-Object -Index 1

If ($null -ne $Date1) {
    $Date1 = $Date1.Split(",")[0]
    $Date1 = $Date1.Replace("`"", "")
    $Date2 = [datetime]::ParseExact($Date1, 'yyyy-MM-ddTHH:mm:ss', $null)
    $intDays = (new-timespan -Start $Date2 -End (get-date)).Days + 1
} else {
    $intDays = 1
}

#$intDays = 140
Remove-Variable -Name ("Date1", "Date2")
Write-Host $intDays

Connect-ExchangeOnline -Organization $org -AppId $appid -Certificate $certificatePath

For ($i = 0; $i -le $intDays; $i++) {
    For ($j = 23; $j -ge 0; $j--) {
        $StartDate = ($Today.AddDays(-$i)).AddHours($j)
        $EndDate = ($Today.AddDays(-$i)).AddHours($j + 1)
        $Audit = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000

        If ($Audit.Count -eq 5000) {
            ForEach ($index in (1..($ArrMinutes.Count - 1))) {
                If ($index -eq 1) {
                    $EndDate = ($Today.AddDays(-$i)).AddHours($j + 1).AddMinutes($ArrMinutes[$index - 1])
                } else {
                    $EndDate = ($Today.AddDays(-$i)).AddHours($j).AddMinutes($ArrMinutes[$index - 1])
                }
                $StartDate = ($Today.AddDays(-$i)).AddHours($j).AddMinutes($ArrMinutes[$index])

                $Audit = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000
                $Records = $Audit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
                
                If ($Audit.Count -ne 0) {
                    ForEach ($Rec in $Records) {
                        $ReportLine = [PSCustomObject] @{
                            CreationTime                  = $Rec.CreationTime
                            Id                            = $Rec.Id
                            Operation                     = $Rec.Operation
                            OrganizationId                = $Rec.OrganizationId
                            RecordType                    = $Rec.RecordType
                            ResultStatus                  = $Rec.ResultStatus
                            UserKey                       = $Rec.UserKey
                            UserType                      = $Rec.UserType
                            Version                       = $Rec.Version
                            Workload                      = $Rec.Workload
                            ClientIP                      = $Rec.ClientIP
                            ObjectId                      = $Rec.ObjectId
                            UserId                        = $Rec.UserId
                            AzureActiveDirectoryEventType = $Rec.AzureActiveDirectoryEventType
                            SiteUrl                       = $Rec.SiteUrl
                            SourceFileName                = $Rec.SourceFileName
                            SourceRelativeUrl             = $Rec.SourceRelativeUrl
                        }
                        $Report.Add($ReportLine)
                    }
                }

                Write-Host $StartDate `t $Audit.Count
                $Counter += $Audit.Count
            }
        } else {
            $Records = $Audit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json

            If ($Audit.Count -ne 0) {
                ForEach ($Rec in $Records) {
                    $ReportLine = [PSCustomObject] @{
                        CreationTime                  = $Rec.CreationTime
                        Id                            = $Rec.Id
                        Operation                     = $Rec.Operation
                        OrganizationId                = $Rec.OrganizationId
                        RecordType                    = $Rec.RecordType
                        ResultStatus                  = $Rec.ResultStatus
                        UserKey                       = $Rec.UserKey
                        UserType                      = $Rec.UserType
                        Version                       = $Rec.Version
                        Workload                      = $Rec.Workload
                        ClientIP                      = $Rec.ClientIP
                        ObjectId                      = $Rec.ObjectId
                        UserId                        = $Rec.UserId
                        AzureActiveDirectoryEventType = $Rec.AzureActiveDirectoryEventType
                        SiteUrl                       = $Rec.SiteUrl
                        SourceFileName                = $Rec.SourceFileName
                        SourceRelativeUrl             = $Rec.SourceRelativeUrl
                    }
                    $Report.Add($ReportLine)
                }
            }

            Write-Host $StartDate `t $Audit.Count
            $Counter += $Audit.Count
        }
    }
}

Write-Host "Total: " `t $Counter

$Report | Export-Csv $OutputFile -Encoding UTF8 -NoTypeInformation -Append
$ReportNew = Import-Csv $OutputFile -Encoding UTF8
$ReportNew | Sort-Object -Uniq * | Sort-Object CreationTime -Descending | Export-Csv $OutputFile -Encoding UTF8 -NoTypeInformation

Disconnect-ExchangeOnline -Confirm:$false
