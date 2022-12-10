# Check what browsers users use on workgroup computers

$creds = Get-Credential -UserName 1\Administrator -Message 1

$Computer = "Computer1.domain.lan", "Computer2.domain.lan", "Computer3.domain.lan", "Computer4.domain.lan"

$Computer | ForEach-Object {
    If (Test-Connection $_ -Count 1 -ErrorAction Continue) {
        New-PSSession $_ -Credential $creds
    }
}

$Sessions = Get-PSSession
#Invoke-Command -Session $Sessions -ScriptBlock { hostname; Get-Process msedgewebview2 | ft; Get-Process chrome | ft; }


foreach ($session in $Sessions) {
    Invoke-Command -Session $Session -ScriptBlock { hostname; Get-Process msedge -ErrorAction SilentlyContinue | Format-Table; Get-Process firefox -ErrorAction SilentlyContinue | Format-Table; Get-Process chrome -ErrorAction SilentlyContinue | Format-Table; }
}

Remove-PSSession -Session $Sessions
