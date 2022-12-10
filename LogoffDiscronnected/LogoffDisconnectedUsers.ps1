# Logoff disconnected users from terminals before Daily backup

#Start-Transcript -path C:\output.txt -append

#Matching server starting from T with numbers. Example: T1, T2, .., T124125
#Our terminal servers (RDS) have names T1, T2, T3, ..

$computers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' -and Enabled -eq $True } -Properties OperatingSystem | Where-Object { $_.Name -match "^T\d" } | Sort-Object
$computers | Format-Table

foreach ($computer in $computers) {
    Write-Host $computer.Name
    $sessions = qwinsta /server $computer.Name | Where-Object { $_ -notmatch '^ SESSIONNAME' } | ForEach-Object {
        $item = "" | Select-Object "Active", "SessionName", "Username", "Id", "State", "Type", "Device"
        $item.Active = $_.Substring(0, 1) -match '>'
        $item.SessionName = $_.Substring(1, 18).Trim()
        $item.Username = $_.Substring(19, 20).Trim()
        $item.Id = $_.Substring(39, 9).Trim()
        $item.State = $_.Substring(48, 8).Trim()
        $item.Type = $_.Substring(56, 12).Trim()
        $item.Device = $_.Substring(68).Trim()
        $item
    }

    $sessions | Format-Table

    if ($sessions) {
        foreach ($session in $sessions) {
            if ( ($session.Username -ne "" -or $session.Username.Length -gt 1) -And ($session.State -eq "Disc") ) {
                logoff /server $computer.Name $session.Id
                Write-Host $session.Username
                Write-Host $session.Id
                #Exit
            }
        }
    }
    #Exit
}

#Stop-Transcript