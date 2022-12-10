# Check where user is logged.
# Uses array with RDS / Terminal servers hostnames

#Start-Transcript -path C:\output.txt -append

$username = Read-Host -Prompt "Enter username or part of Name/Surname"
$username = "*" + $username + "*"

#$computers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' -and Enabled -eq $True } -Properties OperatingSystem | Where-Object {$_.Name -match "^T\d"} | Sort-Object
$computersArray = "Terminal1", "Terminal2", "Servers1", "Servers2", "Servers3"

#$computers | ft

foreach ($computer in $computersArray) {
    Write-Host $computer.Name
    $sessions = qwinsta /server $computer | Where-Object { $_ -notmatch '^ SESSIONNAME' } | ForEach-Object {
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

    #$sessions | ft

    if ($sessions) {
        foreach ($session in $sessions) {
            if ( $session.Username -like $username -and $session.Username.Length -gt 1 ) {
                Write-Host "Server: "$computer
                Write-Host "Username: "$session.Username
                Write-Host "SessionID: "$session.Id
                Write-Host "State: "$session.State
            }
        }
    }
    #Exit
}

Pause

#Stop-Transcript