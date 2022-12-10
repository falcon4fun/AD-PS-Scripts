# PowerShell 7 Only!
# Checks on which computer / server user has active session
# ThreadLimit - 500
# Works quite fast (:


Function Get-OnlineComputers() {
    
    $username = Read-Host -Prompt "Enter username or part of Name/Surname"
    $username = "*" + $username + "*"

    $Limit = (Get-Date).AddHours(-8)
    $computers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows*' -and Enabled -eq $True -and LastLogonTimestamp -lt $Limit } -Properties * | Sort-Object
    #$computers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows*' -and Enabled -eq $True -and LastLogonTimestamp -lt $Limit } -Properties * | Where-Object {$_.Name -like "*"} | Sort-Object
    
    $computers | ForEach-Object -Parallel {
        $computerName = $_.Name

        $sessions = qwinsta /server $computerName 2>$null  | Where-Object { $_ -notmatch '^ SESSIONNAME' }  | ForEach-Object {
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

        $a = $using:username
        $b = $computerName

        $sessions | ForEach-Object -Parallel {
            if ( $_.Username -like $using:a -and $_.Username.Length -gt 1 ) {

                $c = @"
Server: $using:b
Username: $($_.Username)
SessionID: $($_.Id)
State: $($_.State)
-----------------------------
"@
                Write-Host $c -ForegroundColor Cyan
            }
        } -ThrottleLimit 500
    } -ThrottleLimit 500
}

(Measure-Command {
    Get-OnlineComputers
}).TotalSeconds
pause