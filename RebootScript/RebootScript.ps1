# Reboot script for some servers
# ComputerList.txt must be filled with hostnames or IPs. 1 line - 1 entry

$MessagePause1 = 120 #120
$MessagePause2 = 60 #60
$RebootPause = 300 #300
$RebootMsg1 = "ALERT: "+($MessagePause1+$MessagePause2)/60+" min. Reboot. Perkrovimas. Перезагрузка"
$RebootMsg2 = "ALERT: "+$MessagePause2/60+" min. Reboot. Perkrovimas. Перезагрузка"

$contents = Get-Content "ComputerList.txt"

if ($contents) {
    foreach ($computer in $contents)
    {
        Write-Host $computer
        msg * /server:$computer /TIME:$MessagePause1 "$RebootMsg1"
        Sleep $MessagePause1
        msg * /server:$computer /TIME:$MessagePause2 "$RebootMsg2"
        Sleep $MessagePause2
        shutdown -r /m \\$computer -t 0 /d p:4:1 /c 'Planned everyday restart by Anton'
        Sleep $RebootPause
    }
}