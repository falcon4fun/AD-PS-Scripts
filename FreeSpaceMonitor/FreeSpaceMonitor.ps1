# Free space monitoring script on all AD servers

# Sends warning notifications if partition <15 GB. Critical notifications if partition <3 GB
# Checks only disks >1GB
# Change MailParams. Use your SMTP server and your mails


$servers = (Get-ADComputer -Filter 'operatingsystem -like "*server*"').Name


### Get only DriveType 3 (Local Disks) foreach Server
ForEach ($s in $servers) {

    Write-Host "Host: $s "
    $Report = Get-WmiObject win32_logicaldisk -ComputerName $s -Filter "Drivetype=3" -ErrorAction SilentlyContinue -AsJob | Wait-Job -Timeout 5 | Receive-Job | Where-Object { ($_.freespace / $_.size) -le '0.15' }
    $View = ($Report.DeviceID -join ",").Replace(":", "")


    if ($Report) {
        Write-Host "Host: $s "
        Write-Host "Report: $Report "
        Write-Host "View: $View "

        ForEach ($r in $Report) {
            #$r | fl
            Write-Host "Disk: $r.DeviceID"
            $FreeSizeMB = $([math]::floor($r.FreeSpace / 1MB))
            $DiskSizeMB = $([math]::floor($r.Size / 1MB))
            Write-Host "DiskSize: $DiskSizeMB MB"
            Write-Host "FreeSize: $FreeSizeMB MB"
            $DeviceID = $r.DeviceID

            if (($FreeSizeMB -lt 3000) -and ($DiskSizeMB -gt 1000)) {
                #Critical <3GB
                #ALERT!
                Write-Host "!!!!!<3 GB!!!!!! "
                # Host $s, Disk $r.DeviceID, FreeSizeMB $([math]::floor($r.FreeSpace/1MB)), DiskSizeMB $([math]::floor($r.Size/1MB))
                $subject = "[$s] Critical size on disk $DeviceID"
                $msg = "Host: $s<br />Disk: $DeviceID<br />DiskSize: $DiskSizeMB MB<br />FreeSize: $FreeSizeMB MB<br />"
                $subject | Format-List
                $msg | Format-List
                
                $MailParams = @{
                    "From"       = "SpaceMonitor <SpaceMonitor@transimeksa.com>"
                    "To"         = "WhereToSendNoti@domain.com", "WhereToSendNoti2@domain.com"
                    "Subject"    = $subject
                    "Body"       = $msg
                    "SmtpServer" = "<IP OF SMTP SERVER>"
                }
    
                Send-MailMessage -BodyAsHtml @MailParams

            } elseif (($FreeSizeMB -lt 15000) -and ($DiskSizeMB -gt 1000)) {
                #Warning <15GB
                #ALERT!
                Write-Host "!!!!!<15GB !!!!!!! "
                $subject = "[$s] Warning size on disk $DeviceID"
                $msg = "Host: $s<br />Disk: $DeviceID<br />DiskSize: $DiskSizeMB MB<br />FreeSize: $FreeSizeMB MB<br />"
                $subject | Format-List
                $msg | Format-List

                $MailParams = @{
                    "From"       = "SpaceMonitor <SpaceMonitor@domain.com>"
                    "To"         = "WhereToSendNoti@domain.com", "WhereToSendNoti2@domain.com"
                    "Subject"    = $subject
                    "Body"       = $msg
                    "SmtpServer" = "<IP OF SMTP SERVER>"
                }
    
                Send-MailMessage -BodyAsHtml @MailParams

            } elseif ($DiskSizeMB -gt 1000) {
                Write-Host "?????11111?????? "
            } else {
        
            }
            Write-Host ""
        }
        Write-Host ""
    }
    Write-Host "-------------------"
}
