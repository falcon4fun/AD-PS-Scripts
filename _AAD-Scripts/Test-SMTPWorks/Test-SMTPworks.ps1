# Send message using SMTP
# Test authorization and Conditional Access

$msolcred = get-credential -Credential scannermail@domain.com
Send-MailMessage –From scannermail@domain.com –To admin@domain.com –Subject “Test Email” –Body “Test SMTP Service” -SmtpServer smtp.office365.com -Credential $msolcred -UseSsl -Port 587