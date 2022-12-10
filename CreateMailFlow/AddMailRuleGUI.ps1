# Create MailFlow rule to redirect messages from User to User/Group
# Change Exchange server path inside New-PSSession

###########################################################
Add-Type -AssemblyName System.Windows.Forms

$global:FromVal = $global:ToVal = $InDate = $OutDate = $null

function SearchUser($val) {
    $Val = '*'+$Val+'*'

    $UserList = Get-ADUser -Filter {name -like $Val} -Properties Name,Surname,DistinguishedName | Select-Object -Property Name,Surname,DistinguishedName
    $GroupList = Get-ADGroup -filter {GroupCategory -eq 'Distribution' -And name -like $Val} | Select-Object -Property Name,Surname,DistinguishedName
    $CombinedList = @()
    $CombinedList += $UserList
    $CombinedList += $GroupList
    $Users = $CombinedList | Out-GridView -Title "Select OU and Click OK" -OutputMode Single
    Write-Host $Users
    $Users | ft
    $Users | fl
    Write-Host $Users.Name
    Write-Host $Users.Surname
    Write-Host $Users.DistinguishedName

    if ($Users.DistinguishedName) {
        Write-Host 'User Found'
        Return $Users
    } else {
        Write-Host 'User NOT found'
        Return 1
    }
}


# Main Form
$mainForm = New-Object System.Windows.Forms.Form
$font = New-Object System.Drawing.Font(“Consolas”, 13)
$mainForm.Text = ” Pick Time Frame”
$mainForm.Font = $font
$mainForm.ForeColor = “White”
$mainForm.BackColor = “DarkOliveGreen”
$mainForm.Width = 400
$mainForm.Height = 300

# MinDatePicker Label
$mindatePickerLabel = New-Object System.Windows.Forms.Label
$mindatePickerLabel.Text = “StartDate”
$mindatePickerLabel.Location = “15, 10”
$mindatePickerLabel.Height = 22
$mindatePickerLabel.Width = 90
$mainForm.Controls.Add($minDatePickerLabel)

# MaxDatePicker Label
$maxdatePickerLabel = New-Object System.Windows.Forms.Label
$maxdatePickerLabel.Text = “StopDate”
$maxdatePickerLabel.Location = “15, 75”
$maxdatePickerLabel.Height = 22
$maxdatePickerLabel.Width = 90
$mainForm.Controls.Add($maxDatePickerLabel)

# MinTimePicker Label
$minTimePickerLabel = New-Object System.Windows.Forms.Label
$minTimePickerLabel.Text = “min-time”
$minTimePickerLabel.Location = “15, 38”
$minTimePickerLabel.Height = 22
$minTimePickerLabel.Width = 90
$mainForm.Controls.Add($minTimePickerLabel)

# MaxTimePicker Label
$maxTimePickerLabel = New-Object System.Windows.Forms.Label
$maxTimePickerLabel.Text = “max-time”
$maxTimePickerLabel.Location = “15, 100”
$maxTimePickerLabel.Height = 22
$maxTimePickerLabel.Width = 90
$mainForm.Controls.Add($maxTimePickerLabel)

# DatePicker
$mindatePicker = New-Object System.Windows.Forms.DateTimePicker
$mindatePicker.Location = “110, 7”
$mindatePicker.Width = “150”
$mindatePicker.Format = [windows.forms.datetimepickerFormat]::custom
$mindatePicker.CustomFormat = “yyyy/MM/dd”
$mainForm.Controls.Add($mindatePicker)

# DatePicker
$maxdatePicker = New-Object System.Windows.Forms.DateTimePicker
$maxdatePicker.Location = “110, 70”
$maxdatePicker.Width = “150”
$maxdatePicker.Format = [windows.forms.datetimepickerFormat]::custom
$maxdatePicker.CustomFormat = “yyyyy/MM/dd”
$mainForm.Controls.Add($maxdatePicker)


# MinTimePicker
$minTimePicker = New-Object System.Windows.Forms.DateTimePicker
$minTimePicker.Location = “110, 35”
$minTimePicker.Width = “150”
$minTimePicker.Format = [windows.forms.datetimepickerFormat]::custom
$minTimePicker.CustomFormat = “HH:mm”
$minTimePicker.ShowUpDown = $TRUE
$mainForm.Controls.Add($minTimePicker)

# MaxTimePicker
$maxTimePicker = New-Object System.Windows.Forms.DateTimePicker
$maxTimePicker.Location = “110, 98”
$maxTimePicker.Width = “150”
$maxTimePicker.Format = [windows.forms.datetimepickerFormat]::custom
$maxTimePicker.CustomFormat = “HH:mm”
$maxTimePicker.ShowUpDown = $TRUE
$mainForm.Controls.Add($maxTimePicker)

$checkbox1 = new-object System.Windows.Forms.checkbox
$checkbox1.Location = new-object System.Drawing.Size(100,145)
$checkbox1.Size = new-object System.Drawing.Size(250,20)
$checkbox1.Text = "Enable/Disable"
$checkbox1.Checked = $true
$checkbox1.add_CheckedChanged({
    if ($checkbox1.Checked){
        $checkbox2.Checked = $false
    }
})
$mainForm.Controls.Add($checkbox1);

$checkbox2 = new-object System.Windows.Forms.checkbox
$checkbox2.Location = new-object System.Drawing.Size(100,165)
$checkbox2.Size = new-object System.Drawing.Size(250,20)
$checkbox2.Text = "QUIT?"
$checkbox2.Checked = $false
$checkbox2.add_CheckedChanged({
    if ($checkbox2.Checked){
        $checkbox1.Checked = $false
    }
})
$mainForm.Controls.Add($checkbox2);

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = “15, 185”
$textBox1.Width = “150”
$textBox1.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $textBox1.Text | Out-Host
        $global:FromVal = SearchUser($textBox1.Text)
        $global:FromVal | fl
        $label1.Text = $global:FromVal.Name
    }
})
$mainForm.Controls.Add($textBox1)

$label1 = New-Object System.Windows.Forms.Label
$label1.Text = “From User”
$label1.Font = New-Object System.Drawing.Font("Consolas",10)
$label1.Location = “170, 190”
$label1.Width = “250”
$mainForm.Controls.Add($label1)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = “15, 210”
$textBox2.Width = “150”
$textBox2.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $textBox2.Text | Out-Host
        $global:ToVal = SearchUser($textBox2.Text)
        $global:ToVal | fl
        $label2.Text = $global:ToVal.Name
    }
})
$mainForm.Controls.Add($textBox2)

$label2 = New-Object System.Windows.Forms.Label
$label2.Text = “To User/Group”
$label2.Font = New-Object System.Drawing.Font("Consolas",10)
$label2.Location = “170, 215”
$label2.Width = “250”
$mainForm.Controls.Add($label2)

# OD Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = “15, 150”
$okButton.ForeColor = “Black”
$okButton.BackColor = “White”
$okButton.Text = “OK”
$okButton.add_Click({
    If ( ($global:FromVal) -and ($global:ToVal) ){
        $mainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $mainForm.close()
    }
})
$mainForm.Controls.Add($okButton)


$result = $mainForm.ShowDialog()
Write-Host "Result: $result "

if ($result -eq [Windows.Forms.DialogResult]::OK) {

    $FromUser = $global:FromVal
    $ToUser = $global:ToVal

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session

    $InDate = [DateTime]"$($mindatePicker.Text) $($minTimePicker.Text)"
    $OutDate = [DateTime]"$($maxdatePicker.Text) $($maxTimePicker.Text)"

    Write-Host "In: $InDate "
    Write-Host "Out: $OutDate "

    Write-Host "FromUser: $FromUser"
    Write-Host "ToUser: $ToUser"

    If ($ToUser.Surname) {
        $FromTo = $FromUser.Surname+" > "+$ToUser.Surname
    } ElseIf ($ToUser.Name) {
        $FromTo = $FromUser.Surname+" > "+$ToUser.Name
    }

    If ($FromTo.Length -gt 36) {
        $FromTo = $FromTo.subString(0, [System.Math]::Min(36, $FromTo.Length))
    }

    $comment = "Autocreated rule from script"
        If ( $checkbox1.Checked ) {
            $ruleName = "BCC: "+$FromTo+" "+ $($mindatePicker.Text)+"-"+$($maxdatePicker.Text)
        } elseif ( $checkbox2.Checked ) {
            $ruleName = "QUIT: "+$FromTo
        } else {
            $ruleName = "BCC: "+$FromTo
        }


    If ( $checkbox1.Checked ) {
        New-TransportRule -ActivationDate ($InDate) -ExpiryDate ($OutDate) -BlindCopyTo $ToUser.DistinguishedName -Comments $comment -Name $ruleName -SentTo $FromUser.DistinguishedName
    } else {
        New-TransportRule -BlindCopyTo $ToUser.DistinguishedName -Comments $comment -Name $ruleName -SentTo $FromUser.DistinguishedName
    }

    if( -not $? )
    {
        $Error | fl
        [System.Windows.MessageBox]::Show($Error[0].Exception.Message,'Error',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    }

    Remove-PSSession $Session
} else {
    Write-Host "Cancel Pressed"
}
