<#

Need RSAT SERVER MANAGER to be installed, where script is launching
ActiveDirectory module will not work without it.

Script can:
1. Create user
2. Select OU
3. Select Groups
4. Create email in Exchange
5. Add secondary SMTP address
6. Create MS Navision Account

Change Improt-Module path for Exchange
Change vars: CompanyName, DefPass and etc.

#>

#Import needed module.
Import-Module ActiveDirectory

function AddToGroups ($samAcc) {
    $ADGroups = Get-ADGroup -Filter * | Out-GridView -Title "Choose Groups" -PassThru

    # If servers have not been selected write warning in host
    If ( !$ADGroups ) {
        Write-Warning "Groups have not been selected"
    } Else {
        # Display server names and their IP addresses
        Write-Warning "The following servers have been selected:"
 
        $ADGroups | Format-Table
 
        # Confirm if you want to proceed
        Write-Host -nonewline "Do you want to proceed? (Y/N): "
        $Response = Read-Host
        Write-Host " "
 
        # If response was different that Y script will end
        If ( $Response -ne "Y" ) {
            Write-Warning "Script ends"      
        } Else {
            # Servers loop
            ForEach ($ADGroup in $ADGroups.SamAccountName) {
                # Restart command
                Write-Warning "Selected $ADGroup :"
                $ADGroup | Format-List
                Add-ADGroupMember -Identity $ADGroup -Members $samAcc
            }
        }
    }
}

function CreateMailAD ($samAcc) {
    #Add user to mail
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session

    #Add acc here
    Enable-Mailbox -Identity $samAcc | Out-Null
    Set-Mailbox -Identity $samAcc -IssueWarningQuota $DefQuota1 -ProhibitSendQuota $DefQuota2 -ProhibitSendReceiveQuota $DefQuota3 -UseDatabaseQuotaDefaults $false
    Write-Host 'Mailbox Created'

    if (Get-Mailbox -Identity $samAcc) {
        $mailbox = (Get-Mailbox -Identity $samAcc).PrimarySmtpAddress
        Write-Host 'Check: Mailbox Created'
        Write-Host $mailbox
        #        Set-Clipboard -Value $mailbox
    } else {
        Remove-PSSession $Session
        throw 'Error creating mailbox. Code 11. Exiting..'
    }

    $msg = 'Add second mail? [Y/N]'
    $response = Read-Host -Prompt $msg
    if ($response -eq 'y') {
        $secondAddress = Read-Host "Enter second email"
        Write-Warning 'Second mail:'
        Write-Warning $secondAddress
        Set-Mailbox -Identity $samAcc -EmailAddresses @{add = "$secondAddress" }

        $msg = 'Make it primary mail? [Y/N]'
        $response = Read-Host -Prompt $msg
        if ($response -eq 'y') {
            Set-Mailbox -Identity $samAcc -PrimarySmtpAddress "$secondAddress" -EmailAddressPolicyEnabled $false
        }
    }

    $PrimaryMail = (Get-Mailbox -Identity $samAcc).PrimarySMTPAddress
    Write-Host 'Primary Mail: $PrimaryMail'
    Set-Clipboard -Value $PrimaryMail

    Remove-PSSession $Session
}

function CreateNavAD ($samAcc) {
    Import-Module "\\nav\Install\Dynamics.365.BC.39327.RU.DVD\WindowsPowerShellScripts\Cloud\NAVRemoteAdministration\NAVRemoteAdministration.psm1"
    $NavSession = New-NAVAdminSession -RemoteMachineAddress nav
    Import-PSSession $NavSession
    #Get-NAVServerUser -ServerInstance prod
    
    $navuser = Get-NAVServerUser -ServerInstance Prod

    if (!($navuser.UserName -contains "$DomainNavName\$samAcc")) {
        #    if (!($navuser.UserPrincipalName -contains "$samAcc@$DomainNavName.local")) {
        Write-Host 1
        New-NAVServerUser -ServerInstance Prod -WindowsAccount "$DomainNavName\$samAcc" -FullName "$fname $lname" -LicenseType Full -State Enabled
        New-NAVServerUserPermissionSet -PermissionSetId SUPER -CompanyName "$CompanyName" -ServerInstance Prod -WindowsAccount "$DomainNavName\$samAcc"
        Write-Host "Nav user created"
    } else {
        Remove-PSSession $NavSession 
        throw 'Error: NAV user already exists. Code 13. Exiting..'
    }

    Remove-PSSession $NavSession
}

function Remove-Diacritics {
    param ([String]$src = [String]::Empty)
    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
  ($normalized -replace '\p{M}', '')
}

$CompanyName = "TestCompany Ltd"
$DefPass = "test!1234"
$DefQuota1 = [Math]::Round(3GB)
$DefQuota2 = [Math]::Round(3.2GB)
$DefQuota3 = [Math]::Round(3.3GB)
$DomainName = "domain.lan"
$DomainNavName = (Get-ADDomain).NetBIOSName
$ExchangeServer = "exchangeServer"
$PrimaryMail = ""
$SearchBase = "DC=DOMAIN,DC=LAN"

$name = Read-Host "Enter 'Name Surname'"

#Convert tab to spaces
$name = $name -replace '\t', '  '
$nameArr = $name.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
Write-Host $nameArr
Write-Host $nameArr.Length


if ($nameArr.length -eq 2) {
    Write-Host $nameArr[0]
    Write-Host $nameArr[1]
    $fname = [cultureinfo]::GetCultureInfo("en-US").TextInfo.ToTitleCase($nameArr[0])
    $lname = [cultureinfo]::GetCultureInfo("en-US").TextInfo.ToTitleCase($nameArr[1])

    Write-Host "Length" $fname.Length+($lname.Length)

    if ( ($fname.Length + $lname.Length) -ge 20) { 
        if ( ($lname.Length + 2) -ge 20) {
            Write-Host 4
            throw 'Hule takaja faimilija bolshaja? Exiting.. Code 4' 
        } else { 
            Write-Host 1
            $samAcc = [cultureinfo]::GetCultureInfo("en-US").TextInfo.ToTitleCase($fname.Substring(0, 1)).Trim() + "." + [cultureinfo]::GetCultureInfo("en-US").TextInfo.ToTitleCase($lname).Trim()
        }
    } else {
        if ( ($lname.Length + 2) -gt 20) {
            Write-Host 3
            throw 'Hule takaja faimilija bolshaja? Exiting.. Code 3'
        } else {
            Write-Host 2
            $samAcc = [cultureinfo]::GetCultureInfo("en-US").TextInfo.ToTitleCase($fname).Trim() + "." + [cultureinfo]::GetCultureInfo("en-US").TextInfo.ToTitleCase($lname).Trim()
        }
    }
} Else {
    throw 'More than 2 words exists in Name. Exiting..'
}

$samAcc = Remove-Diacritics($samAcc)
Write-Host($samAcc)

$OUList = Get-ADOrganizationalUnit -SearchBase $SearchBase -Filter * -Properties Name, DistinguishedName | Select-Object -Property Name, DistinguishedName
$OU = $OUList | Out-GridView -Title "Select OU and Click OK" -OutputMode Single
Write-Host $OU.DistinguishedName

#Account will be created in the OU listed in the $OU variable in the CSV file; don’t forget to change the domain name in the"-UserPrincipalName" variable
New-ADUser `
    -SamAccountName $samAcc `
    -UserPrincipalName "$samAcc@$DomainName" `
    -Name "$fname $lname" `
    -GivenName $fname `
    -Surname $lname `
    -Enabled $True `
    -ChangePasswordAtLogon $False `
    -DisplayName "$fname $lname" `
    -Path $OU.DistinguishedName `
    -AccountPassword (convertto-securestring $DefPass -AsPlainText -Force)

if (Get-ADUser -Identity $samAcc) {
    Write-Host 'User Created'
} else {
    throw 'Error creating user. Code 10. Exiting..'
}

$msg = 'Add user to Groups? [Y/N]'
do {
    $response = Read-Host -Prompt $msg
    if ($response -eq 'y') {
        AddToGroups($samAcc)
    }
} until ($response -eq 'n')


$msg = 'Create Mail? [Y/N]'
$response = Read-Host -Prompt $msg
if ($response -eq 'y') {
    CreateMailAD($samAcc)
}

$msg = 'Create Dynamics Nav? [Y/N]'
$response = Read-Host -Prompt $msg
if ($response -eq 'y') {
    CreateNavAD($samAcc)
}

