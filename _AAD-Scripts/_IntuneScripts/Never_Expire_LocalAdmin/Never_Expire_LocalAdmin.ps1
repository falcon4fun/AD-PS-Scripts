# Sets local admin to NeverExpire
# Dunno why but it cant be set thru Intune using in-built features.

Set-LocalUser -Name "Admin" -PasswordNeverExpires:$true
