# Fixes problem when username have non-English symbols.
# Applyes to AAD only.

# On-Prem uses sAMAccountName to create user inside SystemDir:\Users\<UserName>
# AAD uses DisplayName to create it. 
# If username have non-english symbols, many programs will have problems. I.e. Outlook will crash with unexpected error on start
# So system locale should be changed using script or manually (WTF are you doing here if you do this manually?)

# Example language settings for non-Unicode applications
Set-WinSystemLocale -SystemLocale lt-LT
