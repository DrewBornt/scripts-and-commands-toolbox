# Collect User Information
$fullname = Read-Host "Enter First and Last name?"
$password = Read-Host "Password?" -AsSecureString
$firstname, $lastname = $fullname.Split(' ')
$username = $fullname.Replace(' ','')
$groups = Read-Host "What groups to assign them to?"
$groups = $groups.Split(' ')

Import-Module ActiveDirectory

# Create User
New-ADUser -Name $fullname `
-DisplayName $fullname `
-GivenName $firstname `
-Surname $lastname `
-EmailAddress "$username@DOMAIN" `
-SamAccountName $username `
-Path "OU=OUNAME, DC=DOMAIN, DC=DOMAIN" `
-Enabled $True `
-ChangePasswordAtLogon $True `
-UserPrincipalName "$username@DOMAIN" `
-AccountPassword $password

Foreach ($group in $groups)
{
    Add-ADGroupMember `
    -Identity $group `
    -Members $username
}


# Create User Shares and Files

New-Item `
-ItemType "directory" `
-Path "\\SERVERNAME\c`$\Shares\$username"

New-Item `
-ItemType "directory" `
-Path "\\SERVERNAME\c`$\Shares\Scans\$username"

New-SmbShare -Name "$username" -CimSession SERVERNAME -Path "c:\Shares\$username"




# Assign License (TO COME)