$Passwords = "PATH-TO-A-COMMON-PASSWORDS-LIST"
 
$Params = @{
    "All"         = $True
    "Server"      = 'DOMAIN-CONTROLLER-NAME'
    "NamingContext" = 'dc=$$$,dc=$$$'
}
 
Get-ADReplAccount @Params | Test-PasswordQuality -WeakPasswordsFile $Passwords | Out-File "C:\Logs\PasswordAudit.txt"