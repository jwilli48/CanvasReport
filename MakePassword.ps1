#Don't have to ever have your password anywhere as plain text
$PasswordType = "Brightcove"
$SetupCred = Get-Credential -Message "Brightcove Credentials needed"
$secureStringText = $SetupCred.Password | ConvertFrom-SecureString
Set-Content $("$PSScriptRoot\Passwords\My"+ $PasswordType + "Password.txt") $secureStringText
Set-Content $("$PSScriptRoot\Passwords\My"+ $PasswordType + "Username.txt") $SetupCred.UserName
