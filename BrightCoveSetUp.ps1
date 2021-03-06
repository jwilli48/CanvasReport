#Brightcove
function Set-BrightcoveCredentials {
    <#
  .DESCRIPTION

  If they don't have credentials, will ask them for some and save a file with username and a secure string password, then it will set the username and password to the variables saved in text files.
  #>
    if (-not (Test-Path "$PSScriptRoot\Passwords")) {
        New-Item -Path "$PsScriptRoot\Passwords" -ItemType Directory
    }
    if (Test-Path "$PSScriptRoot\Passwords\MyBrightcovePassword.txt") {}
    else {
        Write-Host "No Brightcove Credentials found, please enter them now." -ForegroundColor Yellow

        . "$home/Desktop/AccessibilityTools/CanvasReport-master/MakePassword.ps1"

        Write-Host "WARNING: If Brightcove fails to login this script will also continue to throw errors.`nYou may also need to go to the file $PsScriptRoot\Passwords.ps1 and change the username variable there" -ForegroundColor Yellow
    }
    try {
        $username = Get-Content "$PSScriptRoot\Passwords\MyBrightcoveUsername.txt"
        $password = Get-Content "$PSScriptRoot\Passwords\MyBrightcovePassword.txt"
        $securePwd = $password | ConvertTo-SecureString
        $Global:BrightcoveCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd
        Write-Host "WARNING: If Brightcove fails to login you may have saved the wrong Username and Password, please go to $PSScriptRoot\Passwords and delete the text files there to reset them" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Your password and username files at $PSScriptRoot\Passwords threw an error, they may be empty, please delete them and run the program again" -ForegroundColor Red
        while ($true) {
            Read-Host "Please close the program"
        }
    }
}
