Set-ExecutionPolicy Bypass -Scope Process

Import-Module ./ProcessMediaReport.ps1 -Force
Import-Module ./PowerShellSelenium.ps1 -Force
Import-Module ./PoshCanvasNew.ps1 -Force
Import-Module ./BrightCoveSetup.ps1 -Force
Import-Module ./CheckModules.ps1 -Force
Import-Module ./FormatExcel.ps1 -Force
Import-Module ./Notifications.ps1 -Force
Import-Module ./SearchCourse.ps1 -Force

Set-BrightcoveCredentials
$course_id = Read-Host "Enter Course ID"
$sw = [Diagnostics.Stopwatch]::new()
$sw.start()

Get-Modules
Start-Chrome
$chrome.url = "https://signin.brightcove.com/login?redirect=https%3A%2F%2Fstudio.brightcove.com%2Fproducts%2Fvideocloud%2Fmedia"
$chromeWait.until($conditions::ElementIsVisible($by::CssSelector("input[name*='email']"))).sendKeys($BrightcoveCredentials.UserName) | Out-Null
$chromeWait.until($conditions::ElementIsVisible($by::CssSelector("input[id*='password']"))).sendKeys($BrightcoveCredentials.GetNetworkCredential().password) | Out-Null
$chromeWait.until($conditions::ElementIsVisible($by::CssSelector("button[id*='signin']"))).submit() | Out-Null

Search-Course $course_id
Close-Chrome
#Excel formatting
Write-Host 'Formatting Excel Document...' -ForegroundColor Green
Format-MediaExcel1

Get-MediaPivotTables

Write-Host 'Finishing Excel formatting...' -ForegroundColor Green
Format-MediaExcel2

Write-Host "Report Generated" -ForegroundColor Green
Write-Host "WARNING: Number formatting for MediaLength Pivot Charts are not working, change Column B Number Format to Custum: '[h]:mm:ss'`nThe default formatting is in days instead of hours." -ForegroundColor Yellow
$sw.stop()
Send-Notification
