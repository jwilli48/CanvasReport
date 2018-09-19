Set-ExecutionPolicy Bypass -Scope Process
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript"){
  $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else{
  $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}
Get-ChildItem -Recurse -Path "$ScriptPath" | Unblock-File

Import-Module ./ProcessMediaReport.ps1 -Force
Import-Module ./PowerShellSelenium.ps1 -Force
Import-Module ./PoshCanvasNew.ps1 -Force
Import-Module ./BrightCoveSetup.ps1 -Force
Import-Module ./CheckModules.ps1 -Force
Import-Module ./FormatExcel.ps1 -Force
Import-Module ./Notifications.ps1 -Force
Import-Module ./SearchCourse.ps1 -Force

Set-BrightcoveCredentials
Get-GoogleApi

./SetDomain.ps1
$course_id = Read-Host "Enter Canvas Course ID or path to course HTML files"
$directory = $FALSE

if($course_id -match "[A-Z]:\\"){
  $directory = $true
}

$sw = [Diagnostics.Stopwatch]::new()
$sw.start()

Get-Modules
Write-Host -ForegroundColor Magenta "Starting chrome..."
Start-Chrome -Headless
$chrome.url = "https://signin.brightcove.com/login?redirect=https%3A%2F%2Fstudio.brightcove.com%2Fproducts%2Fvideocloud%2Fmedia"
$chromeWait.until($conditions::ElementIsVisible($by::CssSelector("input[name*='email']"))).sendKeys($BrightcoveCredentials.UserName) | Out-Null
$chromeWait.until($conditions::ElementIsVisible($by::CssSelector("input[id*='password']"))).sendKeys($BrightcoveCredentials.GetNetworkCredential().password) | Out-Null
$chromeWait.until($conditions::ElementIsVisible($by::CssSelector("button[id*='signin']"))).submit() | Out-Null

if($directory){
  Search-Directory $course_id
}else{
  Search-Course $course_id
}

Close-Chrome
#Excel formatting
Write-Host 'Formatting Excel Document...' -ForegroundColor Green
Format-MediaExcel1

Get-MediaPivotTables

Write-Host 'Finishing Excel formatting...' -ForegroundColor Green
Format-MediaExcel2

Write-Host "Report Generated" -ForegroundColor Green
$sw.stop()
Send-Notification
