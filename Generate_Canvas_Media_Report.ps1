Set-ExecutionPolicy Bypass -Scope Process
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else {
    $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}
Get-ChildItem -Recurse -Path "$ScriptPath" | Unblock-File

Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/ProcessMediaReport.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/PowerShellSelenium.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/PoshCanvasNew.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/BrightCoveSetup.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/CheckModules.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/FormatExcel.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/Notifications.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/SearchCourse.ps1" -Force

Set-BrightcoveCredentials
Get-GoogleApi


$course_id = Read-Host "Enter Canvas Course ID or path to course HTML files"
$directory = $FALSE

if ($course_id -match "[A-Z]:\\") {
    $directory = $true
}
else {
    ."$home/Desktop/AccessibilityTools/CanvasReport-master/SetDomain.ps1"
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

if ($directory) {
    Search-Directory $course_id
}
else {
    Search-Course $course_id
}

Close-Chrome
#Excel formatting

try {
    #Importing it checks if it is empty
    Import-Excel $ExcelReport | Out-Null
    Write-Host 'Formatting Excel Document...' -ForegroundColor Green
    Format-MediaExcel1

    Get-MediaPivotTables

    Write-Host 'Finishing Excel formatting...' -ForegroundColor Green
    Format-MediaExcel2

    Write-Host "Report Generated" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Excel Sheet may be empty" -ForegroundColor Red
}

$sw.stop()
Send-Notification
