Set-ExecutionPolicy Bypass -Scope Process

Import-Module ./ProcessA11yReport.ps1 -Force
Import-Module ./PoshCanvasNew.ps1 -Force
Import-Module ./CheckModules.ps1 -Force
Import-Module ./FormatExcel.ps1 -Force
Import-Module ./Notifications.ps1 -Force
Import-Module ./SearchCourse.ps1 -Force

$course_id = Read-Host "Enter Course ID"

$sw = [Diagnostics.Stopwatch]::new()
$sw.start()

Get-Modules
Search-Course $course_id

#Excel formatting
Write-Host 'Formatting Excel Document...' -ForegroundColor Green
Format-A11yExcel

#Get-A11yPivotTables
ConvertTo-A11yExcel

Write-Host "Report Generated" -ForegroundColor Green
$sw.stop()
Send-Notification
