Set-ExecutionPolicy Bypass -Scope Process

Import-Module ./ProcessA11yReport.ps1 -Force
Import-Module ./PoshCanvasNew.ps1 -Force
Import-Module ./CheckModules.ps1 -Force
Import-Module ./FormatExcel.ps1 -Force
Import-Module ./Notifications.ps1 -Force
Import-Module ./SearchCourse.ps1 -Force


$course_id = Read-Host "Enter Canvas Course ID or path to course HTML files"
$directory = $FALSE

if($course_id -match "[A-Z]:\\"){
  $directory = $true
}else{
  ./SetDomain.ps1
}

$sw = [Diagnostics.Stopwatch]::new()
$sw.start()

Get-Modules
if($directory){
  Search-Directory $course_id
}else{
  Search-Course $course_id
}

#Excel formatting
Write-Host 'Formatting Excel Document...' -ForegroundColor Green
Format-A11yExcel

#Get-A11yPivotTables
ConvertTo-A11yExcel

Write-Host "Report Generated" -ForegroundColor Green
$sw.stop()
Send-Notification
