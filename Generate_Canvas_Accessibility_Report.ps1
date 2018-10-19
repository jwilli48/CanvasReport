Set-ExecutionPolicy Bypass -Scope Process

Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/ProcessA11yReport.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/PoshCanvasNew.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/CheckModules.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/FormatExcel.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/Notifications.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/SearchCourse.ps1" -Force


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
if ($directory) {
    Search-Directory $course_id
}
else {
    Search-Course $course_id
}

#Excel formatting
Write-Host 'Formatting Excel Document...' -ForegroundColor Green
Format-A11yExcel

#Get-A11yPivotTables
ConvertTo-A11yExcel
#Add-LocationLinks

Write-Host "Report Generated" -ForegroundColor Green
$sw.stop()
Send-Notification
