Set-ExecutionPolicy Bypass -Scope Process

Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/PoshCanvasNew.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/CheckModules.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/FormatExcel.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/Notifications.ps1" -Force
Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/Util.ps1" -Force

$course_id = Read-Host "Enter directory"

$sw = [Diagnostics.Stopwatch]::new()
$sw.start()

Get-Modules

$Global:location = [System.Collections.ArrayList]::new()
$Global:href = [System.Collections.ArrayList]::new()
$Global:status = [System.Collections.ArrayList]::new()

function Search-Directory{
  param(
    [string]$directory
  )

  $Global:courseName = $Directory.split('\')[-2]
  $course_files = Get-ChildItem "$directory\*.html" -Exclude '*old*','*ImageGallery*', '*CourseMedia*', '*GENERIC*'
  if($NULL -eq $course_files){
    Write-Host "ERROR: Directory input is empty"
  }else{
    $i = 0
    foreach($file in $course_files){
      $i++
      Write-Progress -Activity "Checking pages" -Status "Progress:" -PercentComplete ($i/$course_files.length * 100)

      $file_content = Get-Content -Encoding UTF8 -Path $file.PSpath -raw
      $item = Format-TransposeData body, title $file_content, $file.name
      $page_body = $item.body
      Write-Host $item.title -ForegroundColor Green

      if($page_body -eq '' -or $NULL -eq $page_body){
        continue
      }
      Start-ProcessLinks $page_body
    }
  }
}

function Start-ProcessLinks{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | ForEach-Object {$_.Matches.Value}
  $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
  foreach($href in $href_list){
    if($href -match "^#" -or $href -match "^mailto:"){
      #these can't be checked by the program
      continue
    }
    if($href -notmatch "http" -and $href -notmatch "^www\." -and $href -notmatch ".*?\.com$" -and $href -notmatch ".*?\.org$"){
        if($href -match "^\.\."){
          if(-not (Test-Path (("$($course_id.split(`"\`").replace(`"HTML`",`"`") -join `"\`")$($href.split(`"/`").replace(`"..`",$NULL) -join `"\`")") -replace "\\\\", "\"))){
            AddToArray $item.title $href "File path doesn't exist"
          }
        }
        elseif(-not (Test-Path "$course_id\$href")){
          AddToArray $item.title $href "File doesn't exist"
        }
        continue
    }
    try{
      Invoke-WebRequest $href | Out-Null
    }catch{
      #There is also a SecureChannelFailure error status but it seems that for the majority of cases those still work when clicking the link manually
      if($_.Exception.Status -eq "SendFailure"){
        AddToArray $item.title $href "Broken link"
      }
      elseif($_.Exception.Status -eq "SecureChannelFailure"){
        AddToArray $item.title $href "Check link"
      }else{
        AddToArray $item.title $href "Unknown error, check link"
      }
    }
  }

  #Check images to see if their source exists
  $image_list = $page_body | Select-String -pattern "<img.*?>" -AllMatches | ForEach-Object {$_.Matches.Value}
  $src_list = $image_list | Select-String -pattern 'src="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
  foreach($src in $src_list){
        if ($src -notmatch "http" -and $src -notmatch "^www\." -and $src -notmatch ".*?\.com$" -and $src -notmatch ".*?\.org$") {
            if ($src -match "^\.\.") {
                if (-not (Test-Path (("$($course_id.split(`"\`").replace(`"HTML`",`"`") -join `"\`")$($src.split(`"/`").replace(`"..`",$NULL) -join `"\`")") -replace "\\\\", "\"))) {
                    AddToArray $item.title $src "Image file path doesn't exist"
                }
            }
            elseif (-not (Test-Path "$course_id\$src")) {
                AddToArray $item.title $src "Image file doesn't exist"
            }
            continue
        }
  }
}

function AddToArray{
  param(
    [string]$locationIn,
    [string]$hrefIn,
    [string]$statusIn
  )
  $Global:location += $locationIn
  $Global:href += $hrefIn
  $Global:status += $statusIn

}
#Maybe this will get rid of the SecureChannelFailure
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Search-Directory $course_id

$data = Format-TransposeData Location, URL, Status $Global:location, $Global:href, $Global:status
$Global:ExcelReport = $PSScriptRoot + "\Reports\LinkCheck_" + $courseName + ".xlsx"
if(-not ($NULL -eq $data)){
  $data | Export-Excel $ExcelReport -AutoFilter -AutoSize -Append
  Write-Host "Report saved to $ExcelReport"
}else{
  Write-Host "Nothing found" -ForegroundColor Green
}

$sw.stop()
Write-Host "Report Generated" -ForegroundColor Green
Send-Notification
