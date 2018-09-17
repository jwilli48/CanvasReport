Set-ExecutionPolicy Bypass -Scope Process

Import-Module ./PoshCanvasNew.ps1 -Force
Import-Module ./CheckModules.ps1 -Force
Import-Module ./FormatExcel.ps1 -Force
Import-Module ./Notifications.ps1 -Force
Import-Module ./Util.ps1 -Force

$course_id = Read-Host "Enter directory"

$sw = [Diagnostics.Stopwatch]::new()
$sw.start()

Get-Modules

$Global:location = @()
$Global:href = @()
$Global:status = @()

function Search-Directory{
  param(
    [string]$directory
  )

  $Global:courseName = $Directory.split('\')[-2]
  $course_files = Get-ChildItem "$directory\*.html" -Exclude '*old*','*ImageGallery*', '*CourseMedia*', '*GENERIC*'
  if($course_files -eq $NULL){
    Write-Host "ERROR: Directory input is empty"
  }else{
    $i = 0
    foreach($file in $course_files){
      $i++
      Write-Progress -Activity "Checking pages" -Status "Progress:" -PercentComplete ($i/$course_files.length * 100)

      $file_content = Get-Content -Encoding UTF8 -Path $file.PSpath -raw
      $item = Transpose-Data body, title $file_content, $file.name
      $page_body = $item.body
      Write-Host $item.title -ForegroundColor Green

      if($page_body -eq '' -or $page_body -eq $NULL){
        continue
      }
      Process-Links $page_body
    }
  }
}

function Process-Links{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | % {$_.Matches.Value}
  $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
  foreach($href in $href_list){
    if($href -notmatch "http" -and $href -notmatch "^www\."){
      if($href -match ".*?\.html" -or $href -match ".*?\.pdf" -or $href -match ".*?\.docx" -or $href -match ".*?\.xlsx"){
          if($href -match "^\.\."){
            if(-not (Test-Path (("$($course_id.split(`"\`").replace(`"HTML`",`"`") -join `"\`")$($href.split(`"/`").replace(`"..`",$NULL) -join `"\`")") -replace "\\\\", "\"))){
              AddToArray $item.title $href "File path doesn't exist"
            }
          }
          elseif(-not (Test-Path "$course_id\$href")){
            AddToArray $item.title $href "File doesn't exist"
          }
          continue
      }else{
        #not a link
        continue
      }
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
        AddToArray $item.title $href "Unkown error, check link"
      }
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

Search-Directory $course_id

$data = Transpose-Data Location, URL, Status $Global:location, $Global:href, $Global:status
$Global:ExcelReport = $PSScriptRoot + "\Reports\LinkCheck_" + $courseName + ".xlsx"
if(-not ($data -eq $NULL)){
  $data | Export-Excel $ExcelReport -AutoFilter -AutoSize -Append
}

$sw.stop()
Write-Host "Report Generated" -ForegroundColor Green
Send-Notification
