Import-Module ./Util.ps1 -Force

function Process_Contents{
  param(
    [string]$page_body
  )
  $Global:elementList = @()
  $Global:locationList = @()
  $Global:videoIDList = @()
  $Global:videoLengthList = @()
  $Global:textList = @()
  $Global:transcriptAvailability = @()
  $Global:mediaCountList = @()

  $Global:ExcelReport = $PSScriptRoot + "\Reports\MediaReport_" + $courseName + ".xlsx"

  Process-Links
  Process-Iframes

  $data = Transpose-Data Element, Location, VideoID, VideoLength, Text, Transcript, MediaCount $elementList, $locationList, $videoIDList, $videoLengthList, $textList, $transcriptAvailability, $mediaCountList
  $markRed = @((New-ConditionalText -Text "No Title" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))
  $highlight = @((New-ConditionalText -Text "Duplicate Video" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ))
  if(-not ($data -eq $NULL)){
    $data | Export-Excel $ExcelReport -ConditionalText $highlight -AutoFilter -AutoSize -Append
  }
}

function Process-Links{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | % {$_.Matches.Value}
  $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
  for($i = 0; $i -lt $href_list.length; $i++){
    switch -regex ($href_list[$i])
    {
      "youtu\.?be" {
        if($href_list[$i].contains('=')){
          $VideoID = $href_list[$i].split('=')[-1]
        }else{
          $VideoID = $href_list[$i].split('/')[-1]
        }
        try{
          $video_Length = [timespan]::fromseconds((Get-GoogleVideoSeconds -VideoID $VideoID)).toString("hh\:mm\:ss")
        }catch{
          Write-Host "Video not found"
          $video_length = "00:00:00"
        }
        AddToArray "Youtube Link" $page.title $VideoID $video_Length $href_list[$i] "Yes"; break
      }
      Default {}
    }
  }
}

function Process-Iframes{
  $iframeList = $page_body | Select-String -pattern "<iframe.*?>.*?</iframe>" -AllMatches | % {$_.Matches.Value}
  foreach($iframe in $iframeList){
    $title = ""
    if(-not $iframe.contains('title')){
      $title = "No Title Found"
    }else{
      $title = $iframe | Select-String -pattern 'title="(.*?)"' | % {$_.Matches.Groups[1].value}
      if($title -eq $NULL -or $title -eq ""){
        $title = "Title found but was empty or could not be saved."
      }
    }

    if($iframe.contains('youtube')){
      $VideoLink = $iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}
      if($VideoLink.contains('?')){
        $Video_ID = $VideoLink.split('/')[4].split('?')[0]
      }else{
        $Video_ID = $VideoLink.split('/')[-1]
      }
      try{
        $video_Length = [timespan]::fromseconds((Get-GoogleVideoSeconds -VideoID $Video_ID)).toString("hh\:mm\:ss")
      }catch{
        Write-Host "Video not found"
        $video_length = "00:00:00"
      }
      AddToArray "Youtube Video" $page.title $video_ID $video_Length $title "Yes"
    }
    elseif($iframe.contains('brightcove')){
      $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('=')[-1]
      $video_Length = (Get-BrightcoveVideoLength $Video_ID).toString('hh\:mm\:ss')
      $transcript = Get-TranscriptAvailable $iframe
      if($transcript){$transcript = "Yes"}
      else{$transcript = "No"}
      AddToArray "Brightcove Video" $page.title $video_ID $video_Length $title $transcript
    }
    elseif($iframe.contains('H5P')){
      AddToArray "H5P" $page.title "" "00:00:00" $title "N\A"
    }
    elseif($iframe.contains('byu.mediasite')){
      $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('/')[-1]
      if($video_ID -eq ""){
        $video_id = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('/')[-2]
      }
      $video_Length = (Get-BYUMediaSiteVideoLength $Video_ID).toString('hh\:mm\:ss')
      $transcript = Get-TranscriptAvailable $iframe
      if($transcript){$transcript = "Yes"}
      else{$transcript = "No"}
      AddToArray "Byu Mediasite Video" $page.title $video_ID $video_Length $title $transcript
    }
    else{
      AddToArray "Iframe" $page.title "" "00:00:00" $title "N\A"
    }
  }
}

function AddToArray{
  param(
    [string]$element,
    [string]$location,
    [string]$VideoID,
    [TimeSpan]$VideoLength,
    [string]$Text,
    [string]$Transcript,
    [string]$MediaCount = 1
  )
  $excel = Export-Excel $ExcelReport -PassThru

  $Global:elementList += $element
  $Global:locationList += $location
  if($excel -eq $NULL){}
  else{
    if($videoID -eq "" -or $videoID -eq $NULL){
      $Global:videoIDList += $VideoID
      $Global:videoLengthList += $VideoLength
    }
    elseif($excel.Workbook.Worksheets['Sheet1'].Names.Value -contains $videoID){
      $Global:videoIDList += "Duplicate Video: `n$videoID"
      $Global:videoLengthList += ""
    }else{
      $Global:videoIDList += $VideoID
      $Global:videoLengthList += $VideoLength
    }
  }
  $Global:textList += $Text
  $Global:transcriptAvailability += $Transcript
  $Global:mediaCountList += $MediaCount
  $excel.Dispose()
}

function Get-GoogleAPI{
  if(-not (Test-Path "$PSScriptRoot\Passwords\MyGoogleApi.txt")){
    Write-Host "Google API needed to get length of YouTube videos." -ForegroundColor Yellow
    $api = Read-Host "Please enter it now (It will then be saved)"
    Set-Content $PSScriptRoot\Passwords\MyGoogleApi.txt $api
  }
  $Global:GoogleApi = Get-Content "$PSScriptRoot\Passwords\MyGoogleApi.txt"
}
