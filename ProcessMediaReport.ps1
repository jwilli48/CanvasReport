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
  Process-BrightcoveVideoHTML
  Process-VideoTags

  $data = Transpose-Data Element, Location, VideoID, VideoLength, Text, Transcript, MediaCount $elementList, $locationList, $videoIDList, $videoLengthList, $textList, $transcriptAvailability, $mediaCountList
  $markRed = @((New-ConditionalText -Text "Video not found" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))
  $highlight = @((New-ConditionalText -Text "Duplicate Video" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ))
  if(-not ($data -eq $NULL)){
    $data | Export-Excel $ExcelReport -ConditionalText $highlight, $markRed -AutoFilter -AutoSize -Append
  }
}

function Process-Links{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | % {$_.Matches.Value}
  $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
  foreach($href in $href_list){
    $Global:videoNotFound = ""
    switch -regex ($href)
    {
      "youtu\.?be" {
        if($href.contains('=')){
          $href = $href.split("&")[0]
          $VideoID = $href.split('=')[-1]
        }else{
          $VideoID = $href.split('/')[-1]
        }
        try{
          $video_Length = [timespan]::fromseconds((Get-GoogleVideoSeconds -VideoID $VideoID)).toString("hh\:mm\:ss")
        }catch{
          Write-Host "Video not found" -ForegroundColor Magenta
          $video_length = "00:00:00"
          $Global:videoNotFound = "`nVideo not found"
        }
        AddToArray "Youtube Link" $item.title $VideoID $video_Length "$href$videoNotFound" "Yes"; break
      }
      Default {}
    }
  }
}

function Process-Iframes{
  $iframeList = $page_body | Select-String -pattern "<iframe.*?>.*?</iframe>" -AllMatches | % {$_.Matches.Value}
  foreach($iframe in $iframeList){
    $Global:videoNotFound = ""
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
        Write-Host "Video not found" -ForegroundColor Magenta
        $video_length = "00:00:00"
        $Global:videoNotFound = "`nVideo not found"
      }
      AddToArray "Youtube Video" $item.title $video_ID $video_Length "$title$videoNotFound" "Yes"
    }
    elseif($iframe.contains('brightcove')){
      $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('=')[-1].split("&")[0]
      $video_Length = (Get-BrightcoveVideoLength $Video_ID).toString('hh\:mm\:ss')
      $transcript = Get-TranscriptAvailable $iframe
      if($transcript){$transcript = "Yes"}
      else{$transcript = "No"}
      AddToArray "Brightcove Video" $item.title $video_ID $video_Length "$title$videoNotFound" $transcript
    }
    elseif($iframe.contains('H5P')){
      AddToArray "H5P" $item.title "" "00:00:00" $title "N\A"
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
      AddToArray "BYU Mediasite Video" $item.title $video_ID $video_Length "$title$videoNotFound" $transcript
    }
    elseif($iframe.contains('Panopto')){
      $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('=').split('&')[1]
      $video_Length = (Get-PanoptoVideoLength $video_ID).toString('hh\:mm\:ss')
      $transcript = Get-TranscriptAvailable $iframe
      if($transcript){$transcript = "Yes"}
      else{$transcript = "No"}
      AddToArray "Panopto Video" $item.title $video_ID $video_Length "$title$videoNotFound" $transcript
    }
    else{
      AddToArray "Iframe" $item.title "" "00:00:00" $title "N\A"
    }
  }
}

function Process-VideoTags{
    $videotag_list = $page_body | Select-String -pattern '<video.*?>.*?</video>' -AllMatches | %{$_.Matches.Value}
    foreach($video in $videotag_list){
      $src = $video | Select-String -pattern 'src="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
      $videoID = $src.split('=')[1].split("&")[0]
      $transcript = Get-TranscriptAvailable $video
      if($transcript){$transcript = "Yes"}
      else{$transcript = "No"}
      AddToArray "Inline Media Video" $item.title $videoID "00:00:00" "Inline Media:`nUnable to find title or video length for this type of video" $transcript
    }
}

function Process-BrightcoveVideoHTML{
  $brightcove_list = $page_body | Select-String -pattern '<div id="[^\d]*(\d{13})"' -Allmatches | % {$_.Matches.Value}
  $id_list = $brightcove_list | Select-String -pattern '\d{13}' -AllMatches | % {$_.matches.Value}
  foreach($id in $id_list){
    $video_Length = (Get-BrightcoveVideoLength $id).toString('hh\:mm\:ss')
    $transcriptCheck = $page_body.split("`n")
    $i = 0
    while($transcriptCheck[$i] -notmatch "$id"){$i++}
    $transcript = $FALSE
    for($j = $i; $j -lt ($i+5); $j++){
      if($transcript[$j] -eq $NULL){
        #End of file
        break
      }elseif($transcript[$j].contains("transcript")){
        $transcript = $TRUE
        break
      }
    }
    if($transcript){$transcript = "Yes"}
    else{$transcript = "No"}
    AddToArray "Brightcove Video" $item.title $id $video_Length $title $transcript
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
