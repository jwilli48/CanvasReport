function Transpose-Data{
    param(
        [String[]]$Names,
        [Object[][]]$Data
    )
    for($i = 0;; ++$i){
        $Props = [ordered]@{}
        for($j = 0; $j -lt $Data.Length; ++$j){
            if($i -lt $Data[$j].Length){
                $Props.Add($Names[$j], $Data[$j][$i])
            }
        }
        if(!$Props.get_Count()){
            break
        }
        [PSCustomObject]$Props
    }
}

function Get-BrightcoveVideoLength {
  param(
    [string]$videoID
  )
  $chrome.url = ("https://studio.brightcove.com/products/videocloud/media/videos/search/" + $videoID)
  try{
    $length = (Wait-UntilElementIsVisible -Selector div[class*='runtime'] -byCssSelector).text
  }catch{
    Write-Host "Video not found"
    $length = "00:00"
  }
  $length = "00:" + $length
  $length = [TimeSpan]$length
  $length
}

function Get-GoogleVideoSeconds ([string]$VideoID){
 $gdata_uri = "https://www.googleapis.com/youtube/v3/videos?id=$VideoId&key=$GoogleApi&part=contentDetails"
 $metadata = irm $gdata_uri
 $duration = $metadata.items.contentDetails.duration;

 $ts = [Xml.XmlConvert]::ToTimeSpan("$duration")
 '{0:00},{1:00},{2:00}.{3:00}' -f ($ts.Hours+$ts.Days*24), $ts.Minutes, $ts.Seconds, $ts.Milliseconds | Out-Null

 $timespan = [TimeSpan]::Parse($ts)
 $totalSeconds = $timespan.TotalSeconds
 $totalSeconds
}

function Get-BYUMediaSiteVideoLength{
  param(
    [string]$videoID
  )
  $chrome.url = "https://byu.mediasite.com/Mediasite/Play/" + $videoID
  try{
    while($length -eq "0:00" -or $length -eq "" -or $length -eq $NULL){
      $length = (Wait-UntilElementIsVisible -Selector span[class*="duration"] -byCssSelector).text
    }
  }catch{
    Write-Host "Video not found"
    $length = "00:00"
  }
  $length = "00:" + $length
  $length = [TimeSpan]$length
  $length
}

function Get-TranscriptAvailable{
  param(
    [string]$iframe
  )
  $check = $page_body.split("`n")
  $i = 0
  while(-not $check[$i].contains($iframe)){$i++}
  if($check[$i+1].contains('Transcript')){
    return $true
  }elseif($check[$i-1].contains('Transcript')){
    return $true
  }else{
    return $false
  }

}
