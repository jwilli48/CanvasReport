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
  if(-not (Test-Path "$PSScriptRoot\Passwords\MyGoogleApi.txt"){
    $api = Read-Host "Google API needed to get length of YouTube videos. Please enter it now (It will then be saved)"
    Set-Content $PSScriptRoot\Passwords\MyGoogleApi.txt $api
  }
 $gdata_uri = "https://www.googleapis.com/youtube/v3/videos?id=$VideoId&key=$GoogleApi&part=contentDetails"
 $metadata = irm $gdata_uri
 $duration = $metadata.items.contentDetails.duration;

 $ts = [Xml.XmlConvert]::ToTimeSpan("$duration")
 '{0:00},{1:00},{2:00}.{3:00}' -f ($ts.Hours+$ts.Days*24), $ts.Minutes, $ts.Seconds, $ts.Milliseconds | Out-Null

 $timespan = [TimeSpan]::Parse($ts)
 $totalSeconds = $timespan.TotalSeconds
 $totalSeconds
}
