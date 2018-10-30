function Format-TransposeData {
    param(
        [String[]]$Names,
        [Object[][]]$Data
    )
    for ($i = 0; ; ++$i) {
        $Props = [ordered]@{}
        for ($j = 0; $j -lt $Data.Length; ++$j) {
            if ($i -lt $Data[$j].Length) {
                $Props.Add($Names[$j], $Data[$j][$i])
            }
        }
        if (!$Props.get_Count()) {
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
    try {
        $length = (Wait-UntilElementIsVisible -Selector div[class*='runtime'] -byCssSelector).text
    }
    catch {
        try {
            $chrome.url = "https:" + ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value})
            (Wait-UntilElementIsVisible -Selector button.vjs-big-play-button -byCssSelector).click()
            $length = (Wait-UntilElementIsVisible -Selector div[class*="vjs-duration-display"] -byCssSelector).text.split("`n")[-1]
        }
        catch {
            Write-Host "Video not found" -ForegroundColor Magenta
            $length = "00:00"
            $Global:videoNotFound = "`nVideo not found"
        }
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-GoogleVideoSeconds ([string]$VideoID) {
    $gdata_uri = "https://www.googleapis.com/youtube/v3/videos?id=$VideoId&key=$GoogleApi&part=contentDetails"
    $metadata = Invoke-RestMethod $gdata_uri
    $duration = $metadata.items.contentDetails.duration;

    $ts = [Xml.XmlConvert]::ToTimeSpan("$duration")
    '{0:00},{1:00},{2:00}.{3:00}' -f ($ts.Hours + $ts.Days * 24), $ts.Minutes, $ts.Seconds, $ts.Milliseconds | Out-Null

    $timespan = [TimeSpan]::Parse($ts)
    $totalSeconds = $timespan.TotalSeconds
    $totalSeconds
}

function Get-BYUMediaSiteVideoLength {
    param(
        [string]$videoID
    )
    $chrome.url = "https://byu.mediasite.com/Mediasite/Play/" + $videoID
    try {
        while ($length -eq "0:00" -or $length -eq "" -or $length -eq $NULL) {
            $length = (Wait-UntilElementIsVisible -Selector span[class*="duration"] -byCssSelector).text
        }
    }
    catch {
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-PanoptoVideoLength {
    param(
        [string]$videoID
    )
    $chrome.url = "https://byu.hosted.panopto.com/Panopto/Pages/Embed.aspx?id=$videoID&amp;v=1"
    while ($chrome.ExecuteScript("return document.readyState") -ne "complete") {}
    while ($chrome.ExecuteScript("return jQuery.active") -ne 0) {}
    try {
        (Wait-UntilElementIsVisible -Selector 'div[aria-label="Play"]' -byCssSelector).click()
        $length = (Wait-UntilElementIsVisible -Selector span[class*="duration"] -byCssSelector).text
    }
    catch {
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-AlexanderStreetVideoLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://search.alexanderstreet.com/embed/token/$videoId"
    try{
        $length = (Wait-UntilElementIsVisible -ByCssSelector -Selector span.fulltime).text
    }
    catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-AlexanderStreenLinkLength {
     param(
        [string]$videoId
    )
    $chrome.url = "https://lib.byu.edu/remoteauth/?url=https://search.alexanderstreet.com/view/work/bibliog
raphic_entity|video_work|$videoId"
    try{
        $length = (Wait-UntilElementIsVisible -ByCssSelector -Selector span.fulltime).text
    }
    catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}
function Get-KanopyVideoLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://byu.kanopy.com/embed/$videoID"
    try{
        (Wait-UntilElementIsVisible -Selector button.vjs-big-play-button -byCssSelector).click()
        $length = ((Wait-UntilElementIsVisible -ByCssSelector -Selector div.vjs-remaining-time-display).text -split '-')[-1]
    }catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-KanopyLinkLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://byu.kanopy.com/video/$videoID"
    try{
        (Wait-UntilElementIsVisible -Selector button.vjs-big-play-button -byCssSelector).click()
        $length = ((Wait-UntilElementIsVisible -ByCssSelector -Selector div.vjs-remaining-time-display).text -split '-')[-1]
    }catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-AmbroseVideoLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://byu.kanopy.com/embed/$videoID"
    try{
        (Wait-UntilElementIsVisible -Selector div.jw-icon.jw-icon-display.jw-button-color.jw-reset -byCssSelector).click()
        $length = ((Wait-UntilElementIsVisible -ByCssSelector -Selector span.jw-text.jw-reset.jw-text-duration).text -split '-')[-1]
    }catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-FacebookVideoLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://www.facebook.com/video/embed?video_id=$videoID"
    try{
        (Wait-UntilElementIsVisible -Selector img -byCssSelector).click()
        $length = (Wait-UntilElementIsVisible -ByCssSelector -Selector div[playbackdurationtimestamp]).text -replace '-', '0'
    }catch{
        try{
            $chrome.Navigate().Refresh()
            (Wait-UntilElementIsVisible -Selector img -byCssSelector).click()
            $length = (Wait-UntilElementIsVisible -ByCssSelector -Selector div[playbackdurationtimestamp]).text -replace '-', '0'
        }catch{
            Write-Host "Video not found" -ForegroundColor Magenta
            $length = "00:00"
            $Global:videoNotFound = "`nVideo not found"
        }
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-DailyMotionVideoLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://www.dailymotion.com/embed/video/$videoID"
    try{
        (Wait-UntilElementIsVisible -Selector button[aria-label*="Playback"] -byCssSelector).click()
        $length = ((Wait-UntilElementIsVisible -ByCssSelector -Selector span[aria-label*="Duration"]).text)
    }catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-VimeoVideoLength {
    param(
        [string]$videoId
    )
    $chrome.url = "https://player.vimeo.com/video/$videoID"
    try{
        $length = ((Wait-UntilElementIsVisible -ByCssSelector -Selector div.timecode).text)
    }catch{
        Write-Host "Video not found" -ForegroundColor Magenta
        $length = "00:00"
        $Global:videoNotFound = "`nVideo not found"
    }
    $length = "00:" + $length
    $length = [TimeSpan]$length
    $length
}

function Get-TranscriptAvailable {
    param(
        [string]$iframe
    )
    $check = $page_body.split("`n")
    $i = 0
    while (-not $check[$i].contains($iframe)) {$i++}
    if ($NULL -ne $check[$i + 1]) {
        if ($check[$i + 1].contains('Transcript')) {
            return $true
        }
    }
    elseif ($check[$i - 1].contains('Transcript')) {
        return $true
    }
    elseif ($NULL -ne $check[$i + 2] -and $check[$i + 2].contains('Transcript')) {
        return $true
    }
    else {
        return $false
    }
}
