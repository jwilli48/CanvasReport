function Process_Contents {
    param(
        [string]$page_body
    )
    $Global:elementList = [System.Collections.ArrayList]::new()
    $Global:locationList = [System.Collections.ArrayList]::new()
    $Global:videoIDList = [System.Collections.ArrayList]::new()
    $Global:UrlList = [System.Collections.ArrayList]::new()
    $Global:videoLengthList = [System.Collections.ArrayList]::new()
    $Global:textList = [System.Collections.ArrayList]::new()
    $Global:transcriptAvailability = [System.Collections.ArrayList]::new()
    $Global:mediaCountList = [System.Collections.ArrayList]::new()

    $Global:ExcelReport = $PSScriptRoot + "\Reports\MediaReport_" + $courseName + "_$ReportType.xlsx"

    Start-ProcessLinks
    Start-ProcessIframes
    Start-ProcessBrightcoveVideoHTML
    Start-ProcessVideoTags

    $data = Format-TransposeData Element, Location, VideoID, Url, VideoLength, Text, Transcript, MediaCount $elementList, $locationList, $videoIDList, $UrlList, $videoLengthList, $textList, $transcriptAvailability, $mediaCountList
    $markRed = @((New-ConditionalText -Text "Video not found" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))
    $highlight = @((New-ConditionalText -Text "Duplicate Video" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ))
    $markBlue = @((New-ConditionalText -Text "Inline Media:`nUnable to find title or video length for this type of video" -BackgroundColor Cyan -ConditionalTextColor '#000000'))
    if (-not ($data -eq $NULL)) {
        $data | Export-Excel $ExcelReport -ConditionalText $highlight, $markRed, $markBlue -AutoFilter -AutoSize -Append
    }
}

function Start-ProcessLinks {
    $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | ForEach-Object {$_.Matches.Value}
    $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
    foreach ($link in $link_list) {
        switch -regex ($link) {
            "class=.*?video_link" {
                $transcript = Get-TranscriptAvailable $link
                if ($transcript) {$transcript = "Yes"}
                else {$transcript = "No"}
                AddToArray "Canvas Video Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" "00:00:00" "Inline Media:`nUnable to find title or video length for this type of video" $transcript
                break
            }
        }
    }
    foreach ($href in $href_list) {
        $Global:videoNotFound = ""
        switch -regex ($href) {
            "youtu\.?be" {
                if ($href.contains("t=")) {
                    $href.split("?")[0].split("/")[-1]
                }
                elseif ($href.contains('=')) {
                    $VideoID = ($href -split 'v=')[-1].split("&")[0]
                }
                else {
                    $VideoID = $href.split('/')[-1]
                }
                try {
                    $VideoID = $VideoID.split("?")[0]
                    $video_Length = [timespan]::fromseconds((Get-GoogleVideoSeconds -VideoID $VideoID)).toString("hh\:mm\:ss")
                }
                catch {
                    Write-Host "Video not found" -ForegroundColor Magenta
                    $video_length = "00:00:00"
                    $Global:videoNotFound = "`nVideo not found"
                }
                AddToArray "Youtube Link" "$($item.url -split `"api/v\d/`" -join `"`")" $VideoID $video_Length "$href$videoNotFound" "Yes" $href
                break
            }
            Default {}
        }
    }
}

function Start-ProcessIframes {
    $iframeList = $page_body | Select-String -pattern "<iframe.*?>.*?</iframe>" -AllMatches | ForEach-Object {$_.Matches.Value}
    foreach ($iframe in $iframeList) {
        $Global:videoNotFound = ""
        $title = ""
        if (-not $iframe.contains('title')) {
            $title = "No Title Attribute Found"
        }
        else {
            $title = $iframe | Select-String -pattern 'title="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value}
            if ($title -eq $NULL -or $title -eq "") {
                $title = "Title found but was empty or could not be saved."
            }
        }
        $url = $iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value}
        if ($iframe.contains('youtube')) {
            $VideoLink = $iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value}
            if ($VideoLink.contains('?')) {
                $Video_ID = $VideoLink.split('/')[4].split('?')[0]
            }
            else {
                $Video_ID = $VideoLink.split('/')[-1]
            }
            try {
                $Video_ID = $Video_ID.split('?')[0]
                $video_Length = [timespan]::fromseconds((Get-GoogleVideoSeconds -VideoID $Video_ID)).toString("hh\:mm\:ss")
            }
            catch {
                Write-Host "Video not found" -ForegroundColor Magenta
                $video_length = "00:00:00"
                $Global:videoNotFound = "`nVideo not found"
            }
            AddToArray "Youtube Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $video_Length "$title$videoNotFound" "Yes" $Url
        }
        elseif ($iframe.contains('brightcove')) {
            $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value}).split('=')[-1].split("&")[0]
            $video_Length = (Get-BrightcoveVideoLength $Video_ID).toString('hh\:mm\:ss')
            $transcript = Get-TranscriptAvailable $iframe
            if ($transcript) {$transcript = "Yes"}
            else {$transcript = "No"}
            AddToArray "Brightcove Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $video_Length "$title$videoNotFound" $transcript $url
        }
        elseif ($iframe.contains('H5P')) {
            AddToArray "H5P" "$($item.url -split `"api/v\d/`" -join `"`")" "" "00:00:00" $title "N\A" $url
        }
        elseif ($iframe.contains('byu.mediasite')) {
            $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].Value}).split('/')[-1]
            if ($video_ID -eq "") {
                $video_id = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].Value}).split('/')[-2]
            }
            $video_Length = (Get-BYUMediaSiteVideoLength $Video_ID).toString('hh\:mm\:ss')
            $transcript = Get-TranscriptAvailable $iframe
            if ($transcript) {$transcript = "Yes"}
            else {$transcript = "No"}
            AddToArray "BYU Mediasite Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $video_Length "$title$videoNotFound" $transcript $url
        }
        elseif ($iframe.contains('Panopto')) {
            $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].Value}).split('=').split('&')[1]
            $video_Length = (Get-PanoptoVideoLength $video_ID).toString('hh\:mm\:ss')
            $transcript = Get-TranscriptAvailable $iframe
            if ($transcript) {$transcript = "Yes"}
            else {$transcript = "No"}
            AddToArray "Panopto Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $video_Length "$title$videoNotFound" $transcript $url
        }
        else {
            AddToArray "Iframe" "$($item.url -split `"api/v\d/`" -join `"`")" "" "00:00:00" $title "N\A" $url
        }
    }
}

function Start-ProcessVideoTags {
    $videotag_list = $page_body | Select-String -pattern '<video.*?>.*?</video>' -AllMatches | ForEach-Object {$_.Matches.Value}
    foreach ($video in $videotag_list) {
        $src = $video | Select-String -pattern 'src="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
        $videoID = $src.split('=')[1].split("&")[0]
        $transcript = Get-TranscriptAvailable $video
        if ($transcript) {$transcript = "Yes"}
        else {$transcript = "No"}
        AddToArray "Inline Media Video" "$($item.url -split `"api/v\d/`" -join `"`")" $videoID "00:00:00" "Inline Media:`nUnable to find title or video length for this type of video" $transcript $src
    }
}

function Start-ProcessBrightcoveVideoHTML {
    $brightcove_list = $page_body | Select-String -pattern 'id="[^\d]*(\d{13}).*?"' -Allmatches | ForEach-Object {$_.Matches.Value}
    $id_list = $brightcove_list | Select-String -pattern '\d{13}' -AllMatches | ForEach-Object {$_.matches.Value}
    foreach ($id in $id_list) {
        $video_Length = (Get-BrightcoveVideoLength $id).toString('hh\:mm\:ss')
        $transcriptCheck = $page_body.split("`n")
        $i = 0
        while ($transcriptCheck[$i] -notmatch "$id") {$i++}
        $transcript = $FALSE
        for ($j = $i; $j -lt ($i + 10); $j++) {
            if ($transcriptCheck[$j] -eq $NULL) {
                #End of file
                break
            }
            elseif ($transcriptCheck[$j] -match "transcript") {
                $transcript = $TRUE
                break
            }
        }
        if ($transcript) {$transcript = "Yes"}
        else {$transcript = "No"}
        AddToArray "Brightcove Video" "$($item.url -split `"api/v\d/`" -join `"`")" $id $video_Length $title $transcript "No URL for this type"
    }
}

function AddToArray {
    param(
        [string]$element,
        [string]$location,
        [string]$VideoID,
        [TimeSpan]$VideoLength,
        [string]$Text,
        [string]$Transcript,
        [string]$url,
        [string]$MediaCount = 1
    )
    $excel = Export-Excel $ExcelReport -PassThru

    $Global:elementList += $element
    $Global:locationList += $location
    if ($excel -eq $NULL) {}
    else {
        if ($videoID -eq "" -or $videoID -eq $NULL) {
            $Global:videoIDList += $VideoID
            $Global:videoLengthList += $VideoLength
        }
        elseif ($excel.Workbook.Worksheets['Sheet1'].Names.Value -contains $videoID) {
            $Global:videoIDList += "Duplicate Video: `n$videoID"
            $Global:videoLengthList += ""
        }
        else {
            $Global:videoIDList += $VideoID
            $Global:videoLengthList += $VideoLength
        }
    }
    $Global:textList += $Text
    $Global:transcriptAvailability += $Transcript
    $Global:mediaCountList += $MediaCount
    $Global:UrlList += $url
    $excel.Dispose()
}

function Get-GoogleAPI {
    if (-not (Test-Path "$PSScriptRoot\Passwords\MyGoogleApi.txt")) {
        Write-Host "Google API needed to get length of YouTube videos." -ForegroundColor Magenta
        $api = Read-Host "Please enter it now (It will then be saved)"
        Set-Content $PSScriptRoot\Passwords\MyGoogleApi.txt $api
    }
    $Global:GoogleApi = Get-Content "$PSScriptRoot\Passwords\MyGoogleApi.txt"
}
