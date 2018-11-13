function Process_Contents {
    param(
        [string]$page_body
    )
    $Global:elementList = [System.Collections.ArrayList]::new()
    $Global:locationList = [System.Collections.ArrayList]::new()
    $Global:videoIDList = [System.Collections.ArrayList]::new()
    $Global:textList = [System.Collections.ArrayList]::new()
    $Global:AccessibilityList = [System.Collections.ArrayList]::new()
    $Global:issueSeverityList = [System.Collections.ArrayList]::new()

    Start-ProcessLinks
    Start-ProcessIframes
    Start-ProcessImages
    Start-ProcessHeaders
    Start-ProcessTables
    Start-ProcessSemantics
    Start-ProcessVideoTags
    Start-ProcessBrightcoveVideoHTML
    Start-ProcessFlash
    Start-ProcessColor

    $data = Format-TransposeData Element, Location, VideoID, Text, Accessibility, IssueSeverity $elementList, $locationList, $videoIDList, $textList, $AccessibilityList, $issueSeverityList
    $Global:ExcelReport = $PSScriptRoot + "\Reports\A11yReport_" + $courseName + "_$ReportType.xlsx"

    ###ConditionalText no longer used, changed to using a Template instead###
    #$markRed  = @((New-ConditionalText -Text "Adjust Link Text" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Needs a title" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "No Alt Attribute" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Alt Text May Need Adjustment" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "JavaScript links are not accessible" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Check if header" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Broken Link" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "Empty link tag" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "No transcript found" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "Revise table" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))

    #$hightlight = @((New-ConditionalText -Text "<i> tags should be <em> tags" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ),(New-ConditionalText -Text "<b> tags should be <strong> tags" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ))

    if (-not ($NULL -eq $data)) {
        $data | Export-Excel $ExcelReport <#-ConditionalText $markRed, $highlight#> -AutoFilter -AutoSize -Append
    }
}

function Start-ProcessLinks {
    $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | ForEach-Object {$_.Matches.Value}
    foreach ($link in $link_list) {
        if ($link -match 'onclick') {
            AddToArray "JavaScript Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $link "JavaScript links are not accessible"
        }elseif (-not $link.contains('href')) {
            AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $link "Empty link tag"
        }
        elseif ($link -match 'href="\s*?"') {
            AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $link "Empty link tag"
        }
    }

    #$href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
  
    $link_text = $link_list | Select-String -pattern '<a.*?>(.*?)</a>' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
    foreach ($text in $link_text) {
        switch -regex ($text) {
            '<img' {
                #It will be caught by Proccess-Image
                break
            }
            $NULL {
                AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Invisble link with no text" "Adjust Link Text"; break
            }
            "^ ?[A-Za-z\.]+ ?$" {
                #This matches if the link text is a sigle word
                AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $text "Adjust Link Text"; break
            }
            "Click" {
                AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $text "Adjust Link Text"; break
            }
            "http" {
                AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $text "Adjust Link Text"; break
            }
            "https" {
                AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $text "Adjust Link Text"; break
            }
            "www\." {
                AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $text "Adjust Link Text"; break
            }
            "Link" {
                if (-not ($text -match "Links to an external site")) {
                    AddToArray "Link" "$($item.url -split `"api/v\d/`" -join `"`")" "" $text "Adjust Link Text"; break
                }
            }
            Default {}
        }
    }
}

function Start-ProcessImages {
    $image_list = $page_body | Select-String -pattern '<img.*?>' -AllMatches | ForEach-Object {$_.Matches.Value}
    foreach ($img in $image_list) {
        $alt = ""
        if (-not $img.contains('alt')) {
            $Accessibility = "No Alt Attribute"
            AddToArray "Image" "$($item.url -split `"api/v\d/`" -join `"`")" "" $img $Accessibility
        }
        else {
            $alt = $img | Select-String -pattern 'alt="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
            $Accessibility = "Alt Text May Need Adjustment"
            switch -regex ($alt) {
                "banner" {
                    AddToArray "Image" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Alt text:`n$alt" $Accessibility
                    Break
                }
                "Placeholder" {
                    AddToArray "Image" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Alt text:`n$alt" $Accessibility
                    Break
                }
                "\.jpg" {
                    AddToArray "Image" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Alt text:`n$alt" $Accessibility
                    Break
                }
                "\.png" {
                    AddToArray "Image" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Alt text:`n$alt" $Accessibility
                    Break
                }
                "https" {
                    AddToArray "Image" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Alt text:`n$alt" $Accessibility
                    Break
                }
                Default {}
            }
        }
    }

}

function Start-ProcessIframes {
    $iframeList = $page_body | Select-String -pattern "<iframe.*?>.*?</iframe>" -AllMatches | ForEach-Object {$_.Matches.Value}
    foreach ($iframe in $iframeList) {
        $title = ""
        if (-not $iframe.contains('title')) {
            $Accessibility = "Needs a title attribute"

            if ($iframe.contains('youtube')) {
                $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value}).split('/')[4].split('?')[0]
                AddToArray "Youtube Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $title $Accessibility
            }
            elseif ($iframe.contains('brightcove')) {
                $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].value}).split('=')[-1].split("&")[0]
                AddToArray "Brightcove Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $title $Accessibility
            }
            elseif ($iframe.contains('H5P')) {
                AddToArray "H5P" "$($item.url -split `"api/v\d/`" -join `"`")" "" $title $Accessibility
            }
            elseif ($iframe.contains('byu.mediasite')) {
                $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].Value}).split('/')[-1]
                if ($video_ID -eq "") {
                    $video_id = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].Value}).split('/')[-2]
                }
                AddToArray "BYU Mediasite Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $title $Accessibility
            }
            elseif ($iframe.contains('Panopto')) {
                $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | ForEach-Object {$_.Matches.Groups[1].Value}).split('=').split('&')[1]
                AddToArray "Panopto Video" "$($item.url -split `"api/v\d/`" -join `"`")" $video_ID $title $Accessibility
            }
            else {
                AddToArray "Iframe" "$($item.url -split `"api/v\d/`" -join `"`")" "" $title $Accessibility
            }
        }
    }
    #Check for transcripts
    $i = 1;
    foreach ($iframe in $iframeList) {
        if ($iframe.contains('brightcove') -or $iframe.contains('byu.mediasite') -or $iframe.contains('Panopto')) {
            if (-not (Get-TranscriptAvailable $iframe)) {
                AddToArray "Transcript" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Video number $i on page" "No transcript found"
            }
            $i++
        }
    }
}

function Start-ProcessBrightcoveVideoHTML {
    $brightcove_list = $page_body | Select-String -pattern '<div id="[^\d]*(\d{13})"' -Allmatches | ForEach-Object {$_.Matches.Value}
    $id_list = $brightcove_list | Select-String -pattern '\d{13}' -AllMatches | ForEach-Object {$_.matches.Value}
    foreach ($id in $id_list) {
        $transcriptCheck = $page_body.split("`n")
        $i = 0
        while ($transcriptCheck[$i] -notmatch "$id") {$i++}
        $transcript = $FALSE
        for ($j = $i; $j -lt ($i + 10); $j++) {
            if ($NULL -eq $transcriptCheck[$j]) {
                #End of file
                break
            }
            elseif ($transcriptCheck[$j] -match "transcript") {
                $transcript = $TRUE
                break
            }
        }
        if ($transcript) {$transcript = "Yes"}
        else {
            $transcript = "No"
            AddToArray "Transcript" "$($item.url -split `"api/v\d/`" -join `"`")" "$id" "No transcript found for BrightCove video with id:`n$id" "No transcript found"
        }
    }
}

function Start-ProcessHeaders {
    $headerList = $page_body | Select-String -pattern '<h\d.*?>.*?</h\d>' -Allmatches | ForEach-Object {$_.Matches.Value}
    $accessibility = ""
    foreach ($header in $headerList) {
        $headerLevel = $header | Select-String -Pattern "<h(\d)" -Allmatches | ForEach-Object {$_.matches.Groups[1].Value}
        #$headerText = $header | Select-String -pattern '<h\d.*?>(.*?)</h\d>' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
        switch -regex ($header) {
            'class=".*?screenreader-only.*?"' {
                $accessibility = "Check if header is meant to be invisible and is not a duplicate"
                AddToArray "Header Level $headerLevel" "$($item.url -split `"api/v\d/`" -join `"`")" "" $header $Accessibility
                break
            }
        }
    }
}

function Start-ProcessTables {
    if ($page_body.contains("<table")) {
        $tableNumber = 0
        $check = $page_body.split("`n")
        for ($i = 0; $i -lt $check.length; $i++) {
            $issueList = [System.Collections.ArrayList]::new()
            if ($check[$i].contains("<table")) {
                $rowNumber = 0
                $columnNumber = 0
                $tableNumber++
                $hasHeaders = $FALSE
                #Starts going through the whole table line by line
                try {
                    #Try will catch if the table is missing a </table> closing tag
                    while (-not ($check[$i].contains("</table>"))) {
                        #If table contains an heading tags it is an accessibility issue
                        if ($check[$i] -match "<h\d") {
                            $issueList += "Heading tags should not be inside of tables"
                        }

                        if ($check[$i] -match "colspan") {
                            if ($check[$i - 1] -match "<tr" -and $check[$i + 1] -match "</tr") {
                                $issueList += "Stretched cell(s) should possibly be a <caption> title for the table"
                            }
                        }
                        elseif ($check[$i] -match "<th[^e]") {
                            $hasHeaders = $TRUE
                            if ($check[$i] -notmatch "scope") {
                                $issueList += "Table headers should have either scope=`"row`" or scope=`"col`" for screenreaders"
                            }
                            $columnNumber++
                        }
                        elseif ($check[$i] -match "<td") {
                            if ($check[$i] -match "scope") {
                                $issueList += "Non-header table cells should not have scope attributes"
                            }
                            $columnNumber++
                        }
                        elseif ($check[$i] -match "<tr") {
                            $rowNumber++
                        }
                        elseif ($check[$i] -match "<th[^e]" -or $check[$i] -match "<td") {
                            $columnNumber++
                        }
                        elseif ($check[$i] -match "</tr>") {
                            if ($check[$i + 1] -match "<tr") {
                                $columnNumber = 0
                            }
                        }
                        $i++
                    }
                }
                catch {
                    $issueList += "Table does not have an ending </table> tag"
                }
                if (-not $hasHeaders) {
                    if (($rowNumber -le 3) -and ($columnNumber -lt 3)) {

                    }
                    else {
                        $issueList += "Table has no headers"
                    }
                }
                $issueString = ""
                $issueList | Select-Object -Unique | ForEach-Object {$issueString += "$_`n"}
                if ($issueList.count -eq 0) {}
                else {
                    AddToArray "Table" "$($item.url -split `"api/v\d/`" -join `"`")"  "" "Table number $($tableNumber):`n$issueString" "Revise table"
                }
            }
        }
    }
}

function Start-ProcessSemantics {
    $i_tag_list = $page_body | Select-String -pattern "<i.*?>(.*?)</i>" -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
    $b_tag_list = $page_body | Select-String -pattern "<b.*?>(.*?)</b>" -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
    $i = 0
    foreach ($i_tag in $i_tag_list) {
        $i++
    }
    foreach ($b_tag in $b_tag_list) {
        $i++
    }
    if ($i -gt 0) {
        AddToArray "<i> or <b> tags" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Page contains <i> or <b> tags" "<i>/<b> tags should be <em>/<strong> tags"
    }
}

function Start-ProcessVideoTags {
    $videotag_list = $page_body -split "`n" | Select-String -pattern '<video.*?>.*?</video>' -AllMatches | ForEach-Object {$_.Matches.Value}
    foreach ($video in $videotag_list) {
        $src = $video | Select-String -pattern 'src="(.*?)"' -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}
        $videoID = $src.split('=')[1].split("&")[0]
        $transcript = Get-TranscriptAvailable $video
        if ($transcript) {}
        else {
            AddToArray "Inline Media Video" "$($item.url -split `"api/v\d/`" -join `"`")" $videoID "Inline Media Video`n" "No transcript found"
        }
    }
}

function Start-ProcessFlash {
    if ($page_body -match "Content on this page requires a newer version of Adobe Flash Player") {
        AddToArray "Flash Element" "$($item.url -split `"api/v\d/`" -join `"`")" "" "$($page_body.split("`n") -match "Content on this page requires a newer version of Adobe Flash Player" | Measure-Object | Select-Object -ExpandProperty Count) embedded flash elements on this page" "Flash is inaccessible"
    }
}

function Start-ProcessColor {
    $colorList = $page_body -split "`n" | Select-String -Pattern "((?:background-)?color:[^;`"]*)" -Allmatches | 
        ForEach-Object {$c = [PSCustomObject]@{
                                Color = ""
                                BackgroundColor = ""
                                }
                        if($_.Matches.Groups[1].Value -match "background")
                        {
                            $c.BackgroundColor = $_.Matches.Groups[1].Value -replace ".*?:", "" -replace " ", ""
                            $c.Color = $_.Matches.Groups[2].Value -replace ".*?:", "" -replace " ", ""
                        }elseif($_.Matches.Groups[2].Value -match "background"){
                            $c.BackgroundColor = $_.Matches.Groups[2].Value -replace ".*?:", "" -replace " ", ""
                            $c.Color = $_.Matches.Groups[1].Value -replace ".*?:", "" -replace " ", ""
                        }else{
                            $c.Color = $_.Matches.Groups[1].Value -replace ".*?:", "" -replace " ", ""
                        }
                        if($null -eq $c.BackgroundColor -or "" -eq $c.BackgroundColor)
                        {
                            $c.BackgroundColor = "#FFFFFF"
                        }
                        if($null -eq $c.Color -or "" -eq $c.Color)
                        {
                            $c.Color = "#000000"
                        }
                        $c
                        }
    Foreach ($color in $colorList) {
        if ($color.Color -notmatch "#") {
            $convert = @([System.drawing.Color]::($color.Color).R, [System.drawing.Color]::($color.Color).G, [System.drawing.Color]::($color.Color).B)
            $color.Color = '#' + -join (0..2| % {"{0:X2}" -f + ($convert[$_])})
        }
        if ($color.BackgroundColor -notmatch "#") {
            $convert = @([System.drawing.Color]::($color.BackgroundColor).R, [System.drawing.Color]::($color.BackgroundColor).G, [System.drawing.Color]::($color.BackgroundColor).B)
            $color.BackgroundColor = '#' + -join (0..2| % {"{0:X2}" -f + ($convert[$_])})
        }
        $color.Color = $color.Color.replace("#", "")
        $color.BackgroundColor = $color.BackgroundColor.replace("#", "")
        $results = (Invoke-WebRequest -Uri ("https://webaim.org/resources/contrastchecker/?fcolor={0}&bcolor={1}&api" -f $color.Color, $color.BackgroundColor)).Content | 
            ConvertFrom-Json
        if ($results.AA -ne 'pass') {
            AddToArray "Color Contrast" "$($item.url -split `"api/v\d/`" -join `"`")" "" "Color: $($color.Color)`nBackgroundColor: $($color.BackgroundColor)`n$($results -replace `"@{`", `"`" -replace `"}`",`"`" -replace `" `", `"`" -split `"`;`" -join "`n")" "Does not meet AA color contrast"
        }
    }
}

function AddToArray {
    param(
        [string]$element,
        [string]$location,
        [string]$VideoID,
        [string]$Text,
        [string]$Accessibility,
        [int]$issueSeverity = 1
    )
    $Global:elementList += $element
    $Global:locationList += $location
    $Global:videoIDList += $VideoID
    $Global:textList += $Text
    $Global:AccessibilityList += $Accessibility
    $Global:issueSeverityList += $issueSeverity
}