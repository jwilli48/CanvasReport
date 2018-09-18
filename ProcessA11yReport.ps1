function Process_Contents{
  param(
    [string]$page_body
  )
  $Global:elementList = @()
  $Global:locationList = @()
  $Global:videoIDList = @()
  $Global:textList = @()
  $Global:AccessibilityList = @()
  $Global:issueSeverityList = @()

  Process-Links
  Process-Iframes
  Process-Images
  Process-Headers
  Process-Tables
  Process-Semantics
  Process-VideoTags
  Process-BrightcoveVideoHTML

  $data = Transpose-Data Element, Location, VideoID, Text, Accessibility, IssueSeverity $elementList, $locationList, $videoIDList, $textList, $AccessibilityList, $issueSeverityList
  $Global:ExcelReport = $PSScriptRoot + "\Reports\A11yReport_" + $courseName + ".xlsx"

  ###ConditionalText no longer used, changed to using a Template instead###
  #$markRed  = @((New-ConditionalText -Text "Adjust Link Text" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Needs a title" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "No Alt Attribute" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Alt Text May Need Adjustment" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "JavaScript links are not accessible" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Check if header" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Broken Link" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "Empty link tag" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "No transcript found" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "Revise table" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))

  #$hightlight = @((New-ConditionalText -Text "<i> tags should be <em> tags" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ),(New-ConditionalText -Text "<b> tags should be <strong> tags" -BackgroundColor '#ffff8b' -ConditionalTextColor '#000000' ))

  if(-not ($data -eq $NULL)){
    $data | Export-Excel $ExcelReport <#-ConditionalText $markRed, $highlight#> -AutoFilter -AutoSize -Append
  }
}

function Process-Links{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | % {$_.Matches.Value}
  foreach($link in $link_list){
    if($link.contains('onlick')){
      AddToArray "JavaScript Link" $item.title "" $link "JavaScript links are not accessible"
    }<#elseif($link.contains('href=".*?javascript.*?"')){
      AddToArray "JavaScript Link" $item.title "" $link "JavaScript links are not accessible"
    }#>elseif(-not $link.contains('href')){
      AddToArray "Link" $item.title "" $link "Empty link tag"
    }elseif($link -match 'href="\s*?"'){
      AddToArray "Link" $item.title "" $link "Empty link tag"
    }
  }
  
  $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}

  <#Checks broken links, not needed since Canvas has it built in
  if($Global:CheckLinks -eq $NULL){
    $host.ui.rawui.foregroundColor = "Yellow"
    $Global:CheckLinks = Read-Host "Would you like to check for broken links? (Y/N)`nThis on average doubles the time to generate this report"
    $host.ui.rawui.foregroundColor = "White"
  }else{
    if($Global:CheckLinks -match "Y"){
      foreach($href in $href_list){
        if($href.contains("mailto")){
          #email link
          continue
        }
        try{
          Invoke-WebRequest $href | Out-Null
        }catch{
          #There is also a SecureChannelFailure error status but it seems that for the majority of cases those still work when clicking the link manually
          if($_.Exception.Status -eq "SendFailure"){
            AddToArray "Link" $item.title "" $href "Broken link"
          }
        }
      }
    }
  }#>
  $link_text = $link_list | Select-String -pattern '<a.*?>(.*?)</a>' -AllMatches | % {$_.Matches.Groups[1].Value}
  foreach($text in $link_text){
    switch -regex ($text)
    {
      '<img' {
        #It will be caught by Proccess-Image
        break
      }
      $NULL{
        AddToArray "Link" $item.title "" "Invisble link with no text" "Adjust Link Text"; break
      }
      "^ ?[A-Za-z\.]+ ?$" {#This matches if the link text is a sigle word
        AddToArray "Link" $item.title "" $text "Adjust Link Text"; break
      }
      "Click" {
        AddToArray "Link" $item.title "" $text "Adjust Link Text"; break
      }
      "http"{
        AddToArray "Link" $item.title "" $text "Adjust Link Text"; break
      }
      "https"{
        AddToArray "Link" $item.title "" $text "Adjust Link Text"; break
      }
      "www\."{
        AddToArray "Link" $item.title "" $text "Adjust Link Text"; break
      }
      "Link"{
        if(-not ($text -match "Links to an external site")){
        AddToArray "Link" $item.title "" $text "Adjust Link Text"; break
        }
      }
      Default {}
    }
  }
}

function Process-Images{
  $image_list = $page_body | Select-String -pattern '<img.*?>' -AllMatches | % {$_.Matches.Value}
  foreach($img in $image_list){
    $alt = ""
    if(-not $img.contains('alt')){
      $Accessibility = "No Alt Attribute"
      AddToArray "Image" $item.title "" $img $Accessibility
    }else{
      $alt = $img | Select-String -pattern 'alt="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
      $Accessibility = "Alt Text May Need Adjustment"
      switch -regex ($alt)
      {
        "banner"{
          AddToArray "Image" $item.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "Placeholder"{
          AddToArray "Image" $item.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "\.jpg"{
          AddToArray "Image" $item.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "\.png"{
          AddToArray "Image" $item.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "https"{
          AddToArray "Image" $item.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        Default{}
      }
    }
  }

}

function Process-Iframes{
  $iframeList = $page_body | Select-String -pattern "<iframe.*?>.*?</iframe>" -AllMatches | % {$_.Matches.Value}
  foreach($iframe in $iframeList){
    $title = ""
    if(-not $iframe.contains('title')){
      $Accessibility = "Needs a title attribute"

      if($iframe.contains('youtube')){
        $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('/')[4].split('?')[0]
        AddToArray "Youtube Video" $item.title $video_ID $title $Accessibility
      }
      elseif($iframe.contains('brightcove')){
        $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('=')[-1].split("&")[0]
        AddToArray "Brightcove Video" $item.title $video_ID $title $Accessibility
      }
      elseif($iframe.contains('H5P')){
        AddToArray "H5P" $item.title "" $title $Accessibility
      }
      elseif($iframe.contains('byu.mediasite')){
        $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('/')[-1]
        if($video_ID -eq ""){
          $video_id = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('/')[-2]
        }
        AddToArray "BYU Mediasite Video" $item.title $video_ID $title $Accessibility
      }
      elseif($iframe.contains('Panopto')){
        $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('=').split('&')[1]
        AddToArray "Panopto Video" $item.title $video_ID $title $Accessibility
      }
      else{
        AddToArray "Iframe" $item.title "" $title $Accessibility
      }
    }
  }
  #Check for transcripts
  $i = 1;
  foreach($iframe in $iframeList){
    if($iframe.contains('brightcove') -or $iframe.contains('byu.mediasite') -or $iframe.contains('Panopto')){
      if(-not (Get-TranscriptAvailable $iframe)){
        AddToArray "Transcript" "$($item.title)" "" "Video number $i on page" "No transcript found"
      }
    }
    $i++
  }
}

function Process-BrightcoveVideoHTML{
  $brightcove_list = $page_body | Select-String -pattern '<div id="[^\d]*(\d{13})"' -Allmatches | % {$_.Matches.Value}
  $id_list = $brightcove_list | Select-String -pattern '\d{13}' -AllMatches | % {$_.matches.Value}
  foreach($id in $id_list){
    $transcriptCheck = $page_body.split("`n")
    $i = 0
    while($transcriptCheck[$i] -notmatch "$id"){$i++}
    $transcript = $FALSE
    for($j = $i; $j -lt ($i +10); $j++){
      if($transcript[$j] -eq $NULL){
        #End of file
        break
      }elseif($transcript[$j] -match "transcript")){
        $transcript = $TRUE
        break
      }
    }
    if($transcript){$transcript = "Yes"}
    else{
      $transcript = "No"
      AddToArray "Transcript" $item.title "$id" "No transcript found for BrightCove video with id:`n$id" "No transcript found"
    }
  }
}

function Process-Headers{
  $headerList = $page_body | Select-String -pattern '<h\d.*?>.*?</h\d>' -Allmatches | % {$_.Matches.Value}
  $accessibility = ""
  foreach($header in $headerList){
    $headerLevel = $header | Select-String -Pattern "<h(\d)" -Allmatches | % {$_.matches.Groups[1].Value}
    $headerText = $header | Select-String -pattern '<h\d.*?>(.*?)</h\d>' -AllMatches | % {$_.Matches.Groups[1].Value}
    switch -regex ($header)
    {
      'class=".*?screenreader-only.*?"'{
        $accessibility = "Check if header is meant to be invisible and is not a duplicate"
        AddToArray "Header Level $headerLevel" $item.title "" $header $Accessibility
        break
      }
    }
  }
}

function Process-Tables{
  if($page_body.contains("<table")){
    $tableNumber = 0
    $check = $page_body.split("`n")
    for($i = 0; $i -lt $check.length; $i++){
      $issueList = @()
      if($check[$i].contains("<table")){
        $rowNumber = 0
        $columnNumber = 0
        $tableNumber++
        $hasHeaders = $FALSE
        #Starts going through the whole table line by line
        while(-not ($check[$i].contains("</table>"))){
          #If table contains an heading tags it is an accessibility issue
          if($check[$i] -match "<h\d"){
            $issueList += "Heading tags should not be inside of tables"
          }

          if($check[$i] -match "colspan"){
            if($check[$i-1] -match "<tr" -and $check[$i+1] -match "</tr"){
              $issueList += "Stretched cell(s) should possibly be a <caption> title for the table"
            }
          }
          elseif($check[$i] -match "<th[^e]"){
            $hasHeaders = $TRUE
            if($check[$i] -notmatch "scope"){
              $issueList += "Table headers should have either scope=`"row`" or scope=`"col`" for screenreaders"
            }
          }
          elseif($check[$i] -match "<td"){
            if($check[$i] -match "scope"){
              $issueList += "Non-header table cells should not have scope attributes"
            }
          }
          elseif($check[$i] -match "<tr"){
            $rowNumber++
          }elseif($check[$i] -match "<th" -or $check[$i] -match "<td"){
            $columnNumber++
          }elseif($check[$i] -match "</tr>"){
            if($check[$i+1] -notmatch "<tr"){}
            else{
              $columnNumber = 0
            }
          }
          $i++
        }
        if(-not $hasHeaders){
          if(($rowNumber -le 3) -and ($columnNumber -lt 3)){

          }else{
            $issueList += "Table has no headers"
          }
        }
        $issueString = ""
        $issueList | Select-Object -Unique | % {$issueString += "$_`n"}
        if($issueList.count -eq 0){}
        else{
          AddToArray "Table" $item.title  "" "Table number $($tableNumber):`n$issueString" "Revise table"
        }
      }
    }
  }
}

function Process-Semantics{
  $i_tag_list = $page_body | Select-String -pattern "<i.*?>(.*?)</i>" -AllMatches | % {$_.Matches.Groups[1].Value}
  $b_tag_list = $page_body | Select-String -pattern "<b.*?>(.*?)</b>" -AllMatches | % {$_.Matches.Groups[1].Value}
  $i = 0
  foreach($i_tag in $i_tag_list){
    $i++
  }
  foreach($b_tag in $b_tag_list){
    $i++
  }
  if($i -gt 0){
    AddToArray "<i> or <b> tags" $item.title "" "Page contains <i> or <b> tags" "<i>/<b> tags should be <em>/<strong> tags"
  }
}

function Process-VideoTags{
    $videotag_list = $page_body | Select-String -pattern '<video.*?>.*?</video>' -AllMatches | %{$_.Matches.Value}
    foreach($video in $videotag_list){
      $src = $video | Select-String -pattern 'src="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
      $videoID = $src.split('=')[1].split("&")[0]
      $transcript = Get-TranscriptAvailable $video
      if($transcript){}
      else{
        AddToArray "Inline Media Video" $item.title $videoID "Inline Media Video`n" "No transcript found"
      }
    }
}

function AddToArray{
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
