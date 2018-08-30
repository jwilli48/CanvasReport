Import-Module ./Util.ps1 -Force

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

  $data = Transpose-Data Element, Location, VideoID, Text, Accessibility, IssueSeverity $elementList, $locationList, $videoIDList, $textList, $AccessibilityList, $issueSeverityList
  $Global:ExcelReport = $PSScriptRoot + "\Reports\A11yReport_" + $courseName + ".xlsx"
  $markRed  = @((New-ConditionalText -Text "Adjust Link Text" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Needs a title" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "No Alt Attribute" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Alt Text May Need Adjustment" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "JavaScript links are not accessible" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Check if header" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Broken Link" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "Empty link tag" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "No transcript found" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'),(New-ConditionalText -Text "Revise table" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))
  if(-not ($data -eq $NULL)){
    $data | Export-Excel $ExcelReport -ConditionalText $markRed -AutoFilter -AutoSize -Append
  }
}

function Process-Links{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | % {$_.Matches.Value}
  foreach($link in $link_list){
    if($link.contains('onlick')){
      AddToArray "JavaScript Link" $page.title "" $link "JavaScript links are not accessible"
    }elseif(-not $link.contains('href')){
      AddToArray "Link" $page.title "" $link "Empty link tag"
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
            AddToArray "Link" $page.title "" $href "Broken link"
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
        AddToArray "Link" $page.title "" "Invisble link with no text" "Adjust Link Text"; break
      }
      "\bhere\b" {
        AddToArray "Link" $page.title "" $text "Adjust Link Text"; break
      }
      "Click Here" {
        AddToArray "Link" $page.title "" $text "Adjust Link Text"; break
      }
      "http"{
        AddToArray "Link" $page.title "" $text "Adjust Link Text"; break
      }
      "https"{
        AddToArray "Link" $page.title "" $text "Adjust Link Text"; break
      }
      "www\."{
        AddToArray "Link" $page.title "" $text "Adjust Link Text"; break
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
      AddToArray "Image" $page.title "" $img $Accessibility
    }else{
      $alt = $img | Select-String -pattern 'alt="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
      switch -regex ($alt)
      {
        "banner"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "Placeholder"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "\.jpg"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "\.png"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" "Alt text:`n$alt" $Accessibility
          Break
        }
        "https"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" "Alt text:`n$alt" $Accessibility
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
      $Accessibility = "Needs a title"

      if($iframe.contains('youtube')){
        $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('/')[4].split('?')[0]
        AddToArray "Youtube Video" $page.title $video_ID $title $Accessibility
      }
      elseif($iframe.contains('brightcove')){
        $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('=')[-1]
        AddToArray "Brightcove Video" $page.title $video_ID $title $Accessibility
      }
      elseif($iframe.contains('H5P')){
        AddToArray "H5P" $page.title "" $title $Accessibility
      }
      elseif($iframe.contains('byu.mediasite')){
        $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('/')[-1]
        if($video_ID -eq ""){
          $video_id = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('/')[-2]
        }
        AddToArray "BYU Mediasite Video" $page.title $video_ID $title $Accessibility
      }
      elseif($iframe.contains('Panopto')){
        $video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].Value}).split('=').split('&')[1]
        AddToArray "Panopto Video" $page.title $video_ID $title $Accessibility
      }
      else{
        AddToArray "Iframe" $page.title "" $title $Accessibility
      }
    }
  }
  #Check for transcripts
  $i = 1;
  foreach($iframe in $iframeList){
    if($iframe.contains('youtube') -or $iframe.contains('brightcove') -or $iframe.contains('byu.mediasite') -or $iframe.contains('Panopto')){
      if(-not (Get-TranscriptAvailable $iframe)){
        AddToArray "Transcript" "$($page.title)" "" "Video number $i on page" "No transcript found"
      }
    }
    $i++
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
        AddToArray "Header Level $headerLevel" $page.title "" $header $Accessibility
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
          if($check[$i] -match "colspan="){
            if($check[$i-1] -match "<tr" -and $check[$i+1] -match "</tr"){
              $issueList += "Stretched cell should possibly be a <caption> title for the table"
            }
          }
          if($check[$i] -match "<th"){
            $hasHeaders = $TRUE
            if($check[$i] -notmatch "scope"){
              $issueList += "Table headers should have either scope=`"row`" or scope=`"col`" for screenreaders"
            }
          }
          if($check[$i] -match "<tr"){
            $rowNumber++
          }elseif($check[$i] -match "<th" -or $check[$i] -match "<td"){
            $columnNumber++
          }elseif($check[$i] -match "</tr>"){
            if($check[$i] -match "</table>"){

            }else{
              $columnNumber = 0
            }
          }
          $i++
        }
        if(-not $hasHeaders){
          if($rowNumber -lt 3 -and $columnNumber -lt 3){

          }else{
            $issueList += "Table has no headers"
          }
        }
        $issueString = ""
        $issueList | Select-Object -Unique | % {$issueString += "$_`n"}
        if($issueList.count -eq 0){}
        else{
          AddToArray "Table" $page.title  "" "Table number $($tableNumber):`n$issueString" "Revise table"
        }
      }
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
