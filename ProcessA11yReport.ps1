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

  $data = Transpose-Data Element, Location, VideoID, Text, Accessibility, IssueSeverity $elementList, $locationList, $videoIDList, $textList, $AccessibilityList, $issueSeverityList
  $Global:ExcelReport = $PSScriptRoot + "\Reports\A11yReport_" + $courseName + ".xlsx"
  $markRed  = @((New-ConditionalText -Text "Adjust Link Text" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Needs a title" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "No Alt Attribute" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'), (New-ConditionalText -Text "Alt Text May Need Adjustment" -BackgroundColor '#ff5454' -ConditionalTextColor '#000000'))
  if(-not ($data -eq $NULL)){
    $data | Export-Excel $ExcelReport -ConditionalText $markRed -AutoFilter -AutoSize -Append
  }
}

function Process-Links{
  $link_list = $page_body | Select-String -pattern "<a.*?>.*?</a>" -AllMatches | % {$_.Matches.Value}
  $href_list = $link_list | Select-String -pattern 'href="(.*?)"' -AllMatches | % {$_.Matches.Groups[1].Value}
  $link_text = $link_list | Select-String -pattern '<a.*?>(.*?)</a>' -AllMatches | % {$_.Matches.Groups[1].Value}
  for($i = 0; $i -lt $link_text.length; $i++){
    switch -regex ($link_text[$i])
    {
      '<img' {
        #It will be caught by Proccess-Image
        break
      }
      "here" {
        AddToArray "Link" $page.title "" $link_text[$i] "Adjust Link Text"; break
      }
      "Click Here" {
        AddToArray "Link" $page.title "" $link_text[$i] "Adjust Link Text"; break
      }
      "https"{
        AddToArray "Link" $page.title "" $link_text[$i] "Adjust Link Text"; break
      }
      "www\."{
        AddToArray "Link" $page.title "" $link_text[$i] "Adjust Link Text"; break
      }
      Default {}
    }
  }
}

function Process-Images{
  $image_list = $page_body | Select-String -pattern '<img.*>' -AllMatches | % {$_.Matches.Value}
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
          AddToArray "Image" $page.title "" $alt $Accessibility
          Break
        }
        "Placeholder"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" $alt $Accessibility
          Break
        }
        "\.jpg"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" $alt $Accessibility
          Break
        }
        "\.png"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" $alt $Accessibility
          Break
        }
        "https"{
          $Accessibility = "Alt Text May Need Adjustment"
          AddToArray "Image" $page.title "" $alt $Accessibility
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
      }elseif($iframe.contains('brightcove')){
        $Video_ID = ($iframe | Select-String -pattern 'src="(.*?)"' | % {$_.Matches.Groups[1].value}).split('=')[-1]
        AddToArray "Brightcove Video" $page.title $video_ID $title $Accessibility
      }elseif($iframe.contains('H5P')){
        AddToArray "H5P" $page.title "" $title $Accessibility
      }else{
        AddToArray "Iframe" $page.title "" $title $Accessibility
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
