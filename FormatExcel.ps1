#ACCESSIBILITY REPORT FORMATTING
function Format-A11yExcel{
  <#
  .DESCRIPTION

  Formats the excel sheet, which is actually not needed anymore as we now use a template
  #>
  $excel = Export-Excel $ExcelReport -PassThru
  if(-not ($excel -eq $NULL)){
    $excel.Workbook.Worksheets["Sheet1"].Column(1).Width = 25
    $excel.Workbook.Worksheets["Sheet1"].Column(4).Width = 75
    $excel.Workbook.Worksheets["Sheet1"].Column(4).Style.wraptext = $true
    $excel.Workbook.Worksheets["Sheet1"].Column(6).Width = 25
    $excel.Save()
    $excel.Dispose()
  }
}

function ConvertTo-A11yExcel{
  <#
  .DESCRIPTION

  Moves the default excel sheet made into the template excel sheet. Takes the each row of default excel sheet then adds it to the correct table with the correct values in the template then saves over the old one.
  #>
  $template = Open-ExcelPackage -Path "$PsScriptRoot\CAR - Accessibility Review Template.xlsx"
  $data = Import-Excel -path $ExcelReport
  $cell = $template.Workbook.Worksheets[1].Cells
  $rowNumber = 9
  for($i = 0; $i -lt $data.length; $i++){
    $cell[$rowNumber,2].Value = "Not Started"
    $cell[$rowNumber,3].Value = $data[$i].Location
    switch ($data[$i].Accessibility)
    {
      "Needs a title attribute"{
        AddToCell "Semantics" "Missing title/label" "$($data[$i].Element) needs a title attribute`nID: $($data[$i].VideoID)"
        Break
      }
      "Adjust Link Text"{
        AddToCell "Link" "Non-Descriptive Link" "$($data[$i].Text)"
        Break
      }
      "No Alt Attribute"{
        AddToCell "Image" "No Alt Attribute"  ""
        Break
      }
      "Alt Text May Need Adjustment"{
        AddToCell "Image" "Non-Descriptive alt tags" "$($data[$i].Text)"
        break
      }
      "JavaScript links are not accessible"{
        AddToCell "Link" "" "$($data[$i].Text)`n$($data[$i].Accessibility)" 3 3 3
        break
      }
      "Check if header is meant to be invisible and is not a duplicate"{
        AddToCell "Semantics" "Improper Headings" "$($data[$i].Text), a $($data[$i].Element), is invisible" 3 3 3
        break
      }
      "Broken link"{
        AddToCell "Link" "Broken Link" "$($data[$i].Text)" 5 5 5
        break
      }"No transcript found"{
		    AddToCell "Media" "Transcript Needed" "$($data[$i].Text)" 5 5 5
        break
	    }"Revise table"{
	      AddToCell "Table" "" "$($data[$i].Text)"
        break
	    }"<i>/<b> tags should be <em>/<strong> tags"{
        AddToCell "Semantics" "Bad use of <i> and/or <b>" "$($data[$i].Accessibility)"
      }"Empty link tag"{
        AddToCell "Link" "Broken Link" "$($data[$i].Text)"
      }"Flash is inaccessible"{
        AddToCell "Misc" "" "$($data[$i].Text)`n$($data[$i].Accessibility)"
      }default{
        AddToCell "" "" "$($data[$i].Element), $($data[$i].Text)"
      }
    }
    $rowNumber++
  }

  $column = 3 #C
  $row = 9 #start of data
  while($NULL -ne $cell[$row,$column]){
    $cell[$row,$column].Hyperlink = $cell[$row,$column].Value
    $cell[$row,$column].Value = $cell[$row,$column].Value.Split("/").split("\")[-1]
    $row++
  }

  $template.Workbook.Worksheets[1].ConditionalFormatting[0].LowValue.Color = [System.Drawing.Color]::FromArgb(255,146,208,80)
  $template.Workbook.Worksheets[1].ConditionalFormatting[0].MiddleValue.Color = [System.Drawing.Color]::FromArgb(255,255,213,5)
  $template.Workbook.Worksheets[1].ConditionalFormatting[0].HighValue.Color = [System.Drawing.Color]::FromArgb(255,255,71,71)
  $template.Workbook.Worksheets[1].ConditionalFormatting[1].LowValue.Color = [System.Drawing.Color]::FromArgb(255,146,208,80)
  $template.Workbook.Worksheets[1].ConditionalFormatting[1].MiddleValue.Color = [System.Drawing.Color]::FromArgb(255,255,213,5)
  $template.Workbook.Worksheets[1].ConditionalFormatting[1].HighValue.Color = [System.Drawing.Color]::FromArgb(255,255,71,71)
  Set-Format -WorkSheet $template.Workbook.Worksheets[1] -Range "E6:E6" -NumberFormat 'Short Date'
  Close-ExcelPackage $template -SaveAs "$ExcelReport"
}

function Add-LocationLinks{
    $excel = Open-ExcelPackage -path $ExcelReport
    $cells = $excel.Workbook.Worksheets[1].Cells
    $column = 3 #C
    $row = 9 #start of data
    while($NULL -ne $cells[$row,$column]){
      $cells[$row,$column].Hyperlink = $cells[$row,$column].Value
      $cells[$row,$column].Value = $cells[$row,$column].Value.Split("/").split("\")[-1]
      $row++
    }
    Close-ExcelPackage $excel
}

function AddToCell{
  <#
  .DESCRIPTION

  Used in the ConvertTo-A11yExcel function, used to simplify adding data to the correct cells.
  #>
  param(
    [string]$issueType,
    [string]$DescriptiveError,
    [string]$Notes,
    [int]$Serverity = 1,
    [int]$Occurence = 1,
    [int]$Detection = 1
  )
  $cell[$rowNumber,4].Value = $issueType
  $cell[$rowNumber,5].Value = $DescriptiveError
  $cell[$rowNumber,6].Value = $Notes
  $cell[$rowNumber,7].Value = $Serverity
  $cell[$rowNumber,8].Value = $Occurence
  $cell[$rowNumber,9].Value = $Detection
}

function Get-A11yPivotTables{
  <#
  .DESCRIPTION

  Was used when we didn't have the template, it added simple Pivot charts and tables to the excel sheet to give useful information at a glance.
  #>
  Export-Excel $ExcelReport -Numberformat '#############' -IncludePivotTable -IncludePivotChart -PivotRows "A" -PivotData @{F='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName 'IssueSeverity'
  Export-Excel $ExcelReport -Numberformat '#############' -IncludePivotTable -IncludePivotChart -PivotRows "Location" -PivotData @{IssueSeverity='sum'} -ChartType PieExploded3d -ShowPercent -PivotTableName 'IssueLocations' -NoLegend
}

#MEDIA REPORT FORMATTING
function Format-MediaExcel1{
  <#
  .DESCRIPTION

  Formats the media excel sheet which we do not have a template for. Makes sure that some of the rows do not become to long, gives it a max width and makes it so the column has text wrap. Also makes sure the number format is correct for both the video ID numbers in column C, and for the video lengths in column D
  #>
  $excel = Export-Excel $ExcelReport -PassThru
  if(-not ($excel -eq $NULL)){
    $excel.Workbook.Worksheets["Sheet1"].Column(4).Width = 25
    $excel.Workbook.Worksheets["Sheet1"].Column(5).Width = 75
    $excel.Workbook.Worksheets["Sheet1"].Column(5).Style.wraptext = $true
    $excel.Workbook.Worksheets["Sheet1"].Column(6).Width = 25
    $sheet = $excel.Workbook.Worksheets["Sheet1"]
    Set-Format -WorkSheet $sheet -Range "D:D" -NumberFormat "hh:mm:ss"
    Set-Format -WorkSheet $sheet -Range "C:C" -NumberFormat "#############"
    $excel.Save()
    $excel.Dispose()
  }
}

function Format-MediaExcel2{
  <#
  .DESCRIPTION

  After the pivot tables are added there are some additional formatting needed for those tables. The default time for the tables is in days and I want it in hours.
  #>
  $excel = Export-Excel $ExcelReport -PassThru
  $excel.Workbook.Worksheets[3].PivotTables[0].DataFields[0].Format = "[h]:mm:ss"
  $excel.Workbook.Worksheets[4].PivotTables[0].DataFields[0].Format = "[h]:mm:ss"
  $excel.Workbook.Worksheets[5].PivotTables[0].DataFields[0].Format = "[h]:mm:ss"
  Close-ExcelPackage $excel
}

function Get-MediaPivotTables{
  <#
  .DESCRIPTION

  Adds a bunch of tables and charts for quick information:
  1. Table and Chart of how many of each type of media there was found
  2. Table and Chart of how total time for each type of media
  3. Table and Chart of total video time per location in the course
  4. Table and Chart of total time based on transcripts found/not found. Can also be split to show more details, such as within the videos that don't have transcripts you can split it to see how much time for each media type.
  #>
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Element" -PivotData @{MediaCount='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName "MediaTypes"
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Element" -PivotData @{VideoLength='sum'} -ChartType PieExploded3d -ShowPercent -PivotTableName "MediaLength"
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Location" -PivotData @{VideoLength='sum'} -ChartType PieExploded3d -ShowPercent -PivotTableName "MediaLengthByLocation" -NoLegend
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Transcript" -PivotData @{VideoLength='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName "TranscriptsVideoLength"
}
