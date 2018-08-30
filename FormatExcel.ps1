#ACCESSIBILITY REPORT FORMATTING
function Format-A11yExcel{
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
  $template = Open-ExcelPackage -Path "$PsScriptRoot\CAR - Accessibility Review Template.xlsx"
  $data = Import-Excel -path $ExcelReport
  $cell = $template.Workbook.Worksheets[1].Cells
  $rowNumber = 9
  for($i = 0; $i -lt $data.length; $i++){
    $cell[$rowNumber,2].Value = "Not Started"
    $cell[$rowNumber,3].Value = $data[$i].Location
    switch ($data[$i].Accessibility)
    {
      "Needs a title"{
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
        AddToCell "Link" "" "$($data[$i].Text) is a javascript link" 3 3 3
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
	  }"Revise table"{
		AddToCell "Table" "" "$($data[$i].Text)"
	  }default{
        AddToCell "" "" "$($data[$i].Element), $($data[$i].Text)"
      }
    }
    $rowNumber++
  }
  $template.Workbook.Worksheets[1].ConditionalFormatting[0].LowValue.Color = [System.Drawing.Color]::FromArgb(255,146,208,80)
  $template.Workbook.Worksheets[1].ConditionalFormatting[0].MiddleValue.Color = [System.Drawing.Color]::FromArgb(255,255,213,5)
  $template.Workbook.Worksheets[1].ConditionalFormatting[0].HighValue.Color = [System.Drawing.Color]::FromArgb(255,255,71,71)
  $template.Workbook.Worksheets[1].ConditionalFormatting[1].LowValue.Color = [System.Drawing.Color]::FromArgb(255,146,208,80)
  $template.Workbook.Worksheets[1].ConditionalFormatting[1].MiddleValue.Color = [System.Drawing.Color]::FromArgb(255,255,213,5)
  $template.Workbook.Worksheets[1].ConditionalFormatting[1].HighValue.Color = [System.Drawing.Color]::FromArgb(255,255,71,71)
  Close-ExcelPackage $template -SaveAs "$ExcelReport"
}

function AddToCell{
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
  Export-Excel $ExcelReport -Numberformat '#############' -IncludePivotTable -IncludePivotChart -PivotRows "A" -PivotData @{F='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName 'IssueSeverity'
  Export-Excel $ExcelReport -Numberformat '#############' -IncludePivotTable -IncludePivotChart -PivotRows "Location" -PivotData @{IssueSeverity='sum'} -ChartType PieExploded3d -ShowPercent -PivotTableName 'IssueLocations' -NoLegend
}

#MEDIA REPORT FORMATTING
function Format-MediaExcel1{
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
  $excel = Export-Excel $ExcelReport -PassThru
  Set-Format -WorkSheet $excel.Workbook.WorkSheets["MediaLength"] -Range "B:B" -NumberFormat '[h]:mm:ss'
  Set-Format -WorkSheet $excel.Workbook.WorkSheets["MediaLengthByLocation"] -Range "B:B" -NumberFormat '[h]:mm:ss'
  Close-ExcelPackage $excel
}

function Get-MediaPivotTables{
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Element" -PivotData @{MediaCount='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName "MediaTypes"
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Element" -PivotData @{VideoLength='sum'} -ChartType PieExploded3d -ShowPercent -PivotTableName "MediaLength"
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Location" -PivotData @{VideoLength='sum'} -ChartType PieExploded3d -ShowPercent -PivotTableName "MediaLengthByLocation" -NoLegend
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Transcript" -PivotData @{VideoLength='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName "TranscriptsVideoLength"
}
