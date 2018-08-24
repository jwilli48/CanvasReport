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

function Get-A11yPivotTables{
  Export-Excel $ExcelReport -Numberformat '#############' -IncludePivotTable -IncludePivotChart -PivotRows "Element" -PivotData @{IssueSeverity='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName 'IssueSeverity'
  Export-Excel $ExcelReport -IncludePivotTable -IncludePivotChart -PivotRows "Location" -PivotData @{IssueSeverity='sum'} -ChartType PieExploded3d -ShowCategory -ShowPercent -PivotTableName 'IssueLocations' -NoLegend
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
}
