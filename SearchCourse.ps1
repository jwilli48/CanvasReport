Function Search-Course{
  param(
    $course_id
  )
  $Global:courseName = (Get-CanvasCoursesById -Id $course_id).name
  $courseName = $courseName -replace [regex]::escape('+'), ' ' -replace ':',''
  $course_modules = Get-CanvasModule -CourseId $course_id

  $i = 0
  foreach($module in $course_modules){
    $i++
    Write-Progress -Activity "Checking pages" -Status "Progress:" -PercentComplete ($i/$course_modules.length * 100)
    Write-Host "Module: $($module.name)" -ForegroundColor Cyan
    $moduleItems = Get-CanvasModuleItem -Course $course_id -ModuleId $module.id
    foreach($item in $ModuleItems){
      if($item.page_url -eq $NULL){
        #It is not a page
        continue
      }
      $page = Get-CanvasCoursesPagesByCourseIdAndUrl -CourseId $course_id -Url $item.page_url
      Write-Host $page.title -ForegroundColor Green
      $page_body = $page.body
      if($page_body -eq '' -or $page_body -eq $NULL){
        #Page is empty
        continue
      }
      Process_Contents $page_body
    }
  }
}
