Import-Module ./Util.ps1 -Force

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
      if($item.type -eq "Page"){
        $page = Get-CanvasCoursesPagesByCourseIdAndUrl -CourseId $course_id -Url $item.page_url
        $page_body = $page.body
      }elseif($item.type -eq "Discussion"){
        $page = Get-CanvasCoursesDiscussionTopicsByCourseIdAndTopicId -CourseId $course_id -TopicId $item.content_id
        $page_body = $page.message
      }elseif($item.type -eq "Assignment"){
        $page = Get-CanvasCoursesAssignmentsByCourseIdAndId -CourseId $course_id -Id $item.content_id
        $page_body = $page.description
      }elseif($item.type -eq "Quiz"){
        $page = Get-CanvasQuizzesById -CourseId $course_id -Id $item.content_id
        $page_body = $page.description
      }else{
        #if its not any of the above just skip it as it is not yet supported
        continue
      }
      Write-Host $page.title -ForegroundColor Green

      if($page_body -eq '' -or $page_body -eq $NULL){
        #Page is empty
        continue
      }
      Process_Contents $page_body
      if($item.type -eq "Quiz"){
        try{
          $quizQuestions = Get-CanvasQuizQuestion -CourseId $course_id -QuizId $item.content_id
          foreach($question in $quizQuestions){
            Process_Contents $question.question_text
            foreach($answer in $question.answers){
              Process_Contents $answer.html
              Process_Contents $answer.comments_html
            }
          }
        }catch{
          if($_ -match "Unauthorized"){
            Write-Host "ERROR: (401) Unauthorized, can not search quiz questions. Skipping..." -ForegroundColor Red
          }else{
            Write-Host $_ -ForegroundColor Red
          }
        }
      }
    }
  }
}

function Search-Directory{
  param(
    [string]$directory
  )

  $Global:courseName = $Directory.split('\')[-2]
  $course_files = Get-ChildItem "$directory\*.html" -Exclude '*old*','*ImageGallery*', '*CourseMedia*'
  if($course_files -eq $NULL){
    Write-Host "ERROR: Directory input is empty"
  }else{
    $i = 0
    foreach($file in $course_files){
      $i++
      Write-Progress -Activity "Checking pages" -Status "Progress:" -PercentComplete ($i/$course_files.length * 100)

      $file_content = Get-Content -Encoding UTF8 -Path $file.PSpath -raw
      $item = Transpose-Data body, title $file_content, $file.name
      $page_body = $item.body
      Write-Host $item.title -ForegroundColor Green

      if($page_body -eq '' -or $page_body -eq $NULL){
        continue
      }
      Process_Contents $page_body
    }
  }
}
