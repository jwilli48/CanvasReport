Import-Module "$home/Desktop/AccessibilityTools/CanvasReport-master/Util.ps1" -Force

Function Search-Course {
    param(
        $course_id
    )
    $Global:courseName = (Get-CanvasCoursesById -Id $course_id).course_code
    $courseName = $courseName -replace [regex]::escape('+'), ' ' -replace ':', ''
    $course_modules = Get-CanvasModule -CourseId $course_id

    $i = 0
    foreach ($module in $course_modules) {
        $i++
        Write-Progress -Activity "Checking pages" -Status "Progress:" -PercentComplete ($i / $course_modules.length * 100)
        Write-Host "Module: $($module.name)" -ForegroundColor Cyan
        $moduleItems = Get-CanvasModuleItem -Course $course_id -ModuleId $module.id
        foreach ($item in $ModuleItems) {
            if ($item.type -eq "Page") {
                $page = Get-CanvasCoursesPagesByCourseIdAndUrl -CourseId $course_id -Url $item.page_url
                $page_body = $page.body
            }
            elseif ($item.type -eq "Discussion") {
                $page = Get-CanvasCoursesDiscussionTopicsByCourseIdAndTopicId -CourseId $course_id -TopicId $item.content_id
                $page_body = $page.message
            }
            elseif ($item.type -eq "Assignment") {
                $page = Get-CanvasCoursesAssignmentsByCourseIdAndId -CourseId $course_id -Id $item.content_id
                $page_body = $page.description
            }
            elseif ($item.type -eq "Quiz") {
                $page = Get-CanvasQuizzesById -CourseId $course_id -Id $item.content_id
                $page_body = $page.description
            }
            else {
                #if its not any of the above just skip it as it is not yet supported
                continue
            }
            Write-Host $item.title -ForegroundColor Green

            if ('' -eq $page_body -or $NULL -eq $page_body) {
                #Page is empty
                continue
            }
            Process_Contents $page_body
            if ($item.type -eq "Quiz") {
                try {
                    $quizQuestions = Get-CanvasQuizQuestion -CourseId $course_id -QuizId $item.content_id
                    foreach ($question in $quizQuestions) {
                        Process_Contents $question.question_text
                        foreach ($answer in $question.answers) {
                            Process_Contents $answer.html
                            Process_Contents $answer.comments_html
                        }
                    }
                }
                catch {
                    if ($_ -match "Unauthorized") {
                        Write-Host "ERROR: (401) Unauthorized, can not search quiz questions. Skipping..." -ForegroundColor Red
                    }
                    else {
                        Write-Host $_ -ForegroundColor Red
                    }
                }
            }
        }
    }
}

function Search-Directory {
    param(
        [string]$directory
    )

    $Global:courseName = $Directory.split('\')[-2]
    $Global:ReportType = "$($Directory[0])Drive"
    $course_files = Get-ChildItem "$directory\*.html" -Exclude '*old*', '*ImageGallery*', '*CourseMedia*', '*GENERIC*'
    if ($NULL -eq $course_files) {
        Write-Host "ERROR: Directory input is empty"
    }
    else {
        $i = 0
        foreach ($file in $course_files) {
            $i++
            Write-Progress -Activity "Checking pages" -Status "Progress:" -PercentComplete ($i / $course_files.length * 100)

            $file_content = Get-Content -Encoding UTF8 -Path $file.PSpath -raw
            if ($Directory[0] -eq "I") {
                $url = "https://iscontent.byu.edu/$($directory -replace `"I:\\`", `"`")/$($file.name)"
            }
            elseif ($Directory[0] -eq "Q") {
                $url = "https://isdev.byu.edu/courses/$($directory -replace `"Q:\\`", `"`")/$($file.name)"
            }
            else {
                $url = "file:///$directory/$($file.name)"
            }
            
            $item = Format-TransposeData body, title, url $file_content, $file.name, $url
            $page_body = $item.body
            Write-Host $item.title -ForegroundColor Green

            if ('' -eq $page_body -or $NULL -eq $page_body) {
                continue
            }
            Process_Contents $page_body
        }
        
        $PathToCss = "$($Directory.replace("HTML","CSS"))\this_course_only.css"
        if ((Test-Path $PathToCss)) {
            $file_content = Get-Content -Encoding UTF8 -Path $PathToCss -raw
            $url = "file:///$PathToCss"

            $item = Format-TransposeData body, title, url $file_content, $file_content.PSChildName, $url
            Write-Host $item.title -ForegroundColor Green

            $page_body = $item.body
            Process_Contents $page_body
        }
    }
}
