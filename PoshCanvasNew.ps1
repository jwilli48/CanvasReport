0
<#

Generated by the Canvas PowerShell API Generator, by Spencer Varney

8/21/2018 1:41:30 PM

Use at your own risk!

#>

#region Base Canvas API Methods
function Get-CanvasCredentials(){
    if ($global:CanvasApiTokenInfo -eq $null) {

        $ApiInfoPath = "$env:USERPROFILE\Documents\CanvasApiCreds.json"

        #TODO: Once this is a module, load it from the module path: $PSScriptRoot or whatever that is
        if (-not (test-path $ApiInfoPath)) {
            $Token = Read-Host "Please enter your Canvas API API Access Token"
            $BaseUri = Read-Host "Please enter your Canvas API Base URI (for example, https://domain.beta.instructure.com)"

            $ApiInfo = [ordered]@{
                Token = $Token
                BaseUri = $BaseUri
            }

            $ApiInfo | ConvertTo-Json | Out-File -FilePath $ApiInfoPath
        }

        #load the file
        $global:CanvasApiTokenInfo = Get-Content -Path $ApiInfoPath | ConvertFrom-Json
    }

    return $global:CanvasApiTokenInfo
}

function Get-CanvasAuthHeader($Token) {
    return @{"Authorization"="Bearer "+$Token}
}

function Get-CanvasApiResult(){

    Param(
        $Uri,

        $RequestParameters,

        [ValidateSet("GET", "POST", "PUT", "DELETE")]
        $Method="GET"
    )

    $AuthInfo = Get-CanvasCredentials

    if ($RequestParameters -eq $null) { $RequestParameters = @{} }

    $RequestParameters["per_page"] = "10000"

    $Headers = (Get-CanvasAuthHeader $AuthInfo.Token)

    try {
    $Results = Invoke-WebRequest -Uri ($AuthInfo.BaseUri + $Uri) -ContentType "multipart/form-data" -Headers $headers -Method $Method -Body $RequestParameters
    } catch {
        throw $_.Exception.Message
    }

    $Content = $Results.Content | ConvertFrom-Json

    #Either PSCustomObject or Object[]
    if ($Content.GetType().Name -eq "PSCustomObject") {
        return $Content
    }

    $JsonResults = New-Object System.Collections.ArrayList

    $JsonResults.AddRange(($Results.Content | ConvertFrom-Json))

    if ($Results.Headers.link -ne $null) {
        $NextUriLine = $Results.Headers.link.Split(",") | where {$_.Contains("rel=`"next`"")}

        #$PerPage = $NextUriLine.Substring($NextUriLine.IndexOf("per_page=")+9) -replace '(\D).*',""

        if (-not [string]::IsNullOrWhiteSpace($NextUriLine)) {
            while ($Results.Headers.link.Contains("rel=`"next`"")) {

                $nextUri = $Results.Headers.link.Split(",") | `
                            where {$_.Contains("rel=`"next`"")} | `
                            % {$_ -replace ">; rel=`"next`""} |
                            % {$_ -replace "<"}

                #Write-Progress
                Write-Host $nextUri

                $Results = Invoke-WebRequest -Uri $nextUri -Headers $headers -Method Get -Body $RequestParameters -ContentType "multipart/form-data" `

                $JsonResults.AddRange(($Results.Content | ConvertFrom-Json))
            }
        }
    }

    return $JsonResults
}

#endregion
<#
.Synopsis
   Return information on a single course.

Accepts the same include[] parameters as the list action plus:
.EXAMPLE
   PS C:> Get-CanvasCoursesById -Id $SomeIdObj
#>
function Get-CanvasCoursesById {
[CmdletBinding()]
    Param (
        # ID
        [Parameter(Mandatory=$True)]
        [string] $Id,

        # - "all_courses": Also search recently deleted courses.
		# - "permissions": Include permissions the current user has
		#   for the course.
		# - "observed_users": include observed users in the enrollments
		# - "course_image": Optional course image data for when there is a course image
		#   and the course image feature flag has been enabled
        [Parameter(Mandatory=$False)]
         $Include
    )

    $Uri = "/api/v1/courses/$Id"

    $Body = @{}

	$Body["id"] = $Id

	if ($Include) { $Body["include"] = $Include }


    return Get-CanvasApiResult $Uri -Method GET -RequestParameters $Body

}
<#
.Synopsis
   A paginated list of the wiki pages associated with a course or group
.EXAMPLE
   PS C:> Get-CanvasCoursesPagesByCourseId -CourseId $SomeCourseIdObj
#>
function Get-CanvasCoursesPagesByCourseId {
[CmdletBinding()]
    Param (
        # ID
        [Parameter(Mandatory=$True)]
        [string] $CourseId,

        # Sort results by this field.
        [Parameter(Mandatory=$False)]
        [string] $Sort,

        # The sorting order. Defaults to 'asc'.
        [Parameter(Mandatory=$False)]
        [string] $Order,

        # The partial title of the pages to match and return.
        [Parameter(Mandatory=$False)]
        [string] $SearchTerm,

        # If true, include only published paqes. If false, exclude published
		# pages. If not present, do not filter on published status.
        [Parameter(Mandatory=$False)]
        [bool] $Published
    )

    $Uri = "/api/v1/courses/$CourseId/pages"

    $Body = @{}

	$Body["course_id"] = $CourseId

	if ($Sort) { $Body["sort"] = $Sort }

	if ($Order) { $Body["order"] = $Order }

	if ($SearchTerm) { $Body["search_term"] = $SearchTerm }

	if ($Published) { $Body["published"] = $Published }


    return Get-CanvasApiResult $Uri -Method GET -RequestParameters $Body

}
<#
.Synopsis
   Retrieve the content of a wiki page
.EXAMPLE
   PS C:> Get-CanvasCoursesPagesByCourseIdAndUrl -CourseId $SomeCourseIdObj -Url $SomeUrlObj
#>
function Get-CanvasCoursesPagesByCourseIdAndUrl {
[CmdletBinding()]
    Param (
        # ID
        [Parameter(Mandatory=$True)]
        [string] $CourseId,

        # ID
        [Parameter(Mandatory=$True)]
        [string] $Url
    )

    $Uri = "/api/v1/courses/$CourseId/pages/$Url"

    $Body = @{}

	$Body["course_id"] = $CourseId

	$Body["url"] = $Url


    return Get-CanvasApiResult $Uri -Method GET -RequestParameters $Body

}

<#
.Synopsis
   A paginated list of the modules in a course
.EXAMPLE
   PS C:> Get-CanvasModule -CourseId $SomeCourseIdObj
#>
function Get-CanvasModule {
[CmdletBinding()]
    Param (
        # ID
        [Parameter(Mandatory=$True)]
        [string] $CourseId,

        # - "items": Return module items inline if possible.
		#   This parameter suggests that Canvas return module items directly
		#   in the Module object JSON, to avoid having to make separate API
		#   requests for each module when enumerating modules and items. Canvas
		#   is free to omit 'items' for any particular module if it deems them
		#   too numerous to return inline. Callers must be prepared to use the
		#   [api:ContextModuleItemsApiController#index List Module Items API]
		#   if items are not returned.
		# - "content_details": Requires include['items']. Returns additional
		#   details with module items specific to their associated content items.
		#   Includes standard lock information for each item.
        [Parameter(Mandatory=$False)]
         $Include,

        # The partial name of the modules (and module items, if include['items'] is
		# specified) to match and return.
        [Parameter(Mandatory=$False)]
        [string] $SearchTerm,

        # Returns module completion information for the student with this id.
        [Parameter(Mandatory=$False)]
        [string] $StudentId
    )

    $Uri = "/api/v1/courses/$CourseId/modules"

    $Body = @{}

	$Body["course_id"] = $CourseId

	if ($Include) { $Body["include"] = $Include }

	if ($SearchTerm) { $Body["search_term"] = $SearchTerm }

	if ($StudentId) { $Body["student_id"] = $StudentId }


    return Get-CanvasApiResult $Uri -Method GET -RequestParameters $Body

}

<#
.Synopsis
   A paginated list of the items in a module
.EXAMPLE
   PS C:> Get-CanvasModuleItem -CourseId $SomeCourseIdObj -ModuleId $SomeModuleIdObj
#>
function Get-CanvasModuleItem {
[CmdletBinding()]
    Param (
        # ID
        [Parameter(Mandatory=$True)]
        [string] $CourseId,

        # ID
        [Parameter(Mandatory=$True)]
        [string] $ModuleId,

        # If included, will return additional details specific to the content
		# associated with each item. Refer to the [api:Modules:Module%20Item Module
		# Item specification] for more details.
		# Includes standard lock information for each item.
        [Parameter(Mandatory=$False)]
         $Include,

        # The partial title of the items to match and return.
        [Parameter(Mandatory=$False)]
        [string] $SearchTerm,

        # Returns module completion information for the student with this id.
        [Parameter(Mandatory=$False)]
        [string] $StudentId
    )

    $Uri = "/api/v1/courses/$CourseId/modules/$ModuleId/items"

    $Body = @{}

	$Body["course_id"] = $CourseId

	$Body["module_id"] = $ModuleId

	if ($Include) { $Body["include"] = $Include }

	if ($SearchTerm) { $Body["search_term"] = $SearchTerm }

	if ($StudentId) { $Body["student_id"] = $StudentId }


    return Get-CanvasApiResult $Uri -Method GET -RequestParameters $Body

}
