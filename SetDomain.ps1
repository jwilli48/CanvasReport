Write-Host "Which canvas are you using:"
$response = Read-Host "[1] Main, [2] Test, [3] MasterCourses"
if(Test-Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json"){
  $CanvasType =  (Get-Content "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" | ConvertFrom-Json).BaseUri
  switch ($CanvasType)
  {
    "https://byu.instructure.com"{
      Copy-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" -Destination "$env:HOMEDRIVE$env:HOMEPATH\Documents\BYU_CanvasApiCreds.json" -Force
      break
    }
    "https://byuistest.instructure.com"{
      Copy-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" -Destination "$env:HOMEDRIVE$env:HOMEPATH\Documents\TEST_CanvasApiCreds.json" -Force
      break
    }
    "https://byuismastercourses.instructure.com"{
      Copy-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" -Destination "$env:HOMEDRIVE$env:HOMEPATH\Documents\MASTER_CanvasApiCreds.json" -Force
      break
    }
  }
  Remove-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json"
}

switch ($response)
{
  "1"{
    if(Test-Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\BYU_CanvasApiCreds.json"){
      Copy-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\BYU_CanvasApiCreds.json" -Destination "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" -Force
    }else{
      Write-Host "You will need to create an Canvas API for this domain."
    }
    break
  }"2"{
    if(Test-Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\TEST_CanvasApiCreds.json"){
      Copy-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\TEST_CanvasApiCreds.json" -Destination "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" -Force
    }else{
      Write-Host "You will need to create an Canvas API for this domain."
    }
    break
  }"3"{
    if(Test-Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\MASTER_CanvasApiCreds.json"){
      Copy-Item -Path "$env:HOMEDRIVE$env:HOMEPATH\Documents\MASTER_CanvasApiCreds.json" -Destination "$env:HOMEDRIVE$env:HOMEPATH\Documents\CanvasApiCreds.json" -Force
    }else{
      Write-Host "You will need to create an Canvas API for this domain."
    }
    break
  }
}
