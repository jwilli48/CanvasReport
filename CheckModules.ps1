function Get-Modules{
  if(-not (Get-Module -ListAvailable -Name ImportExcel)){
    Write-Host "You need to have the ImportExcel module installed. Intalling now..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser
  }
  if(-not (Get-Module -ListAvailable -Name BurntToast)){
    Write-Host "You need to have the BurntToast module installed. Installing now..." -ForegroundColor Yellow
    Install-Module BurntToast -Scope CurrentUser
  }
  if(-not (Test-Path "$PSScriptRoot\Reports")){
    New-Item -Path "$PsScriptRoot\Reports" -ItemType Directory
  }
}
