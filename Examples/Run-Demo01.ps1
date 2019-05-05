Import-Module .\Excelimo.psd1 -Force
Import-Module PSWriteExcel -Force

$Process = Get-Process | Select-Object -First 5

Excel -FilePath $PSScriptRoot\"Run-Demo01.xlsx" {
    WorkbookProperties -Title 'Test'

    Worksheet -DataTable $Process -Name 'Processes'
    Worksheet -DataTable $Process -Name 'Processes Test' -TabColor Crimson
} -Open