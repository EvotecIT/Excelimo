function Worksheet {
    param(
        [parameter(DontShow)] $ExcelDocument,
        [Array] $DataTable,
        [string] $Name,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Skip',
        [RGBColors] $TabColor = [RGBColors]::None
    )
    $Worksheet = Add-ExcelWorkSheetData -ExcelDocument $Script:ExcelDocument -WorksheetName $Name -Option $Option -Supress $false -DataTable $DataTable -TabColor $TabColor
    #$Worksheet
}