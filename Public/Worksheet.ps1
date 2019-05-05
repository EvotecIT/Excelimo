function Worksheet {
    [CmdletBinding()]
    param(
        [Array] $DataTable,
        [string] $Name,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
        [RGBColors] $TabColor = [RGBColors]::None,
        [switch] $AutoFilter,
        [switch] $AutoFit
    )
    $ScriptBlock = {
        Param (
            $ExcelDocument,
            [Array] $DataTable,
            [string] $Name,
            [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
            [RGBColors] $TabColor = [RGBColors]::None,
            [bool] $Supress,
            [switch] $AutoFilter,
            [switch] $AutoFit
        )
        $addExcelWorkSheetDataSplat = @{
            DataTable          = $DataTable
            TabColor           = $TabColor
            Supress            = $Supress
            Option             = $Option
            ExcelDocument      = $ExcelDocument
            ExcelWorksheetName = $Name
            AutoFit            = $AutoFit
            AutoFilter         = $AutoFilter
        }
        Add-ExcelWorkSheetData @addExcelWorkSheetDataSplat -Verbose
    }
    $ExcelWorkSheetParameters = [ordered] @{
        DataTable     = $DataTable
        TabColor      = $TabColor
        Supress       = $true
        Option        = $Option
        ExcelDocument = $Script:Excel.ExcelDocument
        Name          = $Name
        AutoFit       = $AutoFit
        AutoFilter    = $AutoFilter
    }

    if ($Script:Excel.Runspaces.Parallel) {
        $RunSpace = Start-Runspace -ScriptBlock $ScriptBlock -Parameters $ExcelWorkSheetParameters -RunspacePool $Script:Excel.Runspaces.RunspacesPool -Verbose:$Verbose
        $Script:Excel.Runspaces.Runspaces.Add($RunSpace)
    } else {
        & $ScriptBlock -Parameters @ExcelWorkSheetParameters
    }
}