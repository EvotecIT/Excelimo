function Worksheet {
    [CmdletBinding()]
    param(
        [parameter(DontShow)] $ExcelDocument,
        [Array] $DataTable,
        [string] $Name,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Skip',
        [RGBColors] $TabColor = [RGBColors]::None
    )
    $ScriptBlock = {
        Param (
            $DataTable,
            $TabColor,
            $Supress,
            $Option,
            $ExcelDocument,
            $Name
        )
        $addExcelWorkSheetDataSplat = @{
            DataTable          = $DataTable
            TabColor           = $TabColor
            Supress            = $Supress
            Option             = $Option
            ExcelDocument      = $ExcelDocument
            ExcelWorksheetName = $Name
        }
        Add-ExcelWorkSheetData @addExcelWorkSheetDataSplat
    }
    $ExcelWorkSheetParameters = [ordered] @{
        DataTable     = $DataTable
        TabColor      = $TabColor
        Supress       = $true
        Option        = $Option
        ExcelDocument = $Script:Excel.ExcelDocument
        Name          = $Name
    }

    if ($Script:Excel.Runspaces.Parallel) {
        $RunSpace = Start-Runspace -ScriptBlock $ScriptBlock -Parameters $ExcelWorkSheetParameters -RunspacePool $Script:Excel.Runspaces.RunspacesPool -Verbose:$Verbose
        $Script:Excel.Runspaces.Runspaces.Add($RunSpace)
    } else {
        & $ScriptBlock -Parameters @ExcelWorkSheetParameters
    }
}