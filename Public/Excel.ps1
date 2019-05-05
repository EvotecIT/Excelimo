function Excel {
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Excel requires opening and closing brace."),
        [string] $FilePath,
        [switch] $Open
    )

    $ExcelDocument = New-ExcelDocument
    $Script:ExcelDocument = $ExcelDocument

    # We need to make sure some commands are executed in correct order, therefore we convert scriptblock into text, get the commands that we need to be executed last (Main)
    # and build ScriptBlock back
    [Array] $Output = Get-ASTCode -ScriptBlock $Content
    $Main = ConvertTo-ScriptBlock -Code $Output -Include 'WorkbookProperties'
    $Worksheets = ConvertTo-ScriptBlock -Code $Output -Include 'Worksheet'

    if ($Worksheets) {
        Invoke-Command -ScriptBlock $Worksheets
    }
    if ($Main) {
        Invoke-Command -ScriptBlock $Main
        <#
        [Array] $Parameters = Invoke-Command -ScriptBlock $Main
        foreach ($Parameter in $Parameters) {
            switch ( $Parameter.Type ) {
                WorkbookProperties {
                    $Splat = $Parameter.ExcelProperties
                    Set-ExcelProperties @Splat -ExcelDocument $ExcelDocument
                    break
                }
            }

        }
        #>
    }
    Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $FilePath -OpenWorkBook:$Open
}


function Get-ASTCode {
    param(
        [ScriptBlock] $ScriptBlock
    )
    [Array] $Output = $ScriptBlock.Ast.EndBlock.Statements.Extent
    [Array] $OutputText = foreach ($Line in $Output) {
        [string] $Line + [System.Environment]::NewLine
    }
    return $OutputText
}

function ConvertTo-ScriptBlock {
    param(
        [Array] $Code,
        [string[]] $Include,
        [string[]] $Exclude
    )
    if ($Include) {
        $Output = foreach ($Line in $Code) {
            foreach ($I in $Include) {
                if ($Line.StartsWith($I)) {
                    $Line
                }
            }
        }
    }
    if ($Exclude) {
        $Output = foreach ($Line in $Code) {
            $Tests = foreach ($E in $Exclude) {
                if ($Line.StartsWith($E)) {
                    $true
                }
            }
            if ($Tests -notcontains $true) {
                $Line
            }
        }
    }
    if ($Output) {
        [ScriptBlock]::Create($Output)
    }
}