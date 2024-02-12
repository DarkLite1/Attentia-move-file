<#
    .SYNOPSIS
        Update the Destination field in the input file with data from
        an Excel file.
#>
Param (
    [String]$ExcelFile = 'T:\Test\Brecht\PowerShell\Attentia mapping table.xlsx',
    [String]$TemplateFile = 'T:\Prod\Application specific\Attentia\Attentia move file\Example.json',
    [String]$NewTemplateFile = "$PSScriptRoot\NewTemplate.json"
)

try {
    #region Test input
    if (-not (Test-Path -LiteralPath $ExcelFile -PathType Leaf)) {
        throw "Excel file '$ExcelFile' not found"
    }

    if (-not (Test-Path -LiteralPath $TemplateFile -PathType Leaf)) {
        throw "Import template file '$TemplateFile' not found"
    }
    #endregion

    #region Import Excel and .json file
    $excelFileContent = Import-Excel -Path $ExcelFile

    $params = @{
        LiteralPath = $TemplateFile
        Raw         = $true
        Encoding    = 'UTF8'
        ErrorAction = 'Stop'
    }
    $templateFileContent = Get-Content @params | ConvertFrom-Json
    #endregion

    #region Add Excel data to input file Destination field
    $templateFileContent.Destination = $excelFileContent |
    Select-Object -Property 'Folder', 'CompanyCode', 'LocationCode' |
    Sort-Object 'Folder', 'CompanyCode', 'LocationCode'
    #endregion

    #region Create new input file
    $templateFileContent | ConvertTo-Json -Depth 7 |
    Out-File -FilePath $NewTemplateFile -Encoding utf8
    #endregion
}
catch {
    throw "Failed to create a new input file: $_"
}