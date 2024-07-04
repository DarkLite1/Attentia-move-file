#Requires -Version 7
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, ImportExcel

<#
.SYNOPSIS
    Move files to a specific folder, based on the file name and a mapping table.

.DESCRIPTION
    Move files from the source folder to the destination folder. The destination
    folder is derived from the matching the file name with the CompanyCode and
    LocationCode in the Destination parameter.

.PARAMETER ImportFile
    A .JSON file that contains all the parameters used by the script.

.PARAMETER MailTo
    E-mail addresses of where to send the summary e-mail

.PARAMETER OverwriteFile
    Overwrite a file when a file with the same name already exists in the
    destination folder.

.PARAMETER SourceFolder
    Folder where the original files are stored.

.PARAMETER NoMatchFolder
    Files that cannot be matched with a name in the Destination parameter will
    be moved to this folder. However, when NoMatchFolder is blank the
    non-matched files will not be moved and will remain in the source folder.

.PARAMETER Destination
    When a file name has a matching company code and location code it is moved
    to the correct destination folder.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\Attentia\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Error.Clear()

        #region Create log folder
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        try {
            @(
                'SourceFolder', 'Destination',
                'ExportExcelFile', 'SendMail',
                'Option'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            #region Test Destination mapping table
            foreach ($f in $file.Destination) {
                @('Folder', 'CompanyCode', 'LocationCode') |
                Where-Object { -not $f.$_ } | ForEach-Object {
                    throw "Property 'Destination.$_' not found."
                }
            }
            $duplicateChildFolderNameMappingTable = $file.Destination |
            Group-Object -Property {
                '{0} - {1}' -f $_.CompanyCode, $_.LocationCode
            } | Where-Object {
                $_.Count -ge 2
            }

            if ($duplicateChildFolderNameMappingTable) {
                throw "Property 'Destination' contains a duplicate combination of CompanyCode and LocationCode: {0}" -f ($duplicateChildFolderNameMappingTable.Name -join ', ')
            }
            #endregion

            #region Test SendMail and ExportExcelFile
            @('To', 'When').Where(
                { -not $file.SendMail.$_ }
            ).foreach(
                { throw "Property 'SendMail.$_' not found" }
            )

            @('When').Where(
                { -not $file.ExportExcelFile.$_ }
            ).foreach(
                { throw "Property 'ExportExcelFile.$_' not found" }
            )

            if ($file.SendMail.When -notMatch '^Never$|^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'SendMail.When' with value '$($file.SendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
            }

            if ($file.ExportExcelFile.When -notMatch '^Never$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'ExportExcelFile.When' with value '$($file.ExportExcelFile.When)' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
            }
            #endregion

            #region Test boolean
            try {
                $null = [Boolean]::Parse($file.Option.OverwriteFile)
            }
            catch {
                throw "Property 'Option.OverwriteFile' is not a boolean value"
            }
            #endregion
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Test source folder exits
        if (-not (Test-Path -Path $file.SourceFolder -PathType 'Container')) {
            throw "Source folder '$($file.SourceFolder)' not found."
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $sourceFiles = Get-ChildItem -LiteralPath $file.SourceFolder -File

        if (-not $sourceFiles) {
            Write-Verbose "No files found in folder '$($file.SourceFolder)'"
            Write-EventLog @EventEndParams; Exit
        }

        $results = foreach ($sourceFile in $sourceFiles) {
            try {
                Write-Verbose "File '$sourceFile'"

                $result = [PSCustomObject]@{
                    SourceFile        = $sourceFile
                    DestinationFolder = $null
                    CompanyCode       = $null
                    LocationCode      = $null
                    DateTime          = Get-Date
                    Moved             = $false
                    Action            = @()
                    Error             = $null
                }

                #region Get companyCode and locationCode from file name
                $tmpStrings = $sourceFile.Name.Split('_')

                if (-not ($result.CompanyCode = $tmpStrings[0])) {
                    throw 'No company code found in the file name'
                }

                Write-Verbose "CompanyCode '$($result.CompanyCode)'"

                if (-not ($result.LocationCode = $tmpStrings[2])) {
                    throw 'No location code found in the file name'
                }

                Write-Verbose "LocationCode '$($result.LocationCode)'"
                #endregion

                #region Get destination folder from mapping table
                $result.DestinationFolder =
                ($file.Destination.Where(
                    {
                        ($_.LocationCode -eq $result.LocationCode) -and
                        ($_.CompanyCode -eq $result.CompanyCode)
                    }, 'First'
                )
                ).Folder

                if (-not $result.DestinationFolder) {
                    if (-not $file.NoMatchFolderName) {
                        $errorMessage = "no match with LocationCode '$($result.LocationCode)' and CompanyCode '$($result.CompanyCode)', and MoMatchFolderName is blank"

                        $result.Error += "File not moved, $errorMessage"

                        $M = "File '$($result.SourceFile)' not moved,  $errorMessage"
                        Write-Warning $M
                        Write-EventLog @EventWarnParams -Message $M
                        Continue
                    }
                    $result.DestinationFolder = $file.NoMatchFolderName
                }
                #endregion

                #region Create destination folder
                $testPathParams = @{
                    Path        = $result.DestinationFolder
                    PathType    = 'Container'
                    ErrorAction = 'Stop'
                }

                if (-not (Test-Path @testPathParams)) {
                    try {
                        $M = "Create destination folder '{0}'" -f
                        $testPathParams.Path
                        Write-Verbose $M
                        Write-EventLog @EventVerboseParams -Message $M

                        $newItemParams = @{
                            Path        = $testPathParams.Path
                            ItemType    = 'Directory'
                            ErrorAction = 'Stop'
                        }
                        $null = New-Item @newItemParams

                        $result.Action += 'created destination folder'
                    }
                    catch {
                        $M = "Failed creating destination folder '{0}': $_" -f
                        $testPathParams.Path
                        $Error.RemoveAt(0)
                        throw $M
                    }
                }
                #endregion

                #region Move file
                $moveParams = @{
                    LiteralPath = $result.SourceFile.FullName
                    Destination = $result.DestinationFolder
                    ErrorAction = 'Stop'
                }

                if ($file.Option.OverwriteFile) {
                    $moveParams.Force = $true
                }

                Write-Verbose "Move file '$($moveParams.LiteralPath)' to '$($moveParams.Destination)'"

                Move-Item @moveParams

                $result.Action += 'file moved'
                $result.Moved = $true
                #endregion
            }
            catch {
                $M = "Failed moving file '$($sourceFile.Name)': $_"
                Write-Warning $M
                $result.Error = $_
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try {
        $mailParams = @{}

        $excelParams = @{
            Path         = $logFile + ' - Log.xlsx'
            AutoSize     = $true
            FreezeTopRow = $true
        }

        #region Counters
        $counter = @{
            SourceFiles = $results.Count
            FilesMoved  = $results.Where({ $_.Moved }).Count
            Actions     = $results.Where({ $_.Action }).Count
            Errors      = @{
                Move   = $results.Where({ $_.Error }).Count
                System = (
                    $Error.Exception.Message | Measure-Object
                ).Count
            }
        }

        $counter.TotalErrors = $counter.Errors.Move + $counter.Errors.System
        #endregion

        #region Create Excel worksheet Overview
        $createExcelFile = $false

        if (
            (
                ($file.ExportExcelFile.When -eq 'OnlyOnError') -and
                ($counter.TotalErrors)
            ) -or
            (
                ($file.ExportExcelFile.When -eq 'OnlyOnErrorOrAction') -and
                (($counter.TotalErrors) -or ($counter.Actions))
            )
        ) {
            $createExcelFile = $true
        }
        #endregion

        #region Create Excel worksheet Overview
        if ($createExcelFile) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Overview'

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $results.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $exportToExcel = $results |
            Select-Object -ExcludeProperty 'Action' -Property 'DateTime',
            @{
                Name       = 'SourceFolder'
                Expression = { $file.SourceFolder }
            },
            'DestinationFolder',
            @{
                Name       = 'FileName'
                Expression = {
                    $_.SourceFile.Name
                }
            },
            @{
                Name       = 'CompanyCode'
                Expression = {
                    $_.CompanyCode
                }
            },
            @{
                Name       = 'LocationCode'
                Expression = {
                    $_.LocationCode
                }
            },
            @{
                Name       = 'Successful'
                Expression = { $_.Moved }
            },
            @{
                Name       = 'Action'
                Expression = { $_.Action -join ', ' }
            },
            @{
                Name       = 'Error'
                Expression = { $_.Error -join ', ' }
            }

            $exportToExcel | Export-Excel @excelParams -NoNumberConversion 'CompanyCode', 'LocationCode'

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Create Excel worksheet MappingTable
        if ($createExcelFile) {
            $excelParams.WorksheetName = $excelParams.TableName = 'MappingTable'

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $file.Destination.Count,
            $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $file.Destination | Sort-Object 'Folder' |
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Send mail to user

        #region Check to send mail to user
        $sendMailToUser = $false

        if (
            (
                ($file.SendMail.When -eq 'Always')
            ) -or
            (
                ($file.SendMail.When -eq 'OnlyOnError') -and
                ($counter.TotalErrors)
            ) -or
            (
                ($file.SendMail.When -eq 'OnlyOnErrorOrAction') -and
                (($counter.Actions) -or ($counter.TotalErrors))
            )
        ) {
            $sendMailToUser = $true
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0}/{1} file{2} moved' -f
        $counter.FilesMoved, $counter.SourceFiles,
        $(
            if ($counter.SourceFiles -ne 1) { 's' }
        )

        if ($counter.TotalErrors) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", {0} error{1}" -f $counter.TotalErrors, $(
                if ($counter.TotalErrors -ne 1) { 's' }
            )
        }
        #endregion

        #region Create html lists
        $systemErrorsHtmlList = if ($counter.Errors.System) {
            "<p>Detected <b>{0} non terminating error{1}</b>:{2}</p>" -f $counter.Errors.System,
            $(
                if ($counter.Errors.System -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } |
                ConvertTo-HtmlListHC
            )
        }

        $summaryHtmlTable = "
            <table>
                <tr>
                    <th colspan=`"2`">Summary</th>
                </tr>
                <tr>
                    <td>Files in source folder</td>
                    <td>$($counter.SourceFiles)</td>
                </tr>
                <tr>
                    <td>Files moved</td>
                    <td>$($counter.FilesMoved)</td>
                </tr>
                $(
                    if ($counter.Errors.Move) {
                    "<tr>
                        <td>Errors</td>
                        <td>$($counter.Errors.Move)</td>
                    </tr>"
                    }
                )
            </table>
        "
        #endregion

        $mailParams += @{
            To             = $file.SendMail.To
            Bcc            = $ScriptAdmin
            Message        = "
                        $systemErrorsHtmlList
                        $summaryHtmlTable"
            LogFolder      = $LogParams.LogFolder
            Header         = $ScriptName
            EventLogSource = $ScriptName
            Save           = $LogFile + ' - Mail.html'
            ErrorAction    = 'Stop'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message +=
            "<p><i>* Check the attachment for details</i></p>"
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}