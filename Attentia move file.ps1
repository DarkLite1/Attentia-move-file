#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, ImportExcel

<#
.SYNOPSIS
    Move files to a specific folder, based on the file name and a mapping table.

.DESCRIPTION
    Move files from the source folder to the destination folder. Subfolders
    will be created in the destination folder based on the file name and a
    mapping table.

.PARAMETER ImportFile
    A .JSON file that contains all the parameters used by the script.

.PARAMETER MailTo
    E-mail addresses of where to send the summary e-mail

.PARAMETER OverwriteFile
    Overwrite a file when a file with the same name already exists in the
    destination subfolder.

.PARAMETER SourceFolder
    The source folder where the original files are stored.

.PARAMETER DestinationFolder
    The parent folder where the subfolders are created in which the moved
    files will be saved.

.PARAMETER ChildFolderNameMappingTable
    When a file name has a matching company code and location code, a subfolder
    is created in the destination folder where the file will be stored.
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
                'SourceFolder', 'DestinationFolder', 'ChildFolderNameMappingTable',
                'ExportExcelFile', 'SendMail',
                'Option'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

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

            #region Test ChildFolderNameMappingTable
            foreach ($f in $file.ChildFolderNameMappingTable) {
                @('FolderName', 'CompanyCode', 'LocationCode') |
                Where-Object { -not $f.$_ } | ForEach-Object {
                    throw "Property '$_' with value '' in the 'ChildFolderNameMappingTable' is not valid."
                }
            }
            $duplicateChildFolderNameMappingTable = $file.ChildFolderNameMappingTable |
            Group-Object -Property {
                '{0} - {1}' -f $_.CompanyCode, $_.LocationCode
            } | Where-Object {
                $_.Count -ge 2
            }

            if ($duplicateChildFolderNameMappingTable) {
                throw "Property 'ChildFolderNameMappingTable' contains a duplicate combination of CompanyCode and LocationCode: {0}" -f ($duplicateChildFolderNameMappingTable.Name -join ', ')
            }
            #endregion

            #region Test boolean
            try {
                [Boolean]::Parse($file.Option.OverwriteFile)
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

        #region Test folders
        if (-not (Test-Path -Path $file.SourceFolder -PathType 'Container')) {
            throw "Source folder '$($file.SourceFolder)' not found."
        }

        if (-not (Test-Path -Path $file.DestinationFolder -PathType 'Container')) {
            throw "Destination folder '$($file.DestinationFolder)' not found."
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

        #region Create Excel worksheet Overview
        if ($results) {
            $excelParams.WorksheetName = 'Overview'
            $excelParams.TableName = 'Overview'

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $results.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $results | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Create Excel worksheet FolderNameMappingTable
        if ($Download.ChildFolderNameMappingTable.Count -ne 0) {
            $excelParams.WorksheetName = 'FolderNameMappingTable'
            $excelParams.TableName = 'FolderNameMappingTable'

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $Download.ChildFolderNameMappingTable.Count,
            $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $Download.ChildFolderNameMappingTable | Sort-Object 'FolderName' |
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Send mail to user

        #region Error counters
        $counter = @{
            FilesOnServer       = $results.Count
            FilesDownloaded     = $results.Where({ $_.DownloadedOn }).Count
            DownloadErrors      = $results.Where({ $_.Error }).Count
            RemovedOnSftpServer = $results.Where({ $_.RemovedOnSftpServer }).Count
            SystemErrors        = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0}/{1} file{2} downloaded' -f
        $counter.FilesDownloaded, $counter.FilesOnServer,
        $(
            if ($counter.FilesOnServer -ne 1) { 's' }
        )

        if (
            $totalErrorCount = $counter.DownloadErrors + $counter.SystemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create html lists
        $systemErrorsHtmlList = if ($counter.SystemErrors) {
            "<p>Detected <b>{0} non terminating error{1}</b>:{2}</p>" -f $counter.SystemErrors,
            $(
                if ($counter.SystemErrors -ne 1) { 's' }
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
                    <td>SFTP files</td>
                    <td>$($counter.FilesOnServer)</td>
                </tr>
                <tr>
                    <td>Files downloaded</td>
                    <td>$($counter.FilesDownloaded)</td>
                </tr>
                <tr>
                    <td>Files removed on SFTP server</td>
                    <td>$($counter.RemovedOnSftpServer)</td>
                </tr>
                <tr>
                    <td>Errors</td>
                    <td>$($counter.DownloadErrors)</td>
                </tr>
                <tr>
                    <th colspan=`"2`">Parameters</th>
                </tr>
                <tr>
                    <td>SFTP hostname</td>
                    <td>$($Sftp.ComputerName)</td>
                </tr>
                <tr>
                    <td>SFTP path</td>
                    <td>$($Sftp.Path)</td>
                </tr>
                <tr>
                    <td>Download folder</td>
                    <td><a href=`"$($Download.Path)`">$($($Download.Path))</a></td>
                </tr>
                <tr>
                    <td>Overwrite downloaded files</td>
                    <td>$($Download.OverwriteExistingFile)</td>
                </tr>
                <tr>
                    <td>Remove files from SFTP server</td>
                    <td>$($Sftp.RemoveFileAfterDownload)</td>
                </tr>
            </table>
        "
        #endregion

        $mailParams += @{
            To        = $file.SendMail.To
            Bcc       = $ScriptAdmin
            Message   = "
                        $systemErrorsHtmlList
                        <p>Download files from an SFTP server.</p>
                        $summaryHtmlTable"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $logFile + ' - Mail.html'
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