#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Posh-SSH

<#
.SYNOPSIS
    Get files from an SFTP server.

.DESCRIPTION
    Retrieve a single file or multiple files from an SFTP server and save them
    in the destination folder. 

.PARAMETER UserName
    The user name used to authenticate to the SFTP server.

.PARAMETER Password
    The password used to authenticate to the SFTP server.

.PARAMETER DownloadFolder
    The destination folder where the file will be saved.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$DownloadFolder,
    [Parameter(Mandatory)]
    [String[]]$MailTo, 
    [HashTable]$Sftp = @{
        Credential   = @{
            UserName = $env:ATTENTIA_SFTP_USERNAME_TEST
            Password = $env:ATTENTIA_SFTP_PASSWORD_TEST
        }
        ComputerName = 'ftp.attentia.be'
        Path         = '/Out/BAND'
    },
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

        #region Test download folder
        if (-not (Test-Path -Path $DownloadFolder -PathType 'Container')) { 
            throw "Download folder '$DownloadFolder' not found."
        }
        #endregion

        #region Create SFTP credential
        try {
            $M = 'Create SFTP credential'
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $params = @{
                String      = $Sftp.Credential.Password 
                AsPlainText = $true
                Force       = $true
            }
            $secureStringPassword = ConvertTo-SecureString @params

            $params = @{
                TypeName     = 'System.Management.Automation.PSCredential'
                ArgumentList = $Sftp.Credential.UserName, $secureStringPassword
                ErrorAction  = 'Stop'
            }
            $sftpCredential = New-Object @params
        }
        catch {
            throw "Failed creating the SFTP credential with user name '$($Sftp.Credential.UserName)' and password '$($Sftp.Credential.Password)': $_"
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
        #region Open SFTP session
        try {
            $M = 'Start SFTP session'
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message  $M

            $params = @{
                ComputerName = $Sftp.ComputerName
                Credential   = $sftpCredential
                AcceptKey    = $true
                ErrorAction  = 'Stop'
            }
            $sftpSession = New-SFTPSession @params
        }
        catch {
            throw "Failed creating an SFTP session to '$($Sftp.ComputerName)': $_"
        }
        #endregion

        $sftpSessionParams = @{
            SessionId   = $sftpSession.SessionID
            Path        = $Sftp.Path
            ErrorAction = 'Stop'
        }

        #region Test SFTP path
        if (-not (Test-SFTPPath @sftpSessionParams)) {
            throw "SFTP path '$($Sftp.Path)' not found"
        }    
        #endregion
        
        #region Get SFTP file list
        try {
            $M = "Get SFTP file list in path '{0}'" -f $Sftp.Path
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $sftpFiles = Get-SFTPChildItem @sftpSessionParams
        }
        catch {
            throw "Failed retrieving the SFTP file list from '$($Sftp.ComputerName)' in path '$($Sftp.Path)': $_"
        }
        #endregion
  
        #region Get SFTP files
        $M = 'Download SFTP files'
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $fileDownloadFolders = @{}

        $results = ForEach ($sftpFile in $sftpFiles) {
            try {
                $result = [PSCustomObject]@{
                    FileName          = $sftpFile.Name
                    FileLastWriteTime = $sftpFile.LastWriteTime 
                    Downloaded        = $false
                    DownloadedOn      = $null
                    DownloadFolder    = $null
                    Error             = $null
                }
                
                #region Create file download folder
                $folderName = $result.FileName.SubString(
                    6, $result.FileName.Length - 10
                )

                if (-not $fileDownloadFolders.ContainsKey($folderName)) {
                    try {
                        $testPathParams = @{
                            Path        = Join-Path $DownloadFolder $folderName
                            PathType    = 'Container'
                            ErrorAction = 'Stop'
                        }

                        if (-not (Test-Path @testPathParams)) { 
                            $M = "Create file download folder '{0}'" -f
                            $testPathParams.Path
                            Write-Verbose $M
                            Write-EventLog @EventVerboseParams -Message $M
                
                            $newItemParams = @{
                                Path        = $testPathParams.Path
                                ItemType    = 'Directory'
                                ErrorAction = 'Stop'
                            }
                            $null = New-Item @newItemParams

                            $fileDownloadFolders[$folderName] = $testPathParams.Path
                        }
                    }
                    catch {
                        $M = "Failed creating file download folder '{0}': $_" -f $testPathParams.Path
                        throw $M
                    }
                }
                
                $result.DownloadFolder = $fileDownloadFolders[$folderName]
                #endregion

                #region Download SFTP file to correct folder
                try {
                    $M = "Download SFTP file '{0}' to folder '{1}'" -f 
                    $result.FileName, $result.DownloadFolder
                    Write-Verbose $M
    
                    $params = @{
                        SessionId   = $sftpSession.SessionID
                        Path        = $sftpFile.FullName 
                        Destination = $result.DownloadFolder 
                        Force       = $true
                        ErrorAction = 'Stop'
                    }
                    Get-SFTPItem @params
    
                    $result.DownloadedOn = Get-Date
                    $result.Downloaded = $true    
                }
                catch {
                    throw "Failed downloading file '$($result.FileName)': $_"
                }
                #endregion
            }
            catch {
                Write-Warning $_
                $result.Error = $_
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }  
        #endregion
  
        #region Close SFTP session
        $M = 'Close SFTP session'
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        Remove-SFTPSession -SessionId $sftpSession.SessionID -ErrorAction Ignore
        #endregion
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