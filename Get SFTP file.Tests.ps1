#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        MailTo   = 'bob@contoso.com'
        Download = @{
            OverwriteExistingFile       = $false
            ParentFolder                = (New-Item 'TestDrive:/a' -ItemType Directory).FullName
            ChildFolderNameMappingTable = @(
                @{
                    FolderName   = 'Brussels'
                    CompanyCode  = '577600'
                    LocationCode = '057'
                }
                @{
                    FolderName   = 'London'
                    CompanyCode  = '577601'
                    LocationCode = '057'
                }
            )
        }
        Sftp     = @{
            Credential              = @{
                UserName = 'envVarBob'
                Password = 'envVarPasswordBob'
            }
            ComputerName            = 'PC1'
            Path                    = '/folder'
            RemoveFileAfterDownload = $false
        }
    }

    $testData = @(
        [PSCustomObject]@{
            Name          = 'BAND_577600_A_057_202306301556.pdf'
            FullName      = '\folder\BAND_577600_A_057_202306301556.pdf'
            LastWriteTime = (Get-Date).AddDays(-3)
            Destination   = @{
                Folder   = Join-Path $testInputFile.Download.ParentFolder 'Brussels'
                FilePath = Join-Path $testInputFile.Download.ParentFolder 'Brussels\BAND_577600_A_057_202306301556.pdf'
            }
        }
        [PSCustomObject]@{
            Name          = 'BAND_999900_A_123_202307301544.pdf'
            FullName      = '\folder\BAND_999900_A_123_202307301544.pdf'
            LastWriteTime = (Get-Date).AddDays(-4)
            Destination   = @{
                Folder   = Join-Path $testInputFile.Download.ParentFolder '999900 123'
                FilePath = Join-Path $testInputFile.Download.ParentFolder '999900 123\BAND_999900_A_123_202307301544.pdf'
            }
        }
    )

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@conotoso.com'
    }

    Function Get-EnvironmentVariableValueHC {
        Param(
            [String]$Name
        )
    }
    
    Mock Get-EnvironmentVariableValueHC {
        'bob'
    } -ParameterFilter {
        $Name -eq $testInputFile.SFtp.Credential.UserName
    }
    Mock Get-EnvironmentVariableValueHC {
        'PasswordBob'
    } -ParameterFilter {
        $Name -eq $testInputFile.SFtp.Credential.Password
    }
    Mock Get-SFTPChildItem
    Mock Get-SFTPItem
    Mock New-SFTPSession {
        [PSCustomObject]@{
            SessionID = 1
        }
    }
    Mock Test-SFTPPath {
        $true
    }
    Mock Remove-SFTPSession
    Mock Remove-SFTPItem
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To[0] -eq $ScriptAdmin[0]) -and 
            ($To[1] -eq $ScriptAdmin[1]) -and 
            ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It 'RemoveFileAfterDownload is not a boolean' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Sftp.RemoveFileAfterDownload = 2

                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*$ImportFile*Property 'RemoveFileAfterDownload' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It '<_> is missing' -ForEach @(
                'MailTo', 'Download'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*$ImportFile*No '$_' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'Download' {
                It 'OverwriteExistingFile is not a boolean' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Download.OverwriteExistingFile = 2
    
                    $testNewInputFile | ConvertTo-Json -Depth 5 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'OverwriteExistingFile' is not a boolean value*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It '<_> is missing' -ForEach @(
                    'ParentFolder', 
                    'ChildFolderNameMappingTable'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Download.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 5 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*No '$_' found in 'Download'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                Context 'ChildFolderNameMappingTable' {
                    It '<_> is missing' -ForEach @(
                        'FolderName', 'CompanyCode', 'LocationCode'
                    ) {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Download.ChildFolderNameMappingTable[0].$_ = $null
        
                        $testNewInputFile | ConvertTo-Json -Depth 5 | 
                        Out-File @testOutParams
                        
                        .$testScript @testParams
                        
                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and 
                            ($Message -like "*$ImportFile*No '$_' found in the 'Download.ChildFolderNameMappingTable'*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
            }
            Context 'sftp' {
                It 'credential.<_> is missing' -ForEach @(
                    'UserName', 'Password'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Sftp.Credential.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 5 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*No '$_' found in 'sftp.credential'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It '<_> is missing' -ForEach @(
                    'ComputerName', 'Path'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Sftp.$_ = $null
    
                    $testNewInputFile | ConvertTo-Json -Depth 5 | 
                    Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*No '$_' found in 'sftp'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            it 'ChildFolderNameMappingTable contains duplicates' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Download.ChildFolderNameMappingTable = @(
                    @{
                        FolderName   = 'Brussels'
                        CompanyCode  = '577600'
                        LocationCode = '057'
                    }
                    @{
                        FolderName   = 'Genk'
                        CompanyCode  = '577600'
                        LocationCode = '057'
                    }
                )

                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*$ImportFile*Property 'ChildFolderNameMappingTable' contains a duplicate combination of CompanyCode and LocationCode*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
    It 'the parent download folder does not exist' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Download.ParentFolder = 'c:/notExistingFolder'

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Parent download folder '$($testNewInputFile.Download.ParentFolder)' not found*")
        }
    }
    It 'authentication to the SFTP server fails' {
        Mock New-SFTPSession {
            throw 'Failed authenticating'
        }

        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Failed creating an SFTP session to '$($testInputFile.sftp.ComputerName)'*")
        }
    }
    It 'the SFTP path does not exist' {
        Mock Test-SFTPPath {
            $false
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Sftp.Path = '/notExisting'

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*SFTP path '$($testNewInputFile.Sftp.Path)' not found*")
        }
    }
    It 'the SFTP file list could bot be retrieved' {
        Mock Get-SFTPChildItem {
            throw 'Failed getting list'
        }

        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Failed retrieving the SFTP file list from '$($testInputFile.Sftp.ComputerName)' in path '$($testInputFile.sftp.Path)'*")
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        Mock Get-SFTPChildItem {
            $testData | Select-Object -Property * -ExcludeProperty 'Destination'
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[0].Destination.FilePath
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[0].FullName) -and
            ($Destination -eq $testData[0].Destination.Folder)
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[1].Destination.FilePath
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[1].FullName) -and
            ($Destination -eq $testData[1].Destination.Folder)
        }

        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams
    }
    It 'the SFTP file list is retrieved' {
        Should -Invoke Get-SFTPChildItem -Times 1 -Exactly -Scope Describe
    }
    Context 'download each file on the SFTP server to' {
        It 'the matching FolderName in ChildFolderNameMappingTable' {
            $testData[0].Destination.Folder | Should -Exist
            $testData[0].Destination.FilePath | Should -Exist

            Should -Invoke Get-SFTPItem -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
                ($Path -eq $testData[0].FullName) -and
                ($Destination -eq $testData[0].Destination.Folder)
            }
        }
        It "the folder 'CompanyCode LocationCode' when there is no match" {
            $testData[1].Destination.Folder | Should -Exist
            $testData[1].Destination.FilePath | Should -Exist

            Should -Invoke Get-SFTPItem -Times 1 -Exactly -Scope 'Describe' -ParameterFilter {
                ($Path -eq $testData[1].FullName) -and
                ($Destination -eq $testData[1].Destination.Folder)
            }
        }
    }
    It 'the SFTP session is closed' {
        Should -Invoke Remove-SFTPSession -Times 1 -Exactly -Scope Describe
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    FileName          = $testData[0].Name
                    FileLastWriteTime = $testData[0].LastWriteTime
                    DownloadedOn      = Get-Date
                    DownloadFolder    = $testData[0].Destination.Folder
                    Error             = $null
                }
                @{
                    FileName          = $testData[1].Name
                    FileLastWriteTime = $testData[1].LastWriteTime
                    DownloadedOn      = Get-Date
                    DownloadFolder    = $testData[1].Destination.Folder
                    Error             = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.FileName -eq $testRow.FileName
                }
                $actualRow.FileLastWriteTime.ToString('yyyyMMdd HHmmss') | 
                Should -Be $testRow.FileLastWriteTime.ToString('yyyyMMdd HHmmss')
                $actualRow.DownloadedOn.ToString('yyyyMMdd') | 
                Should -Be $testRow.DownloadedOn.ToString('yyyyMMdd')
                $actualRow.DownloadFolder | Should -Be $testRow.DownloadFolder
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'send an e-mail' {
        It 'to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '2/2 files downloaded') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*table*SFTP files*2*Files downloaded*2*Errors*0*")
            }
        }
    }
}
Describe 'when RemoveFileAfterDownload is' {
    BeforeAll {
        Mock Get-SFTPChildItem {
            $testData | Select-Object -Property * -ExcludeProperty 'Destination'
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[0].Destination.FilePath -Force
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[0].FullName) -and
            ($Destination -eq $testData[0].Destination.Folder) -and
            ($Force)
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[1].Destination.FilePath -Force
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[1].FullName) -and
            ($Destination -eq $testData[1].Destination.Folder) -and
            ($Force)
        }
    }
    It 'true the file on the SFTP server is removed' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Sftp.RemoveFileAfterDownload = $true

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Remove-SFTPItem -Times 2 -Exactly
    }
    It 'false the file is left untouched on the SFTP server' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Sftp.RemoveFileAfterDownload = $false

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Not -Invoke Remove-SFTPItem
    }
}
Describe 'when OverwriteExistingFile is' {
    BeforeAll {
        Mock Get-SFTPChildItem {
            $testData | Select-Object -Property * -ExcludeProperty 'Destination'
        }
    }
    It 'true the file in the download folder is overwritten' {
        $null = New-Item -Path $testData[0].Destination.FilePath -Force
        $null = New-Item -Path $testData[1].Destination.FilePath -Force

        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[0].Destination.FilePath -Force
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[0].FullName) -and
            ($Destination -eq $testData[0].Destination.Folder) -and
            ($Force)
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[1].Destination.FilePath -Force
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[1].FullName) -and
            ($Destination -eq $testData[1].Destination.Folder) -and
            ($Force)
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Download.OverwriteExistingFile = $true

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Get-SFTPItem -Times 2 -Exactly -ParameterFilter {
            $Force
        }

    }
    It 'false the file in the parent folder if now overwritten and an error is logged' {
        $null = New-Item -Path $testData[0].Destination.FilePath -Force
        $null = New-Item -Path $testData[1].Destination.FilePath -Force

        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[0].Destination.FilePath
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[0].FullName) -and
            ($Destination -eq $testData[0].Destination.Folder)
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[1].Destination.FilePath
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[1].FullName) -and
            ($Destination -eq $testData[1].Destination.Folder)
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Download.OverwriteExistingFile = $false

        $testNewInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Get-SFTPItem -Times 2 -Exactly -ParameterFilter {
            -not $Force
        }
    }
}