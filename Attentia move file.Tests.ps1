#Requires -Version 7
#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML

BeforeAll {
    $testInputFile = @{
        SourceFolder                = (New-Item 'TestDrive:/a' -ItemType Directory).FullName
        DestinationFolder           = (New-Item 'TestDrive:/b' -ItemType Directory).FullName
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
        Option                      = @{
            OverwriteFile = $false
        }
        SendMail                    = @{
            To   = @('bob@contoso.com')
            When = 'Always'
        }
        ExportExcelFile             = @{
            When = 'OnlyOnErrorOrAction'
        }
    }

    $testData = @(
        [PSCustomObject]@{
            Name          = 'BAND_577600_A_057_202306301556.pdf'
            FullName      = '\folder\BAND_577600_A_057_202306301556.pdf'
            LastWriteTime = (Get-Date).AddDays(-3)
            Destination   = @{
                Folder   = Join-Path $testInputFile.DestinationFolder 'Brussels'
                FilePath = Join-Path $testInputFile.DestinationFolder 'Brussels\BAND_577600_A_057_202306301556.pdf'
            }
        }
        [PSCustomObject]@{
            Name          = 'BAND_999900_A_123_202307301544.pdf'
            FullName      = '\folder\BAND_999900_A_123_202307301544.pdf'
            LastWriteTime = (Get-Date).AddDays(-4)
            Destination   = @{
                Folder   = Join-Path $testInputFile.DestinationFolder '999900 123'
                FilePath = Join-Path $testInputFile.DestinationFolder '999900 123\BAND_999900_A_123_202307301544.pdf'
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
            ($To -eq $testParams.ScriptAdmin) -and
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
            It '<_> not found' -ForEach @(
                'SourceFolder', 'DestinationFolder', 'ChildFolderNameMappingTable',
                'ExportExcelFile', 'SendMail',
                'Option'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ExportExcelFile.<_> not found' -ForEach @(
                'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ExportExcelFile.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.When = 'wrong'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.When = 'wrong'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Option.OverwriteFile is not a boolean' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Option.OverwriteFile = 2

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Option.OverwriteFile' is not a boolean value*")
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
                    $testNewInputFile.ChildFolderNameMappingTable[0].$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' with value '' in the 'ChildFolderNameMappingTable' is not valid*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            It 'ChildFolderNameMappingTable contains duplicates' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ChildFolderNameMappingTable = @(
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
    It 'the source folder does not exist' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.SourceFolder = 'c:/notExistingFolder'

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like "*Source folder '$($testNewInputFile.SourceFolder)' not found*")
        }
    }
    It 'the destination folder does not exist' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.DestinationFolder = 'c:/notExistingFolder'

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like "*Destination folder '$($testNewInputFile.DestinationFolder)' not found*")
        }
    }
} -Tag test
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