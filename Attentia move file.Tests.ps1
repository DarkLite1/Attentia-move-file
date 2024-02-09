#Requires -Version 7
#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML

BeforeAll {
    $testInputFile = @{
        SourceFolder    = (New-Item 'TestDrive:\a' -ItemType Directory).FullName
        NoMatchFolder   = (New-Item 'TestDrive:\n' -ItemType Directory).FullName
        Destination     = @(
            @{
                FolderName   = 'TestDrive:\b\Brussels'
                CompanyCode  = '577600'
                LocationCode = '057'
            }
        )
        Option          = @{
            OverwriteFile = $false
        }
        SendMail        = @{
            To   = @('bob@contoso.com')
            When = 'Always'
        }
        ExportExcelFile = @{
            When = 'OnlyOnErrorOrAction'
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:\Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:\log' -ItemType Directory
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
        $mailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\\notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$mailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$mailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'SourceFolder', 'Destination',
                'ExportExcelFile', 'SendMail',
                'Option'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$mailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'Destination' {
                It '<_> not found' -ForEach @(
                    'FolderName', 'CompanyCode', 'LocationCode'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Destination[0].$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$mailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Destination.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'contains duplicates' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Destination = @(
                        @{
                            FolderName   = 'a'
                            CompanyCode  = '577600'
                            LocationCode = '057'
                        }
                        @{
                            FolderName   = 'b'
                            CompanyCode  = '577600'
                            LocationCode = '057'
                        }
                    )

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$mailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Destination' contains a duplicate combination of CompanyCode and LocationCode*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'SendMail' {
                It '<_> not found' -ForEach @(
                    'To', 'When'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.SendMail.$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$mailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'When is not valid' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.SendMail.When = 'wrong'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$mailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'ExportExcelFile' {
                It '<_> not found' -ForEach @(
                    'When'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.ExportExcelFile.$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$mailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'ExportExcelFile.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'When is not valid' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.ExportExcelFile.When = 'wrong'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$mailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            It 'Option.OverwriteFile is not a boolean' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Option.OverwriteFile = 2

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$mailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Option.OverwriteFile' is not a boolean value*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
    It 'the source folder does not exist' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.SourceFolder = 'c:\notExistingFolder'

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$mailAdminParams) -and
            ($Message -like "*Source folder '$($testNewInputFile.SourceFolder)' not found*")
        }
    }
}  -Tag test
Describe 'move files from the source folder to the destination folder' {
    It 'to the des' {

    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testData = @(

        )

        $testInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        $testFiles = $testData.ForEach(
            {
                $testNewItemParams = @{
                    Path     = (
                        $testInputFile.SourceFolder + '\' + $_.FileName)
                    ItemType = 'File'
                }
                New-Item @testNewItemParams
            }
        )

        .$testScript @testParams
    }
    Context 'the files are moved from the source to' {
        It 'a specific destination folder' {
            $testData.ForEach(
                { $_.Destination.FilePath | Should -Exist }
            )
            $testFiles.ForEach(
                { $_.FullName | Should -Not -Exist }
            )
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    DateTime     = Get-Date
                    SourceFolder = $testInputFile.SourceFolder
                    Destination  = $testData[0].Destination.Folder
                    FileName     = $testData[0].FileName
                    Successful   = $true
                    Action       = 'created destination folder, file moved'
                    Error        = ''
                }
                @{
                    DateTime     = Get-Date
                    SourceFolder = $testInputFile.SourceFolder
                    Destination  = $testData[1].Destination.Folder
                    FileName     = $testData[1].FileName
                    Successful   = $true
                    Action       = 'created destination folder, file moved'
                    Error        = ''
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
                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should -Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.SourceFolder | Should -Be $testRow.SourceFolder
                $actualRow.Destination.ParentFolder | Should -Be $testRow.Destination.ParentFolder
                $actualRow.Successful | Should -Be $testRow.Successful
                $actualRow.Action | Should -Be $testRow.Action
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
            ($Subject -eq '2\2 files downloaded') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*table*SFTP files*2*Files downloaded*2*Errors*0*")
            }
        }
    }
}
Context 'Option.OverwriteFile' {
    It 'if true the destination file is overwritten' {
        Mock Move-Item

        $testData.ForEach(
            {
                $testNewItemParams = @{
                    Path     = (
                        $testInputFile.SourceFolder + '\' + $_.FileName)
                    ItemType = 'File'
                }
                New-Item @testNewItemParams
            }
        )

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Option.OverwriteFile = $true

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Move-Item -Times $testData.Count -Exactly -ParameterFilter {
            ($Force)
        }
    }
}