#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName     = 'Test (Brecht)'
        MailTo         = 'bob@conotoso.com'
        DownloadFolder = New-Item 'TestDrive:/folder' -ItemType Directory
        LogFolder      = New-Item 'TestDrive:/log' -ItemType Directory
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
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('DownloadFolder', 'ScriptName') {
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
    It 'the download folder does not exist' {
        $testNewParams = $testParams.clone()
        $testNewParams.DownloadFolder = 'c:/notExistingFolder'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Download folder '$($testNewParams.DownloadFolder)' not found*")
        }
    }
    It 'authentication to the SFTP server fails' {
        Mock New-SFTPSession {
            throw 'Failed authenticating'
        }

        $testNewParams = $testParams.clone()
        $testNewParams.Sftp = @{
            Credential   = @{
                UserName = $env:ATTENTIA_SFTP_USERNAME_TEST
                Password = $env:ATTENTIA_SFTP_PASSWORD_TEST
            }
            ComputerName = 'ftp.somewhere'
            Path         = '/folder'
        }

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Failed creating an SFTP session to 'ftp.somewhere'*")
        }
    }
    It 'the SFTP path does not exist' {
        Mock Test-SFTPPath {
            $false
        }

        $testNewParams = $testParams.clone()
        $testNewParams.Sftp = @{
            Credential   = @{
                UserName = $env:ATTENTIA_SFTP_USERNAME_TEST
                Password = $env:ATTENTIA_SFTP_PASSWORD_TEST
            }
            ComputerName = 'ftp.somewhere'
            Path         = '/notExisting'
        }

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*SFTP path '/notExisting' not found*")
        }
    }
    It 'the SFTP file list could bot be retrieved' {
        Mock Get-SFTPChildItem {
            throw 'Failed getting list'
        }

        $testNewParams = $testParams.clone()
        $testNewParams.Sftp = @{
            Credential   = @{
                UserName = $env:ATTENTIA_SFTP_USERNAME_TEST
                Password = $env:ATTENTIA_SFTP_PASSWORD_TEST
            }
            ComputerName = 'ftp.somewhere'
            Path         = '/folder'
        }
      
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Failed retrieving the SFTP file list from 'ftp.somewhere' in path '/folder'*")
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testData = @(
            [PSCustomObject]@{
                Name          = '123456Brussels.txt'
                FullName      = '\folder\123456Brussels.txt'
                LastWriteTime = Get-Date
                Destination   = @{
                    Folder   = Join-Path $testParams.DownloadFolder 'Brussels'
                    FilePath = Join-Path $testParams.DownloadFolder 'Brussels\123456Brussels.txt'
                }
            }
            [PSCustomObject]@{
                Name          = '123456London.txt'
                FullName      = '\folder\123456London.txt'
                LastWriteTime = Get-Date
                Destination   = @{
                    Folder   = Join-Path $testParams.DownloadFolder 'London'
                    FilePath = Join-Path $testParams.DownloadFolder 'London\123456London.txt'
                }
            }
        )
        Mock Get-SFTPChildItem {
            $testData | Select-Object -Property * -ExcludeProperty 'Destination'
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[0].Destination.FilePath
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[0].FullName) -and
            ($Destination -eq $testData[0].Destination.Folder) -and
            ($Force)
        }
        Mock Get-SFTPItem {
            $null = New-Item -Path $testData[1].Destination.FilePath
        } -ParameterFilter {
            ($SessionId) -and
            ($Path -eq $testData[1].FullName) -and
            ($Destination -eq $testData[1].Destination.Folder) -and
            ($Force)
        }

        $testNewParams = $testParams.clone()

        .$testScript @testNewParams
    }
    It 'the SFTP file list is retrieved' {
        Should -Invoke Get-SFTPChildItem -Times 1 -Exactly -Scope Describe
    }
    It 'a folder is created based on the file name' {
        $testData[0].Destination.Folder | Should -Exist
        $testData[1].Destination.Folder | Should -Exist
    }
    It 'the files are downloaded to the correct folder' {
        Should -Invoke Get-SFTPItem -Times 2 -Exactly -Scope Describe

        $testData[0].Destination.FilePath | Should -Exist
        $testData[1].Destination.FilePath | Should -Exist
    }
    It 'the SFTP session is closed' {
        Should -Invoke Remove-SFTPSession -Times 1 -Exactly -Scope Describe
    }
} -Tag test