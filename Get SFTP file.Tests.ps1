#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.General
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName     = 'Test (Brecht)'
        DownloadFolder = New-Item 'TestDrive:/folder' -ItemType Directory
        LogFolder      = New-Item 'TestDrive:/log' -ItemType Directory
    }

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
            ($Message -like  "*Download folder '$($testNewParams.DownloadFolder)' not found*")
        }
    }
}