#Requires -Modules ActiveDirectory, Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $OutParams = @{
        FilePath = (New-Item "TestDrive:/Test.txt" -ItemType File).FullName
        Encoding = 'utf8'
    }
    
    $testParams = @{
        ScriptName  = 'Test'
        ImportFile  = $OutParams.FilePath
        SQLDatabase = 'PowerShell TEST'
        LogFolder   = (New-Item "TestDrive:/log" -ItemType Directory).FullName
    }
    
    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }
    
    $ADOUExisting = Get-ADOrganizationalUnit -Filter * | 
    Select-Object -First 1 -ExpandProperty DistinguishedName

    Mock Invoke-Sqlcmd2
    Mock Send-MailHC
    Mock Write-EventLog
    Mock Get-ADDisplayNameHC
    Mock Get-ADDisplayNameFromSID
    Mock Get-ADCircularGroupsHC
    Mock Get-ADTSProfileHC
    Mock Test-ADOUExistsHC { $true }
}

Describe 'Prerequisites' {
    Context 'ImportFile' {
        It 'mandatory parameter' {
            (Get-Command $testScript).Parameters['ImportFile'].Attributes.Mandatory | 
            Should -BeTrue
        } 
        It 'file not existing' {
            $testNewParams = $testParams.Clone()
            $testNewParams.ImportFile = 'NonExisting'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Cannot find path*NonExisting*because it does not exist*")
            }
        } 
        It 'skip comments' {
            Mock Get-ADUser
            Mock Get-ADGroup
            Mock Get-ADComputer

            $fakeInputFile = @"
MailTo: bob@contoso.com
ADGroup: BEL ATT Leaver
# BEL ATT Leaver

OUs: OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net
OUs: OU=Users,OU=NLD,OU=EU,DC=grouphc,DC=net

#OUs: OU=Users,OU=CZE,OU=EU,DC=grouphc,DC=net
"@
            $fakeInputFile | Out-File @OutParams

            .$testScript @testParams

            $Expected = @(
                'MailTo: bob@contoso.com'
                'ADGroup: BEL ATT Leaver'
                'OUs: OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net'
                'OUs: OU=Users,OU=NLD,OU=EU,DC=grouphc,DC=net'
            )

            $File | Should -BeExactly $Expected
        } -tag test
        It 'OU missing' {
            $fakeInputFile = @"
MailTo: bob@contoso.com
ADGroup: BEL ATT Leaver                                                                      
"@
            $fakeInputFile | Out-File @OutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*no organizational units found*")
            }
        } 
        It 'OU incorrect' {
            $fakeInputFile = @"
MailTo: bob@contoso.com
ADGroup: BEL ATT Leaver
OU=Does,DC=Not,Exist=net
"@
            $fakeInputFile | Out-File @OutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Cannot validate argument on parameter 'OU'*")
            }
        } 
        It 'MailTo missing' {
            @"
ADGroup: BEL ATT Leaver
OU=Does,DC=Not,Exist=net
$ADOUExisting
"@ | Out-File @OutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*MailTo*")
            }
        } 
    }
}
    