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
        ScriptAdmin = 'admin@contoso.com'
    }

    $MailAdminParams = {
        ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
        ($Subject -eq 'FAILURE')
    }

    $ADOUExisting = Get-ADOrganizationalUnit -Filter * |
    Select-Object -First 1 -ExpandProperty DistinguishedName

    Mock Invoke-Sqlcmd
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
        It 'OU missing' {
            $fakeInputFile = @{
                MailTo = 'bob@contoso.com'
                AD     = @{
                    OU        = @()
                    GroupName = 'BEL ATT Leaver'
                }
            }
            $fakeInputFile | ConvertTo-Json | Out-File @OutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*No 'AD.OU' found.")
            }
        }
        It 'OU incorrect' {
            $fakeInputFile = @{
                MailTo = 'bob@contoso.com'
                AD     = @{
                    OU        = @('OU=Does,DC=Not,Exist=net')
                    GroupName = 'BEL ATT Leaver'
                }
            }
            $fakeInputFile | ConvertTo-Json | Out-File @OutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Cannot validate argument on parameter 'OU'*")
            }
        }
        It 'MailTo missing' {
            $fakeInputFile = @{
                # MailTo = 'bob@contoso.com'
                AD = @{
                    OU        = @($ADOUExisting)
                    GroupName = 'BEL ATT Leaver'
                }
            }
            $fakeInputFile | ConvertTo-Json | Out-File @OutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*MailTo*")
            }
        }
    }
}
