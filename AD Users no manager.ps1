﻿#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Toolbox.Remoting

<#
    .SYNOPSIS
        Report about all the users that don't have a manager assigned to them
        in the AD.

    .DESCRIPTION
        Report about all the users that don't have a manager assigned to them
        in the AD. The report is an e-mail containing the total number of users
        per country that don't have a manager. In attachment is an Excel sheet
        containing these specific users.

        When an AD group name is provided with the parameter 'ADGroup' in the
        'ImportFile' we check if users are member of one of these groups and
        add an extra boolean column to the Excel file for each group name.

        All results are stored in an SQL database for use in the Excel sheet by
        the PivotTable. This to generate a line graph with an overview of
        progress throughout time.

    .PARAMETER ImportFile
        Contains all needed parameters:
        - MailTo
        - OU Distinguished names
        - AD Group

    .PARAMETER LogFolder
        Location for the log files.
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName = 'AD Users no manager',
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$SQLServerInstance = 'GRPSDFRAN0049',
    [String]$SQLDatabase = 'PowerShell',
    [String]$SQLTableReportUsersNoManager = 'ReportUsersNoManager',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Users no manager\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        $Now = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
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

        $SQLParams = @{
            ServerInstance         = $SQLServerInstance
            Database               = $SQLDatabase
            TrustServerCertificate = $true
            QueryTimeout           = '1000'
            ConnectionTimeout      = '20'
            ErrorAction            = 'Stop'
        }

        #region Import input file
        $ImportFileName = (Get-Item $ImportFile -EA Stop).BaseName

        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($adOUs = $File.AD.OU)) {
            throw "Input file '$ImportFile': No 'AD.OU' found."
        }

        $adGroupNames = $File.AD.GroupName

        $adGroups = foreach ($groupName in $adGroupNames) {
            [PSCustomObject]@{
                Name    = $groupName
                Members = Get-ADGroupMember $groupName -Recursive
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $UsersNoManager = Get-ADUserNoManagerHC -OU $adOUs -EA Stop

        $HtmlOus = $adOUs | ConvertTo-OuNameHC -OU | Sort-Object |
        ConvertTo-HtmlListHC -Header 'Organizational units:'

        Switch (($UsersNoManager | Measure-Object).Count) {
            '0' {
                $Intro = "<p>All users have a manager assigned.</p>"
                $Subject = "All users have a manager assigned"
                $Priority = 'Normal'
            }
            '1' {
                $Intro = "<p>Only <b>1 user</b> has <b>no manager</b>:</p>"
                $Subject = "1 user has no manager"
                $Priority = 'High'
            }
            Default {
                $Intro = "<p><b>$_ users</b> have <b>no manager</b>:</p>"
                $Subject = "$_ users have no manager"
                $Priority = 'High'
            }
        }

        if ($UsersNoManager) {
            if ($adGroups) {
                $UsersNoManager | ForEach-Object {
                    $Sam = $_.'Logon name'

                    $Properties = [Ordered]@{}
                    $adGroups | ForEach-Object {
                        $Properties.($_.Name) = $_.Members.SamAccountName -contains $Sam
                    }

                    $_ | Add-Member -NotePropertyMembers $Properties
                }
            }

            $Results = $UsersNoManager | Group-Object Country | Select-Object Count, Name

            foreach ($R in $Results) {
                Invoke-Sqlcmd @SQLParams -Query "
                    INSERT INTO $SQLTableReportUsersNoManager
                    (RunDate, ImportFile, Country, Total)
                    VALUES ('$("{0:yyyy-MM-dd HH:mm:ss}" -f $Now)',
                        '$ImportFileName', '$($R.Name)', '$($R.Count)')"
            }

            $Results = Invoke-Sqlcmd @SQLParams -Query "
                SELECT * FROM $SQLTableReportUsersNoManager WHERE ImportFile = '$ImportFileName'"

            $ExcelParams = @{
                Path         = $LogFile + '.xlsx'
                AutoSize     = $true
                FreezeTopRow = $true
            }

            $UsersNoManager | Export-Excel @ExcelParams -WorksheetName Users -TableName User -NoNumberConversion 'Employee ID',
            'OfficePhone', 'HomePhone', 'MobilePhone', 'ipPhone', 'Fax', 'Pager'

            $Results | Export-Excel -Path $ExcelParams.Path -WorksheetName HistoryLine -PivotRows RunDate -PivotColumns Country -PivotData @{Total = 'Sum' } -ChartType Line -IncludePivotTable -IncludePivotChart -HideSheet HistoryLine

            $Results | Export-Excel -Path $ExcelParams.Path -WorksheetName HistoryBar -PivotRows Country -PivotColumns RunDate -PivotData @{Total = 'Sum' } -ChartType ColumnClustered -IncludePivotChart -IncludePivotTable -HideSheet HistoryBar

            $Table = $UsersNoManager | Group-Object Country |
            Select-Object @{Name = "Country"; Expression = { $_."Name" } }, @{Name = "Total"; Expression = { $_."Count" } } |
            Sort-Object Country | ConvertTo-Html -As Table -Fragment

            $Message = "$Intro
                        $Table
                        <p><i>* Check the attachment for details and an historical overview</i></p>"
        }
        else {
            $Message = $Intro
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
    Try {
        $EmailParams = @{
            To          = $MailTo
            Bcc         = $ScriptAdmin
            Subject     = $Subject
            Message     = $Message, $HtmlOus
            Attachments = $ExcelParams.Path
            Priority    = $Priority
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = $LogFile + ' - Mail.html'
        }
        Remove-EmptyParamsHC $EmailParams
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @EmailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}