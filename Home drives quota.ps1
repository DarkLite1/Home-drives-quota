#Requires -Version 5.1

<#
    .SYNOPSIS
        Apply hard quota limits on user's home folders, based on active 
        directory group membership. A summary e-mail is send with an Excel 
        sheet in attachment containing all the details.

    .DESCRIPTION
        The 'Home drive quota' script applies hard quota limits on user's home 
        folders, based on the active directory group membership. A summary 
        e-mail is send with an Excel sheet in attachment containing all the 
        details.

        Home folders defined on the file server for users that are not member 
        of an AD Quota management group are ignored. They are not changed or 
        reported in the attachment of the sent e-mail.

        The script is started by a Scheduled Task that is planned to run every 
        night. This means that changes are only visible the next day..

        Requirements
        -	On the file server:
            o	we need to be local administrator
            o	Set-ExecutionPolicy RemoteSigned
            o	Anonymous mail sending capability
                Check SMTP relay or request it via Group IT
            o	Set SMTP server for quota management
                right click the role 
                'File Server Resource Manager' > 'Configure options'
                Set 'SMTP server name' and 'Default from e-mail address'
            o	An AutoApply template on the parent folder for the home folders
                (A soft limit is advised, this way the hard limits are managed 
                by the script)

        Information
        -	When users are member of multiple quota management groups at the    
            same time, the script does
            nothing and report this as incorrect.

        -	When a user was first member of a quota management group and later 
            on removed from that group,
            he will still have his old hard quota limit applied. To fix this, 
            add the user to the group defined in [-ADGroupRemoveName].

        -	Users that have quota limits applied on the server, but are not 
            member of an AD quota management group, will simply be ignored. 
            This script only manages quotas on home folders for users that
            are member of an AD quota limit group.

        -	Users that are not 'Enabled' or that don't have the attribute 
            'HomeDirectory' set are ignored.

        -	Supported OS Windows Server 2008 or newer

    .PARAMETER ADGroupName
        The prefix used to find the quota management groups in active directory 
        with their size limit. The string found after the prefix needs to be a 
        valid size (ex. 5GB, 250MB, ..).

        Ex. [-ADGroupName 'BEL H Quota'] includes users in groups:
            'BEL H Quota 5GB'   > SizeLimit 5GB
            'BEL H Quota 10GB'  > SizeLimit 10GB
            'BEL H Quota 500MB  > SizeLimit 500MB
            ...

    .PARAMETER ADGroupRemoveName
        Members of this group will have the SourceTemplate of the parent folder 
        re-applied on their home drive.

        This is needed to 'un-manage' users that were previously 'managed' by 
        the script. Users that previously were member of one of the AD quota 
        management groups, have a hard quota size limit set. If the user is 
        removed from one of the latter groups, he will still have the 
        previously applied hard quota. To undo this, the SourceTemplate will 
        have to be re-applied, which usually has a SoftLimit.

    .PARAMETER MailTo
        E-mail address where the summary report, with the Excel file containing 
        all the details, will be send to. It is strongly advised to use a mail 
        enabled distribution list, so membership can be managed within AD.

    .PARAMETER ThresholdFile
        This is a JSON file containing information about the e-mail(s) that 
        users will receive when they surpass/breach the threshold percentage on 
        their home drive. Multiple e-mails with different percentages are 
        supported.

        If no JSON file is provided, the quota management server will not send out e-mails to users when they surpass/breach
        the quota size limit. In case a user will try to add more data to his home drive Windows will simply report that
        the disk is full and block the action when a hard limit is set.

        Ex.
        [
            {
	        "Percentage":  80,
	        "Color":  "Pink",
	        "Action":  {
		        "MailFrom":  "No-Reply@contoso.com",
	            "MailTo":  "[Source Io Owner Email]",
		        "MailCc":  "",
		        "MailBcc":  "",
		        "MailReplyTo":  "",
		        "Subject":  "Your personal home drive has reached [Quota Threshold]% of its maximum allowed volume",
		        "Body":  "Dear user,\r\n\r\nYour personal home drive has reached [Quota Threshold]% of its maximum allowed volume.\r\n\r\nThe quota limit is [Quota Limit MB] MB, and [Quota Used MB] MB currently is in use ([Quota Used Percent]% of limit). Please take into account that once your quota has been reached, you will no longer be able to edit or save files on your home drive.\r\n\r\nIn case you need assistance in handling this issue, we kindly ask you to call us on xxx or visit http://selfservice \r\n\r\nKind regards,\r\nIT Service Desk",
	        "RunLimitInterval":  60,
	        "Type":  2
	        }
            },
            {
	        "Percentage":  90,
	        "Color":  "Orange",
        	    "Action":  {
		            "MailFrom":  "No-Reply@contoso.com",
             	    "MailTo":  "[Source Io Owner Email]",
		            "MailCc":  "",
		            "MailBcc":  "",
		            "MailReplyTo":  "",
		            "Subject":  "Your personal home drive has reached [Quota Threshold]% of its maximum allowed volume",
                                   "Body":  "Dear user,\r\n\r\nYour personal home drive has reached [Quota Threshold]% of its maximum allowed volume.\r\n\r\nThe quota limit is [Quota Limit MB] MB, and [Quota Used MB] MB currently is in use ([Quota Used Percent]% of limit). Please take into account that once your quota has been reached, you will no longer be able to edit or save files on your home drive.\r\n\r\nIn case you need assistance in handling this issue, we kindly ask you to call us on xxx or visit http://selfservice \r\n\r\nKind regards,\r\nIT Service Desk",
	        "RunLimitInterval":  60,
	        "Type":  2
	        }
            },
            {
	        "Percentage":  100,
	        "Color":    "Red",
	        "Action":  {
		        "MailFrom":  "No-Reply@contoso.com",
		        "MailTo":  "[Source Io Owner Email]",
		        "MailCc":  "",
		        "MailBcc":  "",
		        "MailReplyTo":  "",
		        "Subject":  "Your personal home drive has reached or exceeded its maximum allowed volume",
		        "Body":  "Dear user,\r\n\r\nYour personal home drive has now reached [Quota Threshold]% of its maximum allowed volume.\r\nPlease take into account that you will no longer be able to edit or save files on your home drive. You can either take the necessary measures to free up space on your home drive, or submit a formal request for a higher volume limit.\r\n\r\nIn case you need assistance in handling this issue, we kindly ask you to call us on xxx or visit http://selfservice \r\n\r\nKind regards,\r\nIT Service Desk",
		        "RunLimitInterval":  60,
		        "Type":  2
	            }
            }
        ]


    .EXAMPLE
        Normal scenario 1
        Mike is added to AD group 'BEL H Quota 5GB'. The script starts and sets 
        a hard limit of 5GB on the 'HomeDirectory' of Mike

    .EXAMPLE
        Normal scenario 2
        Jim is removed from AD group 'BEL H Quota 5GB' and added to AD group 
        'BEL H Quota 10GB'. The script starts and changes the hard limit from 
        5GB to 10GB on the 'HomeDirectory' of Jim

    .EXAMPLE
        Special scenario
        1. Bob is added to AD group 'BEL H Quota 5GB'
           The script starts and sets a hard limit of 5GB on the 
           'HomeDirectory' of Bob
        2. Bob is removed from AD group 'BEL H Quota 5GB' and is member of no
            other group
           > Script starts and doesn't do anything on Bob's 'HomeDirectory'
           > Bob still has a hard limit of 5GB on his 'HomeDirectory'
        3. Bob is added to AD group 'BEL H Quota REMOVE'
           > Bob's 'HomeDirectory' will have the SourceTemplate applied again
           > If the SourceTemplate is a 'SoftLimit', then the hard limit will be removed

    .LINK
        https://www.simple-talk.com/sysadmin/exchange/implementing-windows-server-2008-file-system-quotas/
 #>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName = 'Home drives quota (BNL)',
    [Parameter(Mandatory)]
    [String]$ADGroupName = 'BEL ATT Quota home',
    [Parameter(Mandatory)]
    [String]$ADGroupRemoveName = 'BEL ATT Quota home REMOVE',
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$ThresholdFile,
    [String]$SetQuotaScriptFile = '.\Set-Quota.ps1',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Home drives\Home drives quota\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        $Error.Clear()
        $LogFile = $null
        Add-Type -Assembly System.Drawing
        # Import-Module PSWorkflow

        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        Set-Culture 'en-US'

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

        #region ThresholdFile
        $Threshold = $null
        $Highlight = [Ordered]@{ }

        if ($ThresholdFile) {
            if (-not (Test-Path $ThresholdFile -PathType Leaf)) {
                throw "Threshold file '$ThresholdFile' not found"
            }

            Try {
                [Array]$Threshold = Get-Content $ThresholdFile -Raw | ConvertFrom-Json

                if (-not $Threshold) {
                    throw 'File is empty'
                }

                foreach ($T in $Threshold) {
                    $Threshold | ForEach-Object {
                        'Color', 'Percentage', 'Action' | ForEach-Object {
                            if (-not $T.$_) {
                                throw "Property '$_' not found"
                            }
                        }
                        foreach ($P in @('MailTo', 'Subject', 'Type', 'Body', 'RunLimitInterval')) {
                            if (-not $T.Action.$P) {
                                throw "Property '$P' not found for Percentage '$($T.Percentage)'"
                            }
                        }
                    }
                }
            }
            Catch {
                throw "Threshold file '$ThresholdFile' does not contain valid (JSON) data: $_"
            }

            $Threshold | Sort-Object 'Percentage' -Descending | ForEach-Object {
                if (-not ($_.Percentage -is [Int])) {
                    Throw "'Percentage' value '$($_.Percentage)' is not valid because it's not numerical"
                }

                Try {
                    $ColorValue = $_.Color
                    $null = [System.Drawing.Color]$_.Color
                }
                Catch {
                    Throw "'Color' value '$ColorValue' is not valid because it's not a proper color"
                }

                Try {
                    $Highlight.Add($_.Percentage, [System.Drawing.Color]$_.Color)
                }
                Catch {
                    Throw "Duplicate threshold values are not possible"
                }
            }

            Write-EventLog @EventVerboseParams -Message "Threshold template '$ThresholdFile' loaded"
        }
        #endregion

        #region SetQuotaScriptFile
        Try {
            if (-not ($SetQuotaScriptFileItem = Get-Item -Path $SetQuotaScriptFile -EA Ignore)) {
                $SetQuotaScriptFileItem = Get-Item -Path (Join-Path $PSScriptRoot $SetQuotaScriptFile) -EA Stop
                if ($SetQuotaScriptFileItem.PSIsContainer) {
                    throw 'Only files are allowed'
                }
            }
        }
        Catch {
            throw "Quota script file '$SetQuotaScriptFile' not found"
        }
        #endregion

        #region Get groups
        if (-not ($Groups = Get-ADGroup -Filter "Name -like '$ADGroupName*' -or Name -like '$ADGroupRemoveName*'")) {
            throw "Couldn't find active directory groups starting with '$ADGroupName'."
        }
        Write-EventLog @EventVerboseParams -Message "Retrieved '$($Groups.Count)' groups with name like '$ADGroupName' or '$ADGroupRemoveName':`n`n$($Groups.SamAccountName -join "`n")"

        if ($Groups.SamAccountName -notcontains $ADGroupRemoveName) {
            throw "Couldn't find the active directory group '$ADGroupRemoveName' for removing quotas."
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message ($env:USERNAME + ' - ' + "FAILURE:`n`n- " + $_)
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Calculate group limits
        $Groups = $Groups | ForEach-Object {
            $Prop = [Ordered]@{
                SamAccountName = $_.SamAccountName
                MemberCount    = $null
            }

            if ($_.SamAccountName -eq $ADGroupRemoveName) {
                $Prop.Limit = 'RemoveQuota'
                $Prop.Size = 0
            }
            else {
                Try {
                    $Prop.Limit = $_.SamAccountName.Split(' ')[-1]
                    $Prop.Size = [ScriptBlock]::Create($Prop.Limit).InvokeReturnAsIs()
                }
                Catch {
                    throw "Failed converting the quota limit string (Ex. 20GB, 50GB, ..) at the end of AD GroupName '$($Prop.SamAccountName)' to a KB size : $_"
                }
            }

            [PSCustomObject]$Prop
        }
        #endregion

        #region Get group members
        $Users = foreach ($G in $Groups) {
            # Avoid pipeline for Pester tests
            $GroupMembers = @(Get-ADGroupMember $G.SamAccountName -Recursive)

            $UserCount = 0

            foreach ($GM in $GroupMembers) {
                if ($UserWithHomeDir = Get-ADUser $GM.SamAccountName -Properties HomeDirectory, HomeDrive |
                    Where-Object { $_.Enabled -and $_.HomeDirectory -and $_.HomeDrive }) {

                    $UserCount += 1

                    $UserWithHomeDir | Select-Object -ExcludeProperty Description -Property *,
                    @{N = 'GroupName'; E = { $G.SamAccountName } },
                    @{N = 'Limit'; E = { $G.Limit } },
                    @{N = 'Size'; E = { $G.Size } },
                    @{N = 'Description'; E = { "$($G.SamAccountName) #PowerShellManaged" } },
                    @{N = 'SoftLimit'; E = { $false } },
                    @{N = 'Threshold'; E = { $Threshold } }
                }
            }

            $G.MemberCount = $UserCount
            Write-EventLog @EventVerboseParams -Message "Group '$($G.SamAccountName)':`n`n- Limit $($G.Limit)`n- Size $($G.Size)`n- Members: $UserCount"
        }
        #endregion

        #region Duplicate group membership
        $Users = $Users | Group-Object SamAccountName | ForEach-Object {
            if ($_.Count -gt 1) {
                Write-Error "User '$($_.Name)' is member of multiple groups '$($_.Group.GroupName -join ', ')'"
            }
            else {
                $_.Group
            }
        }
        #endregion

        if ($Users) {
            Write-EventLog @EventVerboseParams -Message "Get DFS ComputerName and local path for $($Users.Count) users"

            if ($DFSDetails = Get-DFSDetailsHC -Path $Users.HomeDirectory) {
                Write-EventLog @EventVerboseParams -Message "DFS details retrieved"

                foreach ($U in $Users) {
                    if ($DFS = $DFSDetails | Where-Object { $U.HomeDirectory -eq $_.Path }) {
                        $Extra = @{
                            ComputerName = $DFS.ComputerName
                            ComputerPath = $DFS.ComputerPath
                        }
                        $U | Add-Member -NotePropertyMembers $Extra
                    }
                }

                $SessionParams = @{
                    ComputerName = $Users | Group-Object ComputerName | Where-Object Name | Select-Object -ExpandProperty Name
                }
                $Sessions = New-PSSession @SessionParams

                #region Set quotas
                Write-EventLog @EventVerboseParams -Message "Launch the script to set the quotas on the remote machines."

                $InvokeParams = @{
                    FilePath     = $SetQuotaScriptFileItem
                    ArgumentList = , $Users
                    Session      = $Sessions
                }
                $QuotaResults = @(Invoke-Command @InvokeParams)

                Write-EventLog @EventVerboseParams -Message "Script finished with '$($QuotaResults.Count)' results."
                #endregion

                #region Get free disk space
                Write-EventLog @EventVerboseParams -Message "Launch the script to get the free disk space on the remote machines."

                $DriveSizeResults = Invoke-Command -Session $Sessions -ScriptBlock {
                    $args | Where-Object {
                        ($_.ComputerName -eq ($env:COMPUTERNAME + '.' + $env:USERDNSDOMAIN) -or ($_.ComputerName -eq $env:COMPUTERNAME)) } |
                    Group-Object -Property { Split-Path -Path $_.ComputerPath -Qualifier } | ForEach-Object {
                        Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$($_.Name)'"
                    }
                } -ArgumentList $DFSDetails

                Write-EventLog @EventVerboseParams -Message "Script finished with '$($DriveSizeResults.Count)' results."
                #endregion
            }
            else {
                Write-EventLog @EventWarnParams -Message "No DFS details retrieved"
            }
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message ($env:USERNAME + ' - ' + "FAILURE:`n`n- " + $_)
        $Sessions | Remove-PSSession -EA Ignore
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        #region Get the numbers for email reporting
        $TotalUsers = ($Groups | Measure-Object -Property 'MemberCount' -Sum).Sum

        $StatusOk, $StatusNotOk = $QuotaResults.where( { $_.Status -eq 'Ok' }, 'Split')
        #endregion

        $MailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$TotalUsers group members, $($StatusNotOk.Count) changes"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        $LogFile += '.xlsx'
        $LogFile | Remove-Item -EA Ignore

        if ($QuotaResults) {
            #region Export users with their new limits to Excel
            Write-EventLog @EventVerboseParams -Message "Export quota execution results to Excel."
            $ExportParams = @{
                Path          = $LogFile
                WorkSheetName = 'Quotas'
                TableName     = 'Quotas'
                AutoSize      = $true
                FreezeTopRow  = $true
                AutoNameRange = $true
                PassThru      = $true
            }

            $ExcelWorkbook = $QuotaResults | Select-Object -Property GroupName,
            @{Name = 'ComputerName'; Expression = { $_.ComputerName -ireplace ('.' + $env:USERDNSDOMAIN) } },
            ComputerPath, Status, Field,
            @{Name = 'Usage'; Expression = { [Math]::Round(((100 / $_.LimitNew) * $_.Usage), 2) } },
            @{Name = 'UsageSize'; Expression = { [MATH]::Round($_.Usage / 1GB, 2) } },
            @{Name = 'LimitNew'; Expression = { [MATH]::Round($_.LimitNew / 1GB, 2) } }, TypeNew,
            @{Name = 'LimitOld'; Expression = { [MATH]::Round($_.LimitOld / 1GB, 2) } }, TypeOld,
            WarningOld, WarningNew, Template, MatchesTemplate, Description,
            LimitOld, LimitNew, Usage, SamAccountName, HomeDirectory,
            @{Name = 'LimitOld_'; Expression = { $_.LimitOld } },
            @{Name = 'LimitNew_'; Expression = { $_.LimitNew } },
            @{Name = 'Usage_'; Expression = { $_.Usage } } -ExcludeProperty ComputerName,
            LimitOld, LimitNew, Usage | Sort-Object Usage -Descending |
            Export-Excel @ExportParams -CellStyleSB {
                Param (
                    $WorkSheet,
                    $TotalRows,
                    $LastColumn
                )

                @($WorkSheet.Names['Usage'].Style).ForEach( {
                        $_.NumberFormat.Format = '? \%'
                    })

                @($WorkSheet.Names['UsageSize', 'LimitOld', 'LimitNew'].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \G\B'
                    })

                @($WorkSheet.Names['LimitOld_', 'LimitNew_', 'Usage_'].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \B'
                    })

                $WorkSheet.Cells.Style.HorizontalAlignment = 'Center'
            }

            $MailParams.Attachments = $LogFile

            #region Format percentage and set row color
            if ($ThresholdFile) {
                $WorkSheet = $ExcelWorkbook.Workbook.Worksheets[$ExportParams.WorkSheetName]
                $LastRow = $WorkSheet.Dimension.Rows

                $ConditionParams = @{
                    WorkSheet = $WorkSheet
                    Range     = "F2:F$LastRow"
                }

                $FirstTimeThrough = $true
                foreach ($H in $Highlight.GetEnumerator()) {
                    if ($FirstTimeThrough) {
                        $FirstTimeThrough = $False
                        Add-ConditionalFormatting @ConditionParams -BackgroundColor $h.Value.Name -RuleType GreaterThan -ConditionValue $H.Name
                    }
                    else {
                        Add-ConditionalFormatting @ConditionParams -BackgroundColor $h.Value.Name -RuleType Between -ConditionValue $H.Name -ConditionValue2 $PreviousPct
                    }

                    $PreviousPct = $H.Name
                }
            }
            #endregion

            $ExcelWorkbook.Save()
            $ExcelWorkbook.Dispose()
            Write-EventLog @EventVerboseParams -Message "Export quota execution results to Excel is finished."
            #endregion
        }

        if ($DriveSizeResults) {
            #region Export the computer drives and their free space to Excel
            Write-EventLog @EventVerboseParams -Message "Export the drive size results to Excel."

            $ExportParams = @{
                Path          = $LogFile
                AutoNameRange = $true
                WorkSheetName = 'DriveSize'
                TableName     = 'DriveSize'
                AutoSize      = $true
                FreezeTopRow  = $true
            }

            $DriveSizeResults | Sort-Object FreeSpace | Select-Object -Property @{Name = 'ComputerName'; Expression = { $_.PSComputerName -ireplace ('.' + $env:USERDNSDOMAIN) } },
            @{N = 'Drive'; E = { $_.DeviceID } },
            @{Name = 'FreeSpace'; Expression = { [MATH]::Round($_.FreeSpace / 1GB, 2) } },
            @{Name = 'Size'; Expression = { [MATH]::Round($_.Size / 1GB, 2) } },
            @{Name = 'FreeSpace_'; Expression = { $_.FreeSpace } },
            @{Name = 'Size_'; Expression = { $_.Size } } |
            Export-Excel @ExportParams -CellStyleSB {
                Param (
                    $WorkSheet,
                    $TotalRows,
                    $LastColumn
                )

                @($WorkSheet.Names['FreeSpace', 'Size'].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \G\B'
                    })

                @($WorkSheet.Names['FreeSpace_', 'Size_'].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \K\B'
                    })

                $WorkSheet.Cells.Style.HorizontalAlignment = 'Center'
            }

            Write-EventLog @EventVerboseParams -Message "Export the drive size results to Excel is finished."
            #endregion

            $MailParams.Attachments = $LogFile
        }

        if ($Error) {
            #region Export errors to Excel
            Write-EventLog @EventVerboseParams -Message "Export $($Error.Count) errors to Excel."

            $ExportErrorParams = @{
                Path          = $LogFile
                WorkSheetName = 'Errors'
                TableName     = 'Errors'
                AutoSize      = $true
                FreezeTopRow  = $true
            }
            $Error.Exception.Message | Select-Object @{N = 'Error message'; E = { $_ } } |
            Export-Excel @ExportErrorParams

            $Error.ForEach( {
                    Write-EventLog @EventErrorParams -Message $_.Exception.Message
                })
            #endregion

            $MailParams.Subject = "FAILURE $TotalUsers quota group members"
            $MailParams.Priority = 'High'
            $MailParams.Attachments = $LogFile
        }

        #region Email formatting
        $MailParams.Message = if ($Error) {
            "<p><b>$($Error.Count) problems</b> were detected while applying quota limits, check the worksheet 'Errors'.</p>"
        }
        elseif ($QuotaResults.Count -eq $StatusOk.Count) {
            'All quota limits are correct, no changes were required.'
        }
        else {
            'Some quota limits were incorrect and have now been corrected.'
        }

        $MailParams.Message += "
        <ul>
            <li>$($StatusOk.Count) home directories had the correct quota limit applied.</li>
            $(if ($StatusNotOk.Count -ne 0) {
                "<li><b>$($StatusNotOk.Count) home directories</b> had their quota limit corrected.</li>"
            })
        </ul>
        <p>The following users are excluded: disabled users, users without a home directory and users with a home directory that have no drive letter.</p>
        <h3>Active directory groups:</h3>
        $($Groups | Sort-Object Size |
            Select-Object SamAccountName, Limit, @{N = 'Members'; E = { $_.MemberCount } } |
            ConvertTo-Html -Fragment -As Table)
        <p><i>* Check the attachment for details</i></p>"
        #endregion

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject "FAILURE" -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message ($env:USERNAME + ' - ' + "FAILURE:`n`n- " + $_); Exit 1
    }
    Finally {
        $Sessions | Remove-PSSession -EA Ignore
        Write-EventLog @EventEndParams
    }
}