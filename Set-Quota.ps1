<#
.SYNOPSIS
    Set quota limits on folders.

.DESCRIPTION
    This script is intended to run locally or on remote computers and has all 
    functions within the script file, not in modules. It will add/change/delete 
    quota limits, warnings, ... on folders. Quotas are applied on user level
    and are hard coded. so changes in the quota SourceTemplate are ignored.

    Only changes will be executed. In case a folder is already compliant with the request, it will simply be ignored,
    no changes will be done on that folder. Every action is logged in the 'Status' and 'Field' properties.

    Output:
        WARNING: 'E:\HOME\cnorris' Changed 'AddWarning60'
        WARNING: 'E:\HOME\cnorris' Changed 'Description'


        SamAccountName        : cnorris
        Group                 : BEL ATT Quota home 15GB
        HomeDirectory         : \\grouphc.net\BNL\HOME\Centralized\cnorris
        ComputerName          : SERVER1
        Path                  : E:\HOME\cnorris
        CurrentSize           : 5916070912
        LimitOld              : 10737418240
        LimitNew              : 10737418240
        TypeOld               : Hard
        TypeNew               : Hard
        Status                : Changed
        Field                 : AddWarning60, Description
        WarningOld            :
        WarningNew            : 60
        Description           : PowerShell managed quota based on AD membership
        SourceTemplateName    : HOME - 10 GB hard quota - changed to SOFT
        MatchesSourceTemplate : False
        PSComputerName        : SERVER1.grouphc.net
        RunSpaceId            : 6b6a25a3-f6a6-4bb9-b6e3-b75a23dc26df

    In case the user needs the SourceTemplate restored, and all modifications on the user's folder need to be removed,
    the user should be added in AD to a group ending with 'REMOVE' which is translated in property
    'Limit = 'RemoveQuota'.

    For e-mail warnings to work, the server needs to be enabled for anonymous e-mail sending.

.PARAMETER User
    Multiple objects containing all the quota details:

    $User = @(
        [PSCustomObject]@{
            Group          = 'BEL ATT Quota home 15GB'
            SamAccountName = 'cnorris'
            User           = "Chuck Norris"
            Limit          = '15GB'
            LimitBytes     = 16106127360
            HomeDirectory  = '\\grouphc.net\BNL\HOME\Centralized\cnorris'
            ComputerName   = 'SERVER1'
            ComputerPath   = 'E:\HOME\cnorris'
        },
        [PSCustomObject]@{
            Group          = 'BEL ATT Quota home 10GB'
            SamAccountName = 'lswagger'
            User           = "Bob Lee Swagger"
            Limit          = '10GB'
            LimitBytes     = 10737418240
            HomeDirectory  = '\\grouphc.net\BNL\HOME\Centralized\lswagger'
            ComputerName   = 'SERVER1'
            ComputerPath   = 'E:\HOME\lswagger'
        },
        [PSCustomObject]@{
            Group          = 'BEL ATT Quota home REMOVE'
            SamAccountName = 'jbond'
            User           = "James Bond"
            Limit          = 'RemoveQuota'
            LimitBytes     = $null
            HomeDirectory  = '\\grouphc.net\BNL\HOME\Centralized\jbond'
            ComputerName   = 'SERVER1'
            ComputerPath   = 'E:\HOME\jbond'
        }
    )

.PARAMETER DescriptionText
    The description that will be added to the quota applied on the user.

.EXAMPLE
    Set a quota limit of 15GB on path 'E:\HOME\cnorris' on remote server 
    'SERVER1'. When a threshold of 90% or 60% of the quota limit has been 
    reached, send an e-mail to the user by adding two warnings.

    $User = @(
        [PSCustomObject]@{
            Group          = 'BEL ATT Quota home 15GB'
            SamAccountName = 'cnorris'
            User           = "Chuck Norris"
            Limit          = '15GB'
            LimitBytes     = 16106127360
            HomeDirectory  = '\\grouphc.net\BNL\HOME\Centralized\cnorris'
            ComputerName   = 'SERVER1'
            ComputerPath   = 'E:\HOME\cnorris'
        }
    )

    $Warning = @(
        [PSCustomObject]@{
            Threshold        = 90
            Type       = 2 # Mails
            RunLimitInterval = 60 # Minutes
            MailTo           = '[Source Io Owner Email]'
            Subject      = '[Quota Threshold]% quota threshold exceeded'
            Body      = "User [Source Io Owner] has exceeded the [Quota Threshold]% " + `
            "quota threshold for the quota on [Quota Path] on server [Server]. The quota limit " + `
            "is [Quota Limit MB] MB, and " + ` "[Quota Used MB] MB currently is in use ([Quota " + `
            "Used Percent]% of limit)."
        },
        [PSCustomObject]@{
            Threshold        = 60
            Type       = 2 # Mails
            RunLimitInterval = 60 # Minutes
            MailTo           = '[Source Io Owner Email]'
            Subject      = '[Quota Threshold]% quota threshold exceeded'
            Body      = "User [Source Io Owner] has exceeded the [Quota Threshold]% " + `
            "quota threshold for the quota on [Quota Path] on server [Server]. The quota limit " + `
            "is [Quota Limit MB] MB, and " + ` "[Quota Used MB] MB currently is in use ([Quota " + `
            "Used Percent]% of limit)."
        }
    )

    $Session = New-PSSession -ComputerName $User.ComputerName

    $InvokeParams = @{
        FilePath     = $SetQuotaScriptFile
        ArgumentList = $User
        Session      = $Session
    }
    Invoke-Command @InvokeParams
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [Object[]]$User
)

Begin {
    Function Convert-QuotaFlagToStringHC {
        Param (
            $Number
        )

        Try {
            switch ($Number) {
                $null { $null; break }
                '0' { 'Soft'; break }
                '256' { 'Hard'; break }
                '512' { 'Disabled (Soft)'; break }
                '768' { 'Disabled (Hard)'; break }
                Default { "Unknown number '$_'" }
            }
        }
        Catch {
            throw "Failed converting quota limit type '$Number'"
        }
    }

    Function Register-ChangeHC {
        Param (
            [Parameter(Mandatory)]
            [String]$Name
        )

        Try {
            $U.Field += $Name
            $U.Status = 'Changed'
            Write-Warning "'$env:COMPUTERNAME' path '$($U.ComputerPath)' changed '$Name'"
        }
        Catch {
            throw "Failed registering change '$Name': $_"
        }
    }

    Function Add-WarningHC {
        $Quota.AddThreshold($_.Percentage)
        $Action = $Quota.CreateThresholdAction($_.Percentage, $_.Action.Type)
        $Action.MailFrom = $_.Action.MailFrom
        $Action.MailTo = $_.Action.MailTo
        $Action.MailCc = $_.Action.MailCc
        $Action.MailBcc = $_.Action.MailBcc
        $Action.MailReplyTo = $_.Action.MailReplyTo
        $Action.MailSubject = $_.Action.Subject
        $Action.MessageText = $_.Action.Body
        $Action.RunLimitInterval = $_.Action.RunLimitInterval
    }

    Function Test-ThresholdHC {
        Param (
            $Threshold
        )

        Try {
            foreach ($T in $Threshold) {
                #region Percentage
                if (-not $T.Percentage) {
                    throw 'Percentage can not be empty'
                }
                else {
                    if (@(1..100) -notcontains $T.Percentage) {
                        throw "Percentage '$($T.Percentage)' is invalid"
                    }
                }
                #endregion

                #region Type
                if (-not $T.Action.Type) {
                    throw "Action.Type for percentage '$($T.Percentage)' can not be empty"
                }
                else {
                    if (@(1..4) -notcontains $T.Action.Type) {
                        throw "Action.Type '$($T.Action.Type)' for percentage '$($T.Percentage)' is invalid"
                    }
                }
                #endregion

                #region RunLimitInterval
                if (-not $T.Action.RunLimitInterval) {
                    throw "Action.RunLimitInterval for percentage '$($T.Percentage)' can not be empty"
                }
                else {
                    if (@(0..10080) -notcontains $T.Action.RunLimitInterval) {
                        throw "Action.RunLimitInterval '$($T.Action.RunLimitInterval)' for percentage '$($T.Percentage)' is invalid"
                    }
                }
                #endregion

                #region MailFrom
                if ((-not $T.Action.MailFrom) -and ($T.Action.MailFrom -ne '')) {
                    $T.Action | Add-Member -NotePropertyName MailFrom -NotePropertyValue ''
                }
                else {
                    $T.Action.MailFrom = $T.Action.MailFrom.Trim()
                }
                #endregion

                #region MailTo
                if (-not $T.Action.MailTo) {
                    throw "Action.MailTo for percentage '$($T.Percentage)' can not be empty"
                }
                else {
                    $T.Action.MailTo = $T.Action.MailTo.Trim()
                }
                #endregion

                #region MailCc
                if ((-not $T.Action.MailCc) -and ($T.Action.MailCc -ne '')) {
                    $T.Action | Add-Member -NotePropertyName MailCc -NotePropertyValue ''
                }
                else {
                    $T.Action.MailCc = $T.Action.MailCc.Trim()
                }
                #endregion

                #region MailReplyTo
                if ((-not $T.Action.MailReplyTo) -and ($T.Action.MailReplyTo -ne '')) {
                    $T.Action | Add-Member -NotePropertyName MailReplyTo -NotePropertyValue ''
                }
                else {
                    $T.Action.MailReplyTo = $T.Action.MailReplyTo.Trim()
                }
                #endregion

                #region MailBcc
                if ((-not $T.Action.MailBcc) -and ($T.Action.MailBcc -ne '')) {
                    $T.Action | Add-Member -NotePropertyName MailBcc -NotePropertyValue ''
                }
                else {
                    $T.Action.MailBcc = $T.Action.MailBcc.Trim()
                }
                #endregion

                #region Subject
                if (-not $T.Action.Subject) {
                    throw "Action.Subject for percentage '$($T.Percentage)' can not be empty"
                }
                else {
                    $T.Action.Subject = $T.Action.Subject.Trim()
                }
                #endregion

                #region Body
                if (-not $T.Action.Body) {
                    throw "Action.Body for percentage '$($T.Percentage)' can not be empty"
                }
                else {
                    $T.Action.Body = $T.Action.Body.Trim()
                }
                #endregion
            }
        }
        Catch {
            throw "Failed threshold test: $_"
        }
    }
}

Process {
    if (-not ($User.ComputerName)) {
        throw "The property 'ComputerName' is mandatory"
    }

    $User = $User | Where-Object {
        ($_.ComputerName -eq ($env:COMPUTERNAME + '.' + $env:USERDNSDOMAIN) -or
            ($_.ComputerName -eq $env:COMPUTERNAME)) }

    Write-Verbose "Set quotas for '$(($User | Measure-Object).Count)' users on '$env:COMPUTERNAME'"

    #region Load COM objects
    Try {
        $FS = New-Object -com Fsrm.FsrmSetting
        $FQM = New-Object -ComObject Fsrm.FsrmQuotaManager
        $FQMT = New-Object -ComObject Fsrm.FsrmQuotaTemplateManager
    }
    Catch {
        throw "Quota manager role 'Fsrm.FsrmQuotaManager' not installed on '$env:COMPUTERNAME'"
    }
    #endregion

    if (($User.Threshold) -and (-not $FS.SmtpServer)) {
        throw "Please request SMTP relay for '$env:COMPUTERNAME' and configure the server for sending e-mails by running [Set-FsrmSetting -CimSession '$env:COMPUTERNAME' -AdminEmailAddress 'xxx.@heidelbergcement.com' -SmtpServer 'xxx.GROUPHC.NET' -FromEmailAddress 'xxx@heidelbergcement.com']."
    }

    $Result = Foreach ($U in $User) {
        Try {
            $HasQuota = $true
            $TemplateWarning = $null

            #region Convert to HashTable
            if ($U.GetType().Name -ne 'HashTable') {
                Write-Verbose "Convert to type 'HashTable'"
                $HashTable = @{ }
                $U.PSObject.Properties | ForEach-Object {
                    $HashTable[$_.Name] = $_.Value
                }

                $U = $HashTable
            }
            #endregion

            #region Add new properties
            $NewProps = [Ordered]@{
                Usage           = $null
                LimitOld        = $null
                LimitNew        = $U.Size
                TypeOld         = $null
                TypeNew         = if ($U.SoftLimit) { 0 } else { 256 }
                Status          = $null
                Field           = @()
                WarningOld      = @()
                WarningNew      = $U.Threshold
                Template        = $null
                MatchesTemplate = $null
            }

            $U += $NewProps
            #endregion

            #region Test mandatory properties
            'ComputerName', 'ComputerPath', 'Size', 'SoftLimit' | Where-Object { $U.Keys -notcontains $_ } | ForEach-Object {
                throw "The property '$_' is mandatory"
            }
            #endregion

            #region Test if user folder exists
            if (-not (Test-Path -LiteralPath $U.ComputerPath -PathType Container)) {
                Write-Error "Folder '$($U.ComputerPath)' not found on '$env:COMPUTERNAME'"
                Continue
            }
            #endregion

            Test-ThresholdHC -Threshold $U.WarningNew

            #region SourceTemplate
            Try {
                $ParentQuota = $FQM.GetAutoApplyQuota((Split-Path $U.ComputerPath -Parent))
            }
            Catch {
                throw 'No auto apply quota template set on parent folder'
            }

            $U.Template = $ParentQuota.SourceTemplateName
            Write-Verbose "'$($U.ComputerPath)' SourceTemplate '$($ParentQuota.SourceTemplateName)'"
            #endregion

            #region Get quota
            Try {
                Write-Verbose "'$($U.ComputerPath)' Get quota"
                $Quota = $FQM.GetQuota($U.ComputerPath)

                $U.LimitOld = $Quota.QuotaLimit
                $U.TypeOld = $Quota.QuotaFlags
                $U.WarningOld = $Quota.Thresholds
            }
            Catch {
                Write-Verbose "'$($U.ComputerPath)' No quota set"

                $HasQuota = $false

                $Quota = $FQM.CreateQuota($U.ComputerPath)

                Write-Verbose "'$($U.ComputerPath)' Apply SourceTemplate '$($ParentQuota.SourceTemplateName)'"
                Register-ChangeHC -Name 'AddQuota'
                $Quota.ApplyTemplate($ParentQuota.SourceTemplateName)
            }
            #endregion

            #region Remove quota
            if ($U.Limit -eq 'RemoveQuota') {
                $U.LimitNew = $ParentQuota.QuotaLimit
                $U.TypeNew = $ParentQuota.QuotaFlags
                $U.Description = $null
            }

            if ((($U.Limit -eq 'RemoveQuota') -and (-not $HasQuota)) -or
                (($U.Limit -eq 'RemoveQuota') -and (-not $Quota.MatchesSourceTemplate))) {
                Write-Verbose "'$($U.ComputerPath)' Apply SourceTemplate '$($ParentQuota.SourceTemplateName)'"
                # Includes Warnings, limit, type, ...
                Register-ChangeHC -Name 'ApplySourceTemplate'
                $Quota.ApplyTemplate($ParentQuota.SourceTemplateName)
            }
            #endregion

            #region QuotaLimit
            Write-Verbose "'$($U.ComputerPath)' Limit old '$($U.LimitOld)' ($([MATH]::Round(($U.LimitOld/1GB),2)) GB)"
            Write-Verbose "'$($U.ComputerPath)' Limit new '$($U.LimitNew)' ($([MATH]::Round(($U.LimitNew/1GB),2)) GB)"

            if ($U.LimitOld -ne $U.LimitNew) {
                $Quota.QuotaLimit = [Decimal]$U.LimitNew
                Register-ChangeHC 'Limit'
            }
            #endregion

            #region QuotaFlags
            Write-Verbose "'$($U.ComputerPath)' Type old '$($U.TypeOld)' ($(Convert-QuotaFlagToStringHC $U.TypeOld))"
            Write-Verbose "'$($U.ComputerPath)' Type new '$($U.TypeNew)' ($(Convert-QuotaFlagToStringHC $U.TypeNew))"

            if ($U.TypeOld -ne $U.TypeNew) {
                $Quota.QuotaFlags = $U.TypeNew
                Register-ChangeHC 'Type'
            }
            #endregion

            #region Threshold
            if ($U.Limit -ne 'RemoveQuota') {
                if ($U.WarningNew) {
                    $U.WarningNew | Where-Object { $U.WarningOld -notcontains $_.Percentage } | ForEach-Object {
                        Register-ChangeHC ('AddWarning' + $_.Percentage)
                        Add-WarningHC
                    }

                    $U.WarningOld | Where-Object { $U.WarningNew.Percentage -notcontains $_ } | ForEach-Object {
                        Register-ChangeHC ('RemoveWarning' + $_)
                        $Quota.DeleteThreshold($_)
                    }

                    $U.WarningNew | Where-Object { $U.WarningOld -contains $_.Percentage } | ForEach-Object {
                        $OldThreshold = $Quota.EnumThresholdActions($_.Percentage)

                        if ((($OldThreshold | Select-Object -ExpandProperty ActionType) -ne $_.Action.Type) -or
                            (($OldThreshold | Select-Object -ExpandProperty MailFrom) -ne $_.Action.MailFrom) -or
                            (($OldThreshold | Select-Object -ExpandProperty MailReplyTo) -ne $_.Action.MailReplyTo) -or
                            (($OldThreshold | Select-Object -ExpandProperty MailTo) -ne $_.Action.MailTo) -or
                            (($OldThreshold | Select-Object -ExpandProperty MailBcc) -ne $_.Action.MailBcc) -or
                            (($OldThreshold | Select-Object -ExpandProperty MailCc) -ne $_.Action.MailCc) -or
                            (($OldThreshold | Select-Object -ExpandProperty RunLimitInterval) -ne $_.Action.RunLimitInterval) -or
                            (($OldThreshold | Select-Object -ExpandProperty MailSubject) -ne $_.Action.Subject) -or
                            (($OldThreshold | Select-Object -ExpandProperty MessageText) -ne $_.Action.Body)) {
                            Register-ChangeHC ('ModifyWarning' + $_.Percentage)
                            $Quota.DeleteThreshold($_.Percentage)
                            Add-WarningHC
                        }
                    }
                }
                elseif ($U.WarningOld) {
                    $U.WarningOld | ForEach-Object { $Quota.DeleteThreshold($_) }
                    Register-ChangeHC RemoveWarningAll
                }
            }
            #endregion

            #region Description

            #$U.Description = $U.Description
            if (-not $U.Description) {
                $U.Description = ''
            }
            else {
                $U.Description = $U.Description.Trim()
            }

            if ($Quota.Description -ne $U.Description) {
                $Quota.Description = [String]$U.Description
                Register-ChangeHC Description
            }
            #endregion

            #region Commit changes
            if ($U.Status) {
                Write-Verbose "'$($U.ComputerPath)' Commit changes"
                $Quota.Commit()
            }
            else {
                Write-Verbose "'$($U.ComputerPath)' No changes"
                $U.Status = 'Ok'
            }
            #endregion

            #region Format fields
            $U.Field = $U.Field -join ', '
            $U.WarningOld = $U.WarningOld -join ', '
            $U.TypeOld = Convert-QuotaFlagToStringHC $U.TypeOld
            #endregion

            [PSCustomObject]$U
        }
        Catch {
            throw "Failed setting quota limit '$($U.Limit)' on path '$($U.ComputerPath)' on '$($U.ComputerName)': $_"
        }
    }

    #region Verify new quotas
    Write-Verbose "Verify new quotas for '$(($Result | Measure-Object).Count)' users"

    if ($Result.Status -contains 'Changed') {
        Start-Sleep -Seconds 3
    }

    foreach ($R in $Result) {
        Try {
            Write-Verbose "'$($R.ComputerPath)' Get quota"
            $NewQuota = $FQM.GetQuota($R.ComputerPath)

            $R.Usage = $NewQuota.QuotaUsed
            Write-Verbose "'$($R.ComputerPath)' Current size '$($R.Usage)' ($([MATH]::Round(($R.Usage/1GB),2)) GB)"

            $R.WarningNew = $NewQuota.Thresholds -join ', '
            $R.TypeNew = Convert-QuotaFlagToStringHC $NewQuota.QuotaFlags
            $R.LimitNew = $NewQuota.QuotaLimit
            $R.Description = $NewQuota.Description
            $R.MatchesTemplate = $NewQuota.MatchesSourceTemplate

            $R
        }
        Catch {
            Write-Error "Failed retrieving quota details for'$($R.ComputerPath)' on '$env:COMPUTERNAME': $_"
        }
    }
    #endregion

    Write-Verbose "Quotas set on '$(($Result | Measure-Object).Count)' folders"
}