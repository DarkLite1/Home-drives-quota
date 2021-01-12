#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $SourceTemplateName = 'TEST HOME Quota (10GB Soft)'

    Function Set-EqualToTemplateHC {
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [String]$Template
        )
    
        Try {
            Remove-FsrmQuota -Path $Path -Confirm:$false
            New-FsrmQuota -Path $Path -Template $Template
        }
        Catch {
            throw "Cound not set template '$Template' on '$Path': $_"
        }
    }
    Function Test-QuotaDetailsHC {
        Param (
            $Actual,
            $Expected,
            [ValidateSet('ExpectedBefore', 'ExpectedAfter')]
            [String]$When
        )
    
        $Properties = @('Size' , 'Description', 'SoftLimit', 'Template', 'Disabled', 'MatchesTemplate')
        foreach ($P in $Properties) {
            $Actual.$P | Should -BeExactly $Expected.$P -Because "the '$P' is not the same $When"
        }
    
        $ThresholdProperties = @('Percentage')
        foreach ($P in $ThresholdProperties) {
            $Actual.Threshold.$P | Should -BeExactly $Expected.Threshold.$P -Because "the 'Threshold.$P' is not the same $When"
        }
    
        foreach ($E in $Expected.Threshold) {
            $A = $Actual.Threshold | Where-Object { $E.Percentage -eq $_.Percentage }
    
            $ThresholdActionProperties = @('Body', 'MailTo', 'Subject', 'Type', 'RunLimitInterval')
    
            foreach ($P in $ThresholdActionProperties) {
                $A.Action.$P | Should -BeExactly $E.Action.$P -Because "the 'Threshold.Action.$P' for percentage '$($E.Percentage)' is not the same $When"
            }
        }
    }
    Function Test-ScriptReturnHC {
        Param (
            $Actual,
            $Script
        )
    
        $Actual.Size | Should -BeExactly $Script.LimitNew -Because "'LimitNew' returned by the script is not the same as the current 'Size'"
        $Actual.Softlimit | Should -BeExactly (Convert-QuotaStringToSoftlimitHC $Script.TypeNew) -Because "'TypeNew' returned by the script is not the same as the current 'Softlimit'"
        $Actual.Usage | Should -BeExactly $Script.Usage -Because "'Usage' returned by the script is not the same"
        $Actual.MatchesTemplate | Should -BeExactly $Script.MatchesTemplate -Because "'MatchesTemplate' returned by the script is not the same"
        $Actual.Template | Should -BeExactly $Script.Template -Because "'Template' returned by the script is not the same"
        $Actual.Description | Should -BeExactly $Script.Description -Because "'Description' returned by the script is not the same"
    }
    
    #region Create SourceTemplate
    Try {
        Remove-FsrmQuotaTemplate -Name $SourceTemplateName -Confirm:$false -EA Ignore
        $FsrmQuotaTemplate = New-FsrmQuotaTemplate -Name $SourceTemplateName -Size 10GB -SoftLimit
    }
    Catch {
        throw "Please install the role 'FileServerResourceManager'. Windows Server 2012R2 or higher is required"
    }
    #endregion

    $TestCases = @(
        #region Test
        $Description = 'hard limit 5GB with a PSCustomObject'
        $Size = 10GB

        @{
            TestName = $Description
            User     = [PSCustomObject]@{
                Description  = $Description
                Limit        = 'Some text'
                Size         = $Size
                SoftLimit    = $false
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $null
            }
            Expected = @{
                Description     = $Description
                Disabled        = $false
                MatchesTemplate = $false
                Size            = $Size
                SoftLimit       = $false
                Template        = $SourceTemplateName
            }
        }
        #endregion

        #region Test
        $Description = 'hard limit 10GB with a HashTable'
        $Size = 10GB

        @{
            TestName = $Description
            User     = @{
                Description  = $Description
                Limit        = 'not important text'
                Size         = $Size
                SoftLimit    = $false
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $null
            }
            Expected = @{
                Description     = $Description
                Disabled        = $false
                MatchesTemplate = $false
                Size            = $Size
                SoftLimit       = $false
                Template        = $SourceTemplateName
            }
        }
        #endregion

        #region Test
        $Description = 'soft limit 32GB'
        $Size = 32GB

        @{
            TestName = $Description
            User     = @{
                Description  = $Description
                Limit        = $null
                Size         = $Size
                SoftLimit    = $true
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $null
            }

            Expected = @{
                Description     = $Description
                Disabled        = $false
                MatchesTemplate = $false
                Size            = $Size
                SoftLimit       = $true
                Template        = $SourceTemplateName
            }
        }
        #endregion

        #region Test
        $Description = 'hard limit 20GB with warning 70%'
        $Size = 20GB
        $Threshold = @(
            @{
                Percentage = 70
                Action     = @{
                    MailFrom         = 'sd@contoso.com'
                    MailTo           = '[Source Io Owner Email]'
                    Subject          = '[Quota Threshold]% quota threshold exceeded'
                    Body             = "User [Source Io Owner] has exceeded the [Quota Threshold]% " + `
                        "quota threshold for the quota on [Quota Path] on server [Server]. The quota limit " + `
                        "is [Quota Limit MB] MB, and " + "[Quota Used MB] MB currently is in use ([Quota " + `
                        "Used Percent]% of limit)."
                    Type             = 2
                    RunLimitInterval = 60
                }
            }
        )

        @{
            TestName = $Description
            User     = @{
                Description  = $Description
                Limit        = $null
                Size         = $Size
                SoftLimit    = $false
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $null
                Threshold    = $Threshold
            }
            Expected = @{
                Description     = $Description
                Disabled        = $false
                MatchesTemplate = $false
                Size            = $Size
                SoftLimit       = $false
                Template        = $SourceTemplateName
                Threshold       = $Threshold
            }
        }
        #endregion

        #region Test
        $Description = 'hard limit 50GB with warning 50% and 60%'
        $Size = 50GB
        $Threshold = @(
            @{
                Percentage = 50
                Action     = @{
                    MailFrom         = 'sd@contoso.com'
                    MailTo           = '[Source Io Owner Email]'
                    Subject          = '[Quota Threshold]% quota threshold exceeded'
                    Body             = "Be aware you reached the treshold"
                    Type             = 2
                    RunLimitInterval = 60
                }
            }
            @{
                Percentage = 60
                Action     = @{
                    MailFrom         = 'sd@contoso.com'
                    MailTo           = '[Source Io Owner Email]'
                    Subject          = '[Quota Threshold]% quota threshold exceeded'
                    Body             = "Be aware you reached the treshold"
                    Type             = 2
                    RunLimitInterval = 60
                }
            }
        )

        @{
            TestName = $Description
            User     = @{
                Description  = $Description
                Limit        = $null
                Size         = $Size
                SoftLimit    = $false
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $null
                Threshold    = $Threshold
            }
            Expected = @{
                Description     = $Description
                Disabled        = $false
                MatchesTemplate = $false
                Size            = $Size
                SoftLimit       = $false
                Template        = $SourceTemplateName
                Threshold       = $Threshold
            }
        }
        #endregion
    )

    $SourceTemplateQuota = @{
        Description     = $FsrmQuotaTemplate.Description
        Disabled        = $false
        MatchesTemplate = $true
        Size            = $FsrmQuotaTemplate.Size
        SoftLimit       = $FsrmQuotaTemplate.SoftLimit
        Template        = $FsrmQuotaTemplate.Name
        Threshold       = $FsrmQuotaTemplate.Threshold
    }

    $FsrmSetting = Get-FsrmSetting -Verbose:$false
    Set-FsrmSetting -SmtpServer 'test@contoso.com' -AdminEmailAddress $ScriptAdmin -Verbose:$false

    $HomeFolder = (New-Item 'TestDrive:\HOME' -ItemType Directory -Force).FullName
    $UserFolder = (New-Item 'TestDrive:\HOME\user' -ItemType Directory -Force).FullName
    # Set-Location $HomeFolder

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ScriptName  = 'Test'
        OU          = 'contoso.com'
        SQLDatabase = 'PowerShell TEST'
        Environment = 'Test'
    }

    New-FsrmAutoQuota -Path $HomeFolder -Template $SourceTemplateName
    Set-FsrmSetting -SmtpServer 'test@contoso.com' -AdminEmailAddress $ScriptAdmin -Verbose:$false
}
AfterAll {
    Set-FsrmSetting -InputObject $FsrmSetting -Ea ignore
    Remove-FsrmQuotaTemplate -Name $SourceTemplateName -Confirm:$false -Verbose:$false -EA Ignore
}

Describe 'throw a terminating error when' {
    Context 'no quota management is configured on the parent folder' {
        BeforeAll {
            Remove-FsrmAutoQuota -Path $HomeFolder -Confirm:$false -EA Ignore
        }
        AfterAll {
            New-FsrmAutoQuota -Path $HomeFolder -Template $SourceTemplateName
            Set-FsrmSetting -SmtpServer 'test@contoso.com' -AdminEmailAddress $ScriptAdmin -Verbose:$false
        }
        It 'Error: No auto apply quota template set on parent folder' {
            Get-FsrmAutoQuota -Path $HomeFolder -EA ignore | Should -BeNullOrEmpty
            $User = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $UserFolder
                Description  = $null
                Limit        = 'RemoveQuota'
                Size         = 0
                SoftLimit    = $false
                Threshold    = $null
            }
    
            { .$testScript -User $User } | Should -Throw -PassThru |
            Select-Object -ExpandProperty Exception |
            Should -BeLike '*No auto apply quota template set on parent folder'
        }        
    }
    Context 'a mandatory property is missing on the input object' {
        $TestCases = @(
            @{
                Name = 'ComputerName'
                User = @{
                    ComputerPath = $UserFolder
                    Limit        = '22GB'
                    Size         = 23622320128
                    SoftLimit    = $false
                }
            }
            @{
                Name = 'ComputerPath'
                User = @{
                    ComputerName = $env:COMPUTERNAME
                    Limit        = '22GB'
                    Size         = 23622320128
                    SoftLimit    = $false
                }
            }
            @{
                Name = 'Size'
                User = @{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder
                    Limit        = '22GB'
                    SoftLimit    = $false
                }
            }
            @{
                Name = 'SoftLimit'
                User = @{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder
                    Limit        = 'RemoveQuota'
                    Size         = 0
                }
            }
        )
        It "Error: The property '<Name>' is mandatory" -Foreach $TestCases {
            { .$testScript -User $User } | Should -Throw -PassThru |
            Select-Object -ExpandProperty Exception |
            Should -BeLike "*The property '$Name' is mandatory*"
        }
    }
    Context 'the SMTP server is not set when threshold warning mails are to be send' {
        AfterAll {
            Set-FsrmSetting -SmtpServer 'test@contoso.com' -AdminEmailAddress $ScriptAdmin -Verbose:$false
        }
        It 'Error: Please request SMTP relay' {
            $User = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $UserFolder
                Limit        = 'RemoveQuota'
                Size         = 0
                SoftLimit    = $false
                Threshold    = $null
            }
            Set-FsrmSetting -SmtpServer $null -Verbose:$false

            { .$testScript -User $User } | 
            Should -Not -Throw -Because 'Threshold is not used, so no SMTP server needed'

            $User.Threshold = @{
                Percentage = 50
                Action     = @{
                    MailFrom         = 'sd@contoso.com'
                    MailTo           = '[Source Io Owner Email]'
                    Subject          = '[Quota Threshold]% quota threshold exceeded'
                    Body             = "Be aware you reached the threshold"
                    Type             = 2
                    RunLimitInterval = 60
                }
            }

            { .$testScript -User $User } | Should -Throw -PassThru |
            Select-Object -ExpandProperty Exception | Should -BeLike "*Please request SMTP relay * configure the server for sending e-mails*" 
        }
    }
}
Describe 'write a non terminating error when' {
    It 'the user folder is not found' {
        Mock Write-Error

        $User = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            ComputerPath = Join-Path $TestDrive 'NotExistingFolder'
            Description  = $null
            Limit        = '22GB'
            Size         = 23622320128
            SoftLimit    = $false
            Threshold    = $null
        }

        .$testScript -User $User

        Should -Invoke Write-Error -Times 1 -Exactly -ParameterFilter {
            $message -like "Folder '$($user.ComputerPath)' not found*"
        }
    }
}
Describe "when a user is member of the 'RemoveQuota' group" {
    Context 'and a quota limit was set that does not match the parent folder quota' {
        BeforeAll {
            Remove-FsrmQuota -Path $UserFolder -Confirm:$false
            New-FsrmQuota -Path $UserFolder -Size 60GB
    
            $User = @{
                Limit        = 'RemoveQuota'
                Size         = 0
                SoftLimit    = $false
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $UserFolder
                Threshold    = $null
            }
            $Expected = @{
                Description     = ''
                Disabled        = $false
                MatchesTemplate = $true
                Size            = $FsrmQuotaTemplate.Size
                SoftLimit       = $FsrmQuotaTemplate.SoftLimit
                Template        = $FsrmQuotaTemplate.Name
                Threshold       = $FsrmQuotaTemplate.Threshold
            }
    
            $Result = .$testScript -User $User
        }
        It 'the previously set quota is removed and the quota from the parent is applied' {
            $Actual = Get-FsrmQuota -Path $User.ComputerPath
            Test-QuotaDetailsHC -Actual $Actual -Expected $Expected
            Test-ScriptReturnHC -Actual $Actual -Script $Result
        }
        It 'the changes are registered' {
            $Result.Status | Should -BeExactly 'Changed'
            $Result.Field | Should -BeLike '*ApplySourceTemplate*'
        }
    } 
    Context 'and the quota limit is already matching the parent folder quota' {
        BeforeAll {
            $User = @{
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $UserFolder
                Limit        = 'RemoveQuota'
                Size         = $FsrmQuotaTemplate.Size
                SoftLimit    = $false
            }

            Set-EqualToTemplateHC -Path $User.ComputerPath -Template $SourceTemplateName

            $Result = .$testScript -User $User
        }
        It 'the previously set quota remains unchanged' {
            $Actual = Get-FsrmQuota -Path $UserFolder
            Test-QuotaDetailsHC -Actual $Actual -Expected $SourceTemplateQuota
            Test-ScriptReturnHC -Actual $Actual -Script $Result
        } 
        It 'no changes are registered' {
            $Result.Status | Should -BeExactly 'Ok'
            $Result.Field | Should -BeNullOrEmpty
            $Result.Description | Should -BeNullOrEmpty
        }
    } 
}
Describe 'when quote limits are requested for a different computer' {
    BeforeAll {
        $User = @{
            ComputerName = 'AnotherServer'
            ComputerPath = $UserFolder
            Limit        = '20GB'
            Size         = 21474836480
            SoftLimit    = $false
        }

        $result = .$testScript -User $User
    }
    It 'they get filtered out' {
        $user | Should -HaveCount 0
    }
    It 'no actions are executed' {
        $result | Should -BeNullOrEmpty
        $ParentQuota | Should -BeNullOrEmpty
    }
}
Describe 'apply quota' {
    Describe 'changes in properties and fix them' {
        It 'No changes, change nothing' {
            $New = @{
                Path        = $UserFolder
                Template    = $SourceTemplateName
                Description = 'Description A'
                Size        = 5GB
                SoftLimit   = $false
            }
            Remove-FsrmQuota -Path $New.Path -Confirm:$false
            New-FsrmQuota @New

            $User = @{
                Description  = $New.Description
                Limit        = $null
                Size         = $New.Size
                SoftLimit    = $New.SoftLimit
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $New.Path
            }

            $Result = .$testScript -User $User

            $Result.Status | Should -BeExactly 'Ok'
            $Result.Field | Should -BeNullOrEmpty
            $Result.Size | Should -BeExactly $New.Size
            $Result.Description | Should -BeExactly $New.Description

            $Actual = Get-FsrmQuota -Path $User.ComputerPath
            $Actual.Size | Should -BeExactly $New.Size
            $Actual.Description | Should -BeExactly $New.Description
            $Actual.SoftLimit | Should -BeExactly $New.SoftLimit
        } 
        Context 'Description' {
            It 'Set quota with Description' {
                $NewQuotaParams = @{
                    Path        = $UserFolder
                    Template    = $SourceTemplateName
                    Description = 'Description A'
                    Size        = 5GB
                    SoftLimit   = $false
                }
                Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                New-FsrmQuota @NewQuotaParams

                $User = @{
                    Description  = 'Description B'
                    Limit        = $null
                    Size         = $NewQuotaParams.Size
                    SoftLimit    = $NewQuotaParams.SoftLimit
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $NewQuotaParams.Path
                }
                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Changed'
                $Result.Field | Should -BeExactly 'Description'
                $Result.Description | Should -BeExactly 'Description B'
            } 
            It 'Set quota without Description' {
                $NewQuotaParams = @{
                    Path        = $UserFolder
                    Template    = $SourceTemplateName
                    Description = 'Description A'
                    Size        = 5GB
                    SoftLimit   = $false
                }
                Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                New-FsrmQuota @NewQuotaParams

                $User = @{
                    Limit        = $null
                    Size         = $NewQuotaParams.Size
                    SoftLimit    = $NewQuotaParams.SoftLimit
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $NewQuotaParams.Path
                }
                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Changed'
                $Result.Field | Should -BeExactly 'Description'
                $Result.Description | Should -BeNullOrEmpty
            } 
            It 'Set quota with Description $null' {
                $NewQuotaParams = @{
                    Path        = $UserFolder
                    Template    = $SourceTemplateName
                    Description = 'Description A'
                    Size        = 5GB
                    SoftLimit   = $false
                }
                Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                New-FsrmQuota @NewQuotaParams

                $User = @{
                    Description  = $null
                    Limit        = $null
                    Size         = $NewQuotaParams.Size
                    SoftLimit    = $NewQuotaParams.SoftLimit
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $NewQuotaParams.Path
                }
                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Changed'
                $Result.Field | Should -BeExactly 'Description'
                $Result.Description | Should -BeNullOrEmpty
            } 
            It 'Set quota with Description empty' {
                $NewQuotaParams = @{
                    Path        = $UserFolder
                    Template    = $SourceTemplateName
                    Description = 'Description A'
                    Size        = 5GB
                    SoftLimit   = $false
                }
                Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                New-FsrmQuota @NewQuotaParams

                $User = @{
                    Description  = ''
                    Limit        = $null
                    Size         = $NewQuotaParams.Size
                    SoftLimit    = $NewQuotaParams.SoftLimit
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $NewQuotaParams.Path
                }
                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Changed'
                $Result.Field | Should -BeExactly 'Description'
                $Result.Description | Should -BeNullOrEmpty
            } 
            Context "user member of 'RemoveQuota' group so 'Description' will be blank" {
                It 'Description requested' {
                    $NewQuotaParams = @{
                        Path        = $UserFolder
                        Template    = $SourceTemplateName
                        Description = 'Description A'
                        Size        = 5GB
                        SoftLimit   = $false
                    }
                    Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                    New-FsrmQuota @NewQuotaParams

                    $User = @{
                        Description  = 'Description not applied'
                        Limit        = 'RemoveQuota'
                        Size         = $NewQuotaParams.Size
                        SoftLimit    = $NewQuotaParams.SoftLimit
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $NewQuotaParams.Path
                    }
                    $Result = .$testScript -User $User

                    $Result.Description | Should -BeNullOrEmpty
                } 
                It 'No Description requested' {
                    $NewQuotaParams = @{
                        Path        = $UserFolder
                        Template    = $SourceTemplateName
                        Description = 'Description A'
                        Size        = 5GB
                        SoftLimit   = $false
                    }
                    Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                    New-FsrmQuota @NewQuotaParams

                    $User = @{
                        Limit        = 'RemoveQuota'
                        Size         = $NewQuotaParams.Size
                        SoftLimit    = $NewQuotaParams.SoftLimit
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $NewQuotaParams.Path
                    }
                    $Result = .$testScript -User $User

                    $Result.Description | Should -BeNullOrEmpty
                } 
                It 'No Description before, no Description afterwards' {
                    $NewQuotaParams = @{
                        Path      = $UserFolder
                        Template  = $SourceTemplateName
                        #Description = 'Description A'
                        Size      = 5GB
                        SoftLimit = $false
                    }
                    Remove-FsrmQuota -Path $NewQuotaParams.Path -Confirm:$false
                    New-FsrmQuota @NewQuotaParams

                    $User = @{
                        Limit        = 'RemoveQuota'
                        Size         = $NewQuotaParams.Size
                        SoftLimit    = $NewQuotaParams.SoftLimit
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $NewQuotaParams.Path
                    }
                    $Result = .$testScript -User $User

                    $Result.Description | Should -BeNullOrEmpty
                } 
            }
        }
        It 'Limit' {
            $NewLimit = 9GB

            $New = @{
                Path      = $UserFolder
                Template  = $SourceTemplateName
                Size      = 5GB
                SoftLimit = $false
            }
            Remove-FsrmQuota -Path $New.Path -Confirm:$false
            New-FsrmQuota @New

            $User = @{
                Limit        = $null
                Size         = $NewLimit
                SoftLimit    = $New.SoftLimit
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $New.Path
            }

            $Result = .$testScript -User $User

            $Result.Status | Should -BeExactly 'Changed'
            $Result.Field | Should -BeExactly 'Limit'
            $Result.Size | Should -BeExactly $NewLimit

            $Actual = Get-FsrmQuota -Path $User.ComputerPath
            $Actual.Size | Should -BeExactly $NewLimit
            $Actual.SoftLimit | Should -BeExactly $New.SoftLimit
        } 
        It 'Type' {
            $NewSoftLimit = $true

            $New = @{
                Path        = $UserFolder
                Template    = $SourceTemplateName
                Description = 'Description A'
                Size        = 5GB
                SoftLimit   = $false
            }
            Remove-FsrmQuota -Path $New.Path -Confirm:$false
            New-FsrmQuota @New

            $User = @{
                Description  = $New.Description
                Limit        = $null
                Size         = $New.Size
                SoftLimit    = $NewSoftLimit
                ComputerName = $env:COMPUTERNAME
                ComputerPath = $New.Path
            }

            $Result = .$testScript -User $User

            $Result.Status | Should -BeExactly 'Changed'
            $Result.Field | Should -BeExactly 'Type'
            $Result.TypeNew | Should -BeExactly 'Soft'
            $Result.TypeOld | Should -BeExactly 'Hard'

            $Actual = Get-FsrmQuota -Path $User.ComputerPath
            $Actual.SoftLimit | Should -BeExactly $NewSoftLimit
        } 
        Context 'Threshold' {
            It 'Correct, nothing to do' {
                $Threshold = @(
                    @{
                        Percentage = 70
                        Action     = @{
                            MailTo           = '[Source Io Owner Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                    @{
                        Percentage = 90
                        Action     = @{
                            MailTo           = '[Admin Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                )

                $Q1Params = @{
                    Type             = 'Email'
                    MailTo           = $Threshold[0].Action.MailTo
                    Subject          = $Threshold[0].Action.Subject
                    Body             = $Threshold[0].Action.Body
                    RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                }

                $Q2Params = @{
                    Type             = 'Email'
                    MailTo           = $Threshold[1].Action.MailTo
                    Subject          = $Threshold[1].Action.Subject
                    Body             = $Threshold[1].Action.Body
                    RunLimitInterval = $Threshold[1].Action.RunLimitInterval
                }

                $New = @{
                    Path      = $UserFolder
                    Template  = $SourceTemplateName
                    Size      = 20GB
                    SoftLimit = $false
                    Threshold = @(
                        (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                        (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                    )
                }
                Remove-FsrmQuota -Path $New.Path -Confirm:$false
                $Before = New-FsrmQuota @New

                $User = [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $New.Path
                    Limit        = $null
                    Size         = $New.Size
                    SoftLimit    = $New.SoftLimit
                    Threshold    = $Threshold
                }

                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Ok'
                $Result.Field | Should -BeNullOrEmpty

                $After = Get-FsrmQuota -Path $User.ComputerPath
                Test-QuotaDetailsHC -Actual $Before -Expected $After
                Test-ScriptReturnHC -Actual $After -Script $Result
            } 
            It 'Remove' {
                $Threshold = @(
                    @{
                        Percentage = 70
                        Action     = @{
                            MailTo           = '[Source Io Owner Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                    @{
                        Percentage = 90
                        Action     = @{
                            MailTo           = '[Admin Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                )

                $Q1Params = @{
                    Type             = 'Email'
                    MailTo           = $Threshold[0].Action.MailTo
                    Subject          = $Threshold[0].Action.Subject
                    Body             = $Threshold[0].Action.Body
                    RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                }

                $Q2Params = @{
                    Type             = 'Email'
                    MailTo           = $Threshold[1].Action.MailTo
                    Subject          = $Threshold[1].Action.Subject
                    Body             = $Threshold[1].Action.Body
                    RunLimitInterval = $Threshold[1].Action.RunLimitInterval
                }

                $New = @{
                    Path      = $UserFolder
                    Template  = $SourceTemplateName
                    Size      = 20GB
                    SoftLimit = $false
                    Threshold = @(
                        (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                        (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                    )
                }
                Remove-FsrmQuota -Path $New.Path -Confirm:$false
                $Before = New-FsrmQuota @New

                $User = [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $New.Path
                    Limit        = $null
                    Size         = $New.Size
                    SoftLimit    = $New.SoftLimit
                    Threshold    = $Threshold[0]
                }

                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Changed'
                $Result.Field | Should -BeExactly "RemoveWarning$($Threshold[1].Percentage)"
                $Result.WarningOld | Should -BeExactly ('{0}, {1}' -f $Threshold[0].Percentage, $Threshold[1].Percentage)
                $Result.WarningNew | Should -BeExactly ('{0}' -f $Threshold[0].Percentage)

                $After = Get-FsrmQuota -Path $User.ComputerPath
                $After.Threshold.Percentage | Should -BeExactly $Threshold[0].Percentage

                Test-ScriptReturnHC -Actual $After -Script $Result
            } 
            It 'Add' {
                $Threshold = @(
                    @{
                        Percentage = 70
                        Action     = @{
                            MailTo           = '[Source Io Owner Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                )

                $Q1Params = @{
                    Type             = 'Email'
                    MailTo           = $Threshold[0].Action.MailTo
                    Subject          = $Threshold[0].Action.Subject
                    Body             = $Threshold[0].Action.Body
                    RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                }

                $New = @{
                    Path      = $UserFolder
                    Template  = $SourceTemplateName
                    Size      = 20GB
                    SoftLimit = $false
                    Threshold = @(
                        (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params))
                    )
                }
                Remove-FsrmQuota -Path $New.Path -Confirm:$false
                $Before = New-FsrmQuota @New

                $Threshold += @(
                    @{
                        Percentage = 95
                        Action     = @{
                            MailTo           = '[Source Io Owner Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                )

                $User = [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $New.Path
                    Limit        = $null
                    Size         = $New.Size
                    SoftLimit    = $New.SoftLimit
                    Threshold    = $Threshold
                }

                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Changed'
                $Result.Field | Should -BeExactly "AddWarning$($Threshold[1].Percentage)"
                $Result.WarningOld | Should -BeExactly ('{0}' -f $Threshold[0].Percentage)
                $Result.WarningNew | Should -BeExactly ('{0}, {1}' -f $Threshold[0].Percentage, $Threshold[1].Percentage)

                $After = Get-FsrmQuota -Path $User.ComputerPath
                $After.Threshold.Percentage | Should -Contain $Threshold[0].Percentage
                $After.Threshold.Percentage | Should -Contain $Threshold[1].Percentage

                Test-ScriptReturnHC -Actual $After -Script $Result
            } 
            Context 'Modify' {
                It 'MailTo' {
                    $Threshold = @(
                        @{
                            Percentage = 70
                            Action     = @{
                                MailTo           = '[Source Io Owner Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                        @{
                            Percentage = 90
                            Action     = @{
                                MailTo           = '[Admin Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                    )

                    $Q1Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[0].Action.MailTo
                        Subject          = $Threshold[0].Action.Subject
                        Body             = $Threshold[0].Action.Body
                        RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                    }

                    $Q2Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[1].Action.MailTo
                        Subject          = $Threshold[1].Action.Subject
                        Body             = $Threshold[1].Action.Body
                        RunLimitInterval = $Threshold[1].Action.RunLimitInterval
                    }

                    $New = @{
                        Path      = $UserFolder
                        Template  = $SourceTemplateName
                        Size      = 20GB
                        SoftLimit = $false
                        Threshold = @(
                            (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                            (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                        )
                    }
                    Remove-FsrmQuota -Path $New.Path -Confirm:$false
                    $Before = New-FsrmQuota @New

                    $Threshold[0].Action.MailTo = 'chuck@norris.com'

                    $User = [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $New.Path
                        Limit        = $null
                        Size         = $New.Size
                        SoftLimit    = $New.SoftLimit
                        Threshold    = $Threshold
                    }

                    $Result = .$testScript -User $User

                    $Result.Status | Should -BeExactly 'Changed'
                    $Result.Field | Should -BeExactly "ModifyWarning$($Threshold[0].Percentage)"

                    $After = Get-FsrmQuota -Path $User.ComputerPath
                    $After.Threshold[0].Action.MailTo | Should -BeExactly $Threshold[0].Action.MailTo
                    $After.Threshold[1].Action.MailTo | Should -BeExactly $Threshold[1].Action.MailTo
                    Test-ScriptReturnHC -Actual $After -Script $Result
                } 
                It 'Subject' {
                    $Threshold = @(
                        @{
                            Percentage = 70
                            Action     = @{
                                MailTo           = '[Source Io Owner Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                        @{
                            Percentage = 90
                            Action     = @{
                                MailTo           = '[Admin Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                    )

                    $Q1Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[0].Action.MailTo
                        Subject          = $Threshold[0].Action.Subject
                        Body             = $Threshold[0].Action.Body
                        RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                    }

                    $Q2Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[1].Action.MailTo
                        Subject          = $Threshold[1].Action.Subject
                        Body             = $Threshold[1].Action.Body
                        RunLimitInterval = $Threshold[1].Action.RunLimitInterval
                    }

                    $New = @{
                        Path      = $UserFolder
                        Template  = $SourceTemplateName
                        Size      = 20GB
                        SoftLimit = $false
                        Threshold = @(
                            (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                            (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                        )
                    }
                    Remove-FsrmQuota -Path $New.Path -Confirm:$false
                    $Before = New-FsrmQuota @New

                    $Threshold[0].Action.Subject = 'Another topic'

                    $User = [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $New.Path
                        Limit        = $null
                        Size         = $New.Size
                        SoftLimit    = $New.SoftLimit
                        Threshold    = $Threshold
                    }

                    $Result = .$testScript -User $User

                    $Result.Status | Should -BeExactly 'Changed'
                    $Result.Field | Should -BeExactly "ModifyWarning$($Threshold[0].Percentage)"

                    $After = Get-FsrmQuota -Path $User.ComputerPath
                    $After.Threshold[0].Action.Subject | Should -BeExactly $Threshold[0].Action.Subject
                    $After.Threshold[1].Action.Subject | Should -BeExactly $Threshold[1].Action.Subject
                    Test-ScriptReturnHC -Actual $After -Script $Result
                } 
                It 'Body' {
                    $Threshold = @(
                        @{
                            Percentage = 70
                            Action     = @{
                                MailTo           = '[Source Io Owner Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                        @{
                            Percentage = 90
                            Action     = @{
                                MailTo           = '[Admin Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                    )

                    $Q1Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[0].Action.MailTo
                        Subject          = $Threshold[0].Action.Subject
                        Body             = $Threshold[0].Action.Body
                        RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                    }

                    $Q2Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[1].Action.MailTo
                        Subject          = $Threshold[1].Action.Subject
                        Body             = $Threshold[1].Action.Body
                        RunLimitInterval = $Threshold[1].Action.RunLimitInterval
                    }

                    $New = @{
                        Path      = $UserFolder
                        Template  = $SourceTemplateName
                        Size      = 20GB
                        SoftLimit = $false
                        Threshold = @(
                            (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                            (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                        )
                    }
                    Remove-FsrmQuota -Path $New.Path -Confirm:$false
                    $Before = New-FsrmQuota @New

                    $Threshold[0].Action.Body = 'Another text in the body'

                    $User = [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $New.Path
                        Limit        = $null
                        Size         = $New.Size
                        SoftLimit    = $New.SoftLimit
                        Threshold    = $Threshold
                    }

                    $Result = .$testScript -User $User

                    $Result.Status | Should -BeExactly 'Changed'
                    $Result.Field | Should -BeExactly "ModifyWarning$($Threshold[0].Percentage)"

                    $After = Get-FsrmQuota -Path $User.ComputerPath
                    $After.Threshold[0].Action.Body | Should -BeExactly $Threshold[0].Action.Body
                    $After.Threshold[1].Action.Body | Should -BeExactly $Threshold[1].Action.Body
                    Test-ScriptReturnHC -Actual $After -Script $Result
                } 
                It 'RunLimitInterval' {
                    $Threshold = @(
                        @{
                            Percentage = 70
                            Action     = @{
                                MailTo           = '[Source Io Owner Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                        @{
                            Percentage = 90
                            Action     = @{
                                MailTo           = '[Admin Email]'
                                Subject          = '[Quota Threshold]% quota threshold exceeded'
                                Body             = "Be aware you reached the treshold"
                                Type             = 2
                                RunLimitInterval = 60
                            }
                        }
                    )

                    $Q1Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[0].Action.MailTo
                        Subject          = $Threshold[0].Action.Subject
                        Body             = $Threshold[0].Action.Body
                        RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                    }

                    $Q2Params = @{
                        Type             = 'Email'
                        MailTo           = $Threshold[1].Action.MailTo
                        Subject          = $Threshold[1].Action.Subject
                        Body             = $Threshold[1].Action.Body
                        RunLimitInterval = $Threshold[1].Action.RunLimitInterval
                    }

                    $New = @{
                        Path      = $UserFolder
                        Template  = $SourceTemplateName
                        Size      = 20GB
                        SoftLimit = $false
                        Threshold = @(
                            (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                            (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                        )
                    }
                    Remove-FsrmQuota -Path $New.Path -Confirm:$false
                    $Before = New-FsrmQuota @New

                    $Threshold[0].Action.RunLimitInterval = 120

                    $User = [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        ComputerPath = $New.Path
                        Limit        = $null
                        Size         = $New.Size
                        SoftLimit    = $New.SoftLimit
                        Threshold    = $Threshold
                    }

                    $Result = .$testScript -User $User

                    $Result.Status | Should -BeExactly 'Changed'
                    $Result.Field | Should -BeExactly "ModifyWarning$($Threshold[0].Percentage)"

                    $After = Get-FsrmQuota -Path $User.ComputerPath
                    $After.Threshold[0].Action.RunLimitInterval | Should -BeExactly $Threshold[0].Action.RunLimitInterval
                    $After.Threshold[1].Action.RunLimitInterval | Should -BeExactly $Threshold[1].Action.RunLimitInterval
                    Test-ScriptReturnHC -Actual $After -Script $Result
                } 
            }
        }
        Context 'special' {
            It 'ignore leading and trailing spaces in properties' {
                $Threshold = @(
                    @{
                        Percentage = 70
                        Action     = @{
                            MailTo           = '[Source Io Owner Email]'
                            Subject          = '[Quota Threshold]% quota threshold exceeded'
                            Body             = "Be aware you reached the treshold"
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                )

                $Q1Params = @{
                    Type             = 'Email'
                    MailTo           = $Threshold[0].Action.MailTo
                    Subject          = $Threshold[0].Action.Subject
                    Body             = $Threshold[0].Action.Body
                    RunLimitInterval = $Threshold[0].Action.RunLimitInterval
                }

                $New = @{
                    Path        = $UserFolder
                    Description = 'The text'
                    Template    = $SourceTemplateName
                    Size        = 20GB
                    SoftLimit   = $false
                    Threshold   = @(
                        (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params))
                    )
                }
                Remove-FsrmQuota -Path $New.Path -Confirm:$false
                $Before = New-FsrmQuota @New

                $Threshold = @(
                    @{
                        Percentage = 70
                        Action     = @{
                            MailTo           = '   ' + $Threshold[0].Action.MailTo + '   '
                            Subject          = '   ' + $Threshold[0].Action.Subject + '   '
                            Body             = '   ' + $Threshold[0].Action.Body + '   '
                            Type             = 2
                            RunLimitInterval = 60
                        }
                    }
                )

                $User = [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $New.Path
                    Description  = '    ' + $New.Description + '   '
                    Limit        = $null
                    Size         = $New.Size
                    SoftLimit    = $New.SoftLimit
                    Threshold    = $Threshold
                }

                $Result = .$testScript -User $User

                $Result.Status | Should -BeExactly 'Ok'
                $Result.Field | Should -BeNullOrEmpty

                $After = Get-FsrmQuota -Path $User.ComputerPath
                Test-QuotaDetailsHC -Actual $Before -Expected $After
                Test-ScriptReturnHC -Actual $After -Script $Result
            } 
        }
    }
    Describe 'returned objects' {
        It 'extra properties passed to the script are returned' {
            Remove-FsrmQuota -Path $UserFolder -Confirm:$false
            New-FsrmQuota -Path $UserFolder -Size 60GB

            $User = @{
                ComputerName   = $env:COMPUTERNAME
                ComputerPath   = $UserFolder
                Limit          = 'RemoveQuota'
                Size           = 0
                SoftLimit      = $false
                ExtraProperty1 = 1
                ExtraProperty2 = 2
            }

            $Result = .$testScript -User $User

            $Result.ExtraProperty1 | Should -BeExactly 1
            $Result.ExtraProperty2 | Should -BeExactly 2
        } 
        It "have type 'PSCustomObject'" {
            Remove-FsrmQuota -Path $UserFolder -Confirm:$false
            New-FsrmQuota -Path $UserFolder -Size 60GB

            $User = @{
                ComputerName   = $env:COMPUTERNAME
                ComputerPath   = $UserFolder
                Limit          = 'RemoveQuota'
                Size           = 0
                SoftLimit      = $false
                ExtraProperty1 = 1
                ExtraProperty2 = 2
            }

            $Result = .$testScript -User $User

            $Result.GetType().Name | Should -BeExactly PSCustomObject
        } 
    }
    Context "SourceTemplate '$SourceTemplateName' applied" {
        It 'Set new quota <TestName>' -TestCases $TestCases {
            Param (
                $TestName,
                $User,
                $Expected
            )

            $User.ComputerPath = $UserFolder
            Set-EqualToTemplateHC -path $User.ComputerPath -Template $SourceTemplateName

            $Actual = Get-FsrmQuota -Path $User.ComputerPath
            Test-QuotaDetailsHC -Actual $Actual -Expected $SourceTemplateQuota -When ExpectedBefore

            $Result = .$testScript -User $User

            $Actual = Get-FsrmQuota -Path $User.ComputerPath
            Test-QuotaDetailsHC -Actual $Actual -Expected $Expected
            Test-ScriptReturnHC -Actual $Actual -Script $Result
        } 
    }
    Context 'multiple users at the same time' {
        BeforeAll {
            $UserFolder1 = (New-Item (Join-Path $HomeFolder '1') -ItemType Directory -Force).FullName
            $UserFolder2 = (New-Item (Join-Path $HomeFolder '2') -ItemType Directory -Force).FullName
            
            # Wait for quota's to be applied
            Start-Sleep -Seconds 3
        }
        It '2 users in hashtable' {
            $User = @(
                @{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder1
                    Limit        = $null
                    Size         = 5GB
                    SoftLimit    = $false
                }
                @{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder2
                    Limit        = 'RemoveQuota'
                    Size         = 0
                    SoftLimit    = $false
                }
            )

            { .$testScript -User $User } | Should -Not -Throw
        } 
        It '2 users as PSCustomObject' {
            $User = @(
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder1
                    Limit        = $null
                    Size         = 5GB
                    SoftLimit    = $false
                }
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder2
                    Limit        = 'RemoveQuota'
                    Size         = 0
                    SoftLimit    = $false
                }
            )

            { .$testScript -User $User } | Should -Not -Throw
        } 
        It '2 users as PSCustomObject with Threshold' {
            $Threshold = @(
                [PSCustomObject]@{
                    Percentage = 80
                    Color      = 'Pink'
                    Action     = @{
                        MailFrom         = 'cnorris@conotoso.com'
                        MailTo           = '[Source Io Owner Email]'
                        MailCc           = ''
                        MailBcc          = $null
                        MailReplyTo      = $null
                        Subject          = '[Quota Threshold]% quota threshold exceeded'
                        Body             = "Be aware you reached the treshold"
                        Type             = 2
                        RunLimitInterval = 60
                    }
                }
                [PSCustomObject]@{
                    Percentage = 90
                    Color      = 'Gray'
                    Action     = @{
                        MailFrom         = 'cnorris@conotoso.com'
                        MailTo           = '[Admin Email]'
                        MailCc           = $null
                        MailBcc          = $null
                        MailReplyTo      = $null
                        Subject          = '[Quota Threshold]% quota threshold exceeded'
                        Body             = "Be aware you reached the treshold"
                        Type             = 2
                        RunLimitInterval = 60
                    }
                }
                [PSCustomObject]@{
                    Percentage = 100
                    Color      = 'Gray'
                    Action     = @{
                        MailFrom         = 'cnorris@conotoso.com'
                        MailTo           = '[Admin Email]'
                        MailCc           = $null
                        MailBcc          = $null
                        MailReplyTo      = $null
                        Subject          = '[Quota Threshold]% quota threshold exceeded'
                        Body             = "Be aware you reached the treshold"
                        Type             = 2
                        RunLimitInterval = 60
                    }
                }
            )

            $User = @(
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder1
                    Limit        = $null
                    Size         = 5GB
                    SoftLimit    = $false
                    Threshold    = $Threshold
                }
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    ComputerPath = $UserFolder2
                    Limit        = 'RemoveQuota'
                    Size         = 0
                    SoftLimit    = $false
                    Threshold    = $Threshold
                }
            )

            { .$testScript -User $User } | Should -Not -Throw
        } 
        It 'JSON AD data, add quota for 2 users' {
            $JSONtestDataAD = @"
[
    {
        "DistinguishedName":  "CN=Schuermans\\, Natascha (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net",
        "Enabled":  true,
        "GivenName":  "Natascha",
        "HomeDirectory":  "\\\\grouphc.net\\BNL\\HOME\\Centralized\\nschuerm",
        "Name":  "Schuermans, Natascha (Braine L’Alleud) BEL",
        "ObjectClass":  "user",
        "ObjectGUID":  "5ef73093-4907-40b0-a21f-c98e191fd545",
        "SamAccountName":  "nschuerm",
        "SID":  {
                    "BinaryLength":  28,
                    "AccountDomainSid":  "S-1-5-21-1078081533-261478967-839522115",
                    "Value":  "S-1-5-21-1078081533-261478967-839522115-680611"
                },
        "Surname":  "Schuermans",
        "UserPrincipalName":  "nschuerm@grouphc.net",
        "PropertyNames":  [
                              "DistinguishedName",
                              "Enabled",
                              "GivenName",
                              "HomeDirectory",
                              "Name",
                              "ObjectClass",
                              "ObjectGUID",
                              "SamAccountName",
                              "SID",
                              "Surname",
                              "UserPrincipalName"
                          ],
        "AddedProperties":  [

                            ],
        "RemovedProperties":  [

                              ],
        "ModifiedProperties":  [

                               ],
        "PropertyCount":  11,
        "GroupName":  "BEL ATT Quota home 120GB",
        "Limit":  "120GB",
        "Size":  128849018880,
        "Description":  "BEL ATT Quota home 120GB #PowerShellManaged",
        "SoftLimit":  false,
        "Threshold":  null,
        "ComputerName":  "DEUSFFRAN0008.grouphc.net",
        "ComputerPath":  "E:\\HOME\\nschuerm"
    },
    {
        "DistinguishedName":  "CN=Gijbels\\, Brecht (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net",
        "Enabled":  true,
        "GivenName":  "Brecht",
        "HomeDirectory":  "\\\\grouphc.net\\BNL\\HOME\\Centralized\\gijbelsb",
        "Name":  "Gijbels, Brecht (Braine L’Alleud) BEL",
        "ObjectClass":  "user",
        "ObjectGUID":  "c1a3e1e7-cfc8-4569-885c-4d6bd2f6c433",
        "SamAccountName":  "gijbelsb",
        "SID":  {
                    "BinaryLength":  28,
                    "AccountDomainSid":  "S-1-5-21-1078081533-261478967-839522115",
                    "Value":  "S-1-5-21-1078081533-261478967-839522115-138487"
                },
        "Surname":  "Gijbels",
        "UserPrincipalName":  "gijbelsb@grouphc.net",
        "PropertyNames":  [
                              "DistinguishedName",
                              "Enabled",
                              "GivenName",
                              "HomeDirectory",
                              "Name",
                              "ObjectClass",
                              "ObjectGUID",
                              "SamAccountName",
                              "SID",
                              "Surname",
                              "UserPrincipalName"
                          ],
        "AddedProperties":  [

                            ],
        "RemovedProperties":  [

                              ],
        "ModifiedProperties":  [

                               ],
        "PropertyCount":  11,
        "GroupName":  "BEL ATT Quota home 150GB",
        "Limit":  "150GB",
        "Size":  161061273600,
        "Description":  "BEL ATT Quota home 150GB #PowerShellManaged",
        "SoftLimit":  false,
        "Threshold":  null,
        "ComputerName":  "DEUSFFRAN0008.grouphc.net",
        "ComputerPath":  "E:\\HOME\\gijbelsb"
    }
]
"@
            $User = $JSONtestDataAD -join "`n" | ConvertFrom-Json

            $User[0].ComputerName = $env:COMPUTERNAME
            $User[0].ComputerPath = $UserFolder1

            $User[1].ComputerName = $env:COMPUTERNAME
            $User[1].ComputerPath = $UserFolder2
            { .$testScript -User $User } | Should -Not -Throw
        } 
        It 'JSON AD data, one user no limit, one user with' {
            $JSONtestDataAD = @"
[
    {
        "DistinguishedName":  "CN=Schuermans\\, Natascha (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net",
        "Enabled":  true,
        "GivenName":  "Natascha",
        "HomeDirectory":  "\\\\grouphc.net\\BNL\\HOME\\Centralized\\nschuerm",
        "Name":  "Schuermans, Natascha (Braine L’Alleud) BEL",
        "ObjectClass":  "user",
        "ObjectGUID":  "5ef73093-4907-40b0-a21f-c98e191fd545",
        "SamAccountName":  "nschuerm",
        "SID":  {
                    "BinaryLength":  28,
                    "AccountDomainSid":  "S-1-5-21-1078081533-261478967-839522115",
                    "Value":  "S-1-5-21-1078081533-261478967-839522115-680611"
                },
        "Surname":  "Schuermans",
        "UserPrincipalName":  "nschuerm@grouphc.net",
        "PropertyNames":  [
                              "DistinguishedName",
                              "Enabled",
                              "GivenName",
                              "HomeDirectory",
                              "Name",
                              "ObjectClass",
                              "ObjectGUID",
                              "SamAccountName",
                              "SID",
                              "Surname",
                              "UserPrincipalName"
                          ],
        "AddedProperties":  [

                            ],
        "RemovedProperties":  [

                              ],
        "ModifiedProperties":  [

                               ],
        "PropertyCount":  11,
        "GroupName":  "BEL ATT Quota home 120GB",
        "Limit":  "120GB",
        "Size":  128849018880,
        "Description":  "BEL ATT Quota home 120GB #PowerShellManaged",
        "SoftLimit":  false,
        "Threshold":  null,
        "ComputerName":  "DEUSFFRAN0008.grouphc.net",
        "ComputerPath":  "E:\\HOME\\nschuerm"
    },
    {
        "DistinguishedName":  "CN=Gijbels\\, Brecht (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net",
        "Enabled":  true,
        "GivenName":  "Brecht",
        "HomeDirectory":  "\\\\grouphc.net\\BNL\\HOME\\Centralized\\gijbelsb",
        "Name":  "Gijbels, Brecht (Braine L’Alleud) BEL",
        "ObjectClass":  "user",
        "ObjectGUID":  "c1a3e1e7-cfc8-4569-885c-4d6bd2f6c433",
        "SamAccountName":  "gijbelsb",
        "SID":  {
                    "BinaryLength":  28,
                    "AccountDomainSid":  "S-1-5-21-1078081533-261478967-839522115",
                    "Value":  "S-1-5-21-1078081533-261478967-839522115-138487"
                },
        "Surname":  "Gijbels",
        "UserPrincipalName":  "gijbelsb@grouphc.net",
        "PropertyNames":  [
                              "DistinguishedName",
                              "Enabled",
                              "GivenName",
                              "HomeDirectory",
                              "Name",
                              "ObjectClass",
                              "ObjectGUID",
                              "SamAccountName",
                              "SID",
                              "Surname",
                              "UserPrincipalName"
                          ],
        "AddedProperties":  [

                            ],
        "RemovedProperties":  [

                              ],
        "ModifiedProperties":  [

                               ],
        "PropertyCount":  11,
        "GroupName":  "BEL ATT Quota home 150GB",
        "Limit":  "150GB",
        "Size":  161061273600,
        "Description":  "BEL ATT Quota home 150GB #PowerShellManaged",
        "SoftLimit":  false,
        "Threshold":  null,
        "ComputerName":  "DEUSFFRAN0008.grouphc.net",
        "ComputerPath":  "E:\\HOME\\gijbelsb"
    }
]
"@
            $User = $JSONtestDataAD -join "`n" | ConvertFrom-Json

            $User[0].ComputerName = $env:COMPUTERNAME
            $User[0].ComputerPath = $UserFolder1

            $User[1].ComputerName = $env:COMPUTERNAME
            $User[1].ComputerPath = $UserFolder2

            Remove-FsrmQuota -Path $UserFolder1 -Confirm:$false
            $NewParams = @{
                Path        = $UserFolder1
                Size        = 5GB
                Description = 'BEL ATT Quota home 5GB #PowerShellManaged'
            }
            New-FsrmQuota @NewParams

            #Remove-FsrmQuota -Path $UserFolder2 -Confirm:$false
            #$NewParams = @{
            #    Path = $UserFolder2
            #    Size = 5GB
            #    Description = 'BEL ATT Quota home 5GB #PowerShellManaged'
            #}
            #New-FsrmQuota @NewParams

            { .$testScript -User $User } | Should -Not -Throw
        } 
        It 'JSON AD data, with thresholds' {
            $Threshold = @(
                @{
                    Percentage = 80
                    Action     = @{
                        MailTo           = '[Source Io Owner Email]'
                        Subject          = '[Quota Threshold]% quota threshold exceeded'
                        Body             = "Be aware you reached the treshold"
                        Type             = 2
                        RunLimitInterval = 60
                    }
                }
                @{
                    Percentage = 90
                    Action     = @{
                        MailTo           = '[Admin Email]'
                        Subject          = '[Quota Threshold]% quota threshold exceeded'
                        Body             = "Be aware you reached the treshold"
                        Type             = 2
                        RunLimitInterval = 60
                    }
                }
            )

            $Q1Params = @{
                Type             = 'Email'
                MailTo           = $Threshold[0].Action.MailTo
                Subject          = $Threshold[0].Action.Subject
                Body             = $Threshold[0].Action.Body
                RunLimitInterval = $Threshold[0].Action.RunLimitInterval
            }

            $Q2Params = @{
                Type             = 'Email'
                MailTo           = $Threshold[1].Action.MailTo
                Subject          = $Threshold[1].Action.Subject
                Body             = $Threshold[1].Action.Body
                RunLimitInterval = $Threshold[1].Action.RunLimitInterval
            }

            $New = @{
                Path        = $UserFolder1
                Template    = $SourceTemplateName
                Description = 'BEL ATT Quota home 5GB #PowerShellManaged'
                Size        = 5GB
                SoftLimit   = $false
                Threshold   = @(
                    (New-FsrmQuotaThreshold -Percentage $Threshold[0].Percentage -Action (New-FsrmAction @Q1Params)),
                    (New-FsrmQuotaThreshold -Percentage $Threshold[1].Percentage -Action (New-FsrmAction @Q2Params))
                )
            }
            Remove-FsrmQuota -Path $New.Path -Confirm:$false
            New-FsrmQuota @New

            $JSONtestDataAD = @"
[
    {
        "DistinguishedName":  "CN=Schuermans\\, Natascha (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net",
        "Enabled":  true,
        "GivenName":  "Natascha",
        "HomeDirectory":  "\\\\grouphc.net\\BNL\\HOME\\Centralized\\nschuerm",
        "Name":  "Schuermans, Natascha (Braine L’Alleud) BEL",
        "ObjectClass":  "user",
        "ObjectGUID":  "5ef73093-4907-40b0-a21f-c98e191fd545",
        "SamAccountName":  "nschuerm",
        "SID":  {
                    "BinaryLength":  28,
                    "AccountDomainSid":  "S-1-5-21-1078081533-261478967-839522115",
                    "Value":  "S-1-5-21-1078081533-261478967-839522115-680611"
                },
        "Surname":  "Schuermans",
        "UserPrincipalName":  "nschuerm@grouphc.net",
        "PropertyNames":  [
                              "DistinguishedName",
                              "Enabled",
                              "GivenName",
                              "HomeDirectory",
                              "Name",
                              "ObjectClass",
                              "ObjectGUID",
                              "SamAccountName",
                              "SID",
                              "Surname",
                              "UserPrincipalName"
                          ],
        "AddedProperties":  [

                            ],
        "RemovedProperties":  [

                              ],
        "ModifiedProperties":  [

                               ],
        "PropertyCount":  11,
        "GroupName":  "BEL ATT Quota home 120GB",
        "Limit":  "120GB",
        "Size":  128849018880,
        "Description":  "BEL ATT Quota home 120GB #PowerShellManaged",
        "SoftLimit":  false,
        "Threshold":  null,
        "ComputerName":  "DEUSFFRAN0008.grouphc.net",
        "ComputerPath":  "E:\\HOME\\nschuerm"
    },
    {
        "DistinguishedName":  "CN=Gijbels\\, Brecht (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=grouphc,DC=net",
        "Enabled":  true,
        "GivenName":  "Brecht",
        "HomeDirectory":  "\\\\grouphc.net\\BNL\\HOME\\Centralized\\gijbelsb",
        "Name":  "Gijbels, Brecht (Braine L’Alleud) BEL",
        "ObjectClass":  "user",
        "ObjectGUID":  "c1a3e1e7-cfc8-4569-885c-4d6bd2f6c433",
        "SamAccountName":  "gijbelsb",
        "SID":  {
                    "BinaryLength":  28,
                    "AccountDomainSid":  "S-1-5-21-1078081533-261478967-839522115",
                    "Value":  "S-1-5-21-1078081533-261478967-839522115-138487"
                },
        "Surname":  "Gijbels",
        "UserPrincipalName":  "gijbelsb@grouphc.net",
        "PropertyNames":  [
                              "DistinguishedName",
                              "Enabled",
                              "GivenName",
                              "HomeDirectory",
                              "Name",
                              "ObjectClass",
                              "ObjectGUID",
                              "SamAccountName",
                              "SID",
                              "Surname",
                              "UserPrincipalName"
                          ],
        "AddedProperties":  [

                            ],
        "RemovedProperties":  [

                              ],
        "ModifiedProperties":  [

                               ],
        "PropertyCount":  11,
        "GroupName":  "BEL ATT Quota home 150GB",
        "Limit":  "150GB",
        "Size":  161061273600,
        "Description":  "BEL ATT Quota home 150GB #PowerShellManaged",
        "SoftLimit":  false,
        "Threshold":  null,
        "ComputerName":  "DEUSFFRAN0008.grouphc.net",
        "ComputerPath":  "E:\\HOME\\gijbelsb"
    }
]
"@
            $User = $JSONtestDataAD -join "`n" | ConvertFrom-Json

            $User[0].ComputerName = $env:COMPUTERNAME
            $User[0].ComputerPath = $UserFolder1

            $User[1].ComputerName = $env:COMPUTERNAME
            $User[1].ComputerPath = $UserFolder2

            #Remove-FsrmQuota -Path $UserFolder1 -Confirm:$false
            #$NewParams = @{
            #    Path = $UserFolder1
            #    Size = 5GB
            #    Description = 'BEL ATT Quota home 5GB #PowerShellManaged'
            #}
            #New-FsrmQuota @NewParams

            #Remove-FsrmQuota -Path $UserFolder2 -Confirm:$false
            #$NewParams = @{
            #    Path = $UserFolder2
            #    Size = 5GB
            #    Description = 'BEL ATT Quota home 5GB #PowerShellManaged'
            #}
            #New-FsrmQuota @NewParams

            { .$testScript -User $User } | Should -Not -Throw
        }
    } -Tag test
}