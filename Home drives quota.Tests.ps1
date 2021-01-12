#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Import-Module ActiveDirectory -Verbose:$false -Force
    $ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'

    $testADgroupNamePrefix = 'Group H drive quota'
    $testThresholdData = @(
        @{
            Percentage = 85
            Color      = 'Red'
            Action     = @{
                MailTo           = 'Brecht.Goijbels@heidelbergcement.com'
                Subject          = "Your personal home drive has reached [Quota Threshold]% of its maximum allowed volume"
                Body             = @"
Dear [Source Io Owner],

Your personal home drive has reached [Quota Threshold]% of its maximum allowed volume.

The quota limit is [Quota Limit MB] MB, and [Quota Used MB] MB currently is in use ([Quota Used Percent]% of limit). Please take into account that once your quota has been reached, you will no longer be able to edit or save files on your home drive.

In case you need assistance in handling this issue, we kindly ask you to contact us on bnl.servicedesk@heidelbergcement.com

Kind regards,
IT Service Desk
"@
                Type             = 2
                RunLimitInterval = 2880
            }
        }
        @{
            Percentage = 100
            Color      = 'Green'
            Action     = @{
                MailTo           = 'Brecht.Goijbels@heidelbergcement.com'
                Subject          = 'Your personal home drive has reached or exceeded its maximum allowed volume'
                Body             = @"
Dear [Source Io Owner],

Your personal home drive has now reached [Quota Threshold]% of its maximum allowed volume.
Please take into account that you will no longer be able to edit or save files on your home drive. You can either take the necessary measures to free up space on your home drive, or submit a formal request for a higher volume limit.

In case you need assistance in handling this issue, we kindly ask you to contact us on bnl.servicedesk@heidelbergcement.com

Kind regards,
IT Service Desk
"@
                Type             = 2
                RunLimitInterval = 2880
            }
        }
    )

    $testQuotaRemoveGroup = New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
        SamAccountName = '{0} REMOVE' -f $testADgroupNamePrefix
        Description    = 'Home drives remove quota limit'
        CanonicalName  = 'contoso.com/EU/BEL/Groups/{0} REMOVE' -f $testADgroupNamePrefix
        GroupCategory  = 'Security'
        GroupScope     = 'Universal'
    }

    $testQuotaGroups = @(
        New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
            SamAccountName = "$($testADgroupNamePrefix) 20GB"
            Description    = 'Home drives 20GB'
            CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f "$($testADgroupNamePrefix) 20GB"
            GroupCategory  = 'Security'
            GroupScope     = 'Universal'
        }
    )

    $testUsers = @(
        New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
            SamAccountName = 'cnorris'
            GivenName      = 'Chcuk'
            Surname        = 'Norris'
            Enabled        = $true
            Description    = 'Normal users'
            HomeDirectory  = '\\contoso\home\cnorris'
            HomeDrive      = 'H:'
            CanonicalName  = 'contoso.com/EU/BEL/Users/{0}' -f 'cnorris'
        }
        New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
            SamAccountName = 'bswagger'
            GivenName      = 'Bob Lee'
            Surname        = 'NSwagger'
            Enabled        = $true
            Description    = 'Normal users'
            HomeDirectory  = '\\contoso\home\bswagger'
            HomeDrive      = 'H:'
            CanonicalName  = 'contoso.com/EU/BEL/Users/{0}' -f 'bswagger'
        }
    )

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ADGroupName        = $testADgroupNamePrefix
        ADGroupRemoveName  = $testQuotaRemoveGroup.SamAccountName
        ScriptName         = 'Test'
        LogFolder          = (New-Item 'TestDrive:\Log' -ItemType Directory).FullName
        ThresholdFile      = (New-Item 'TestDrive:\Thresholds.json' -ItemType File).FullName
        SetQuotaScriptFile = (New-Item 'TestDrive:\Script.ps1' -ItemType File).FullName
        MailTo             = $ScriptAdmin
    }

    $testThresholdData | ConvertTo-Json | Out-File $testParams.ThresholdFile -Force

    Mock Export-Excel
    Mock Get-ADGroup
    Mock Get-ADGroupMember
    Mock Get-ADUser
    Mock Get-DFSDetailsHC {
        @{
            Path         = $testUsers[0].HomeDirectory
            ComputerName = $env:COMPUTERNAME
            ComputerPath = $null
        }
        @{
            Path         = $testUsers[1].HomeDirectory
            ComputerName = $env:COMPUTERNAME
            ComputerPath = $null
        }
    }
    Mock Invoke-Command
    Mock Send-MailHC
    Mock Write-EventLog
    Mock Set-Culture
}

Describe 'send an error mail to tha admin when' {
    It 'an ADGroupName has an incorrect quota limit at the end of its name' {
        Mock Get-ADGroup {
            $testQuotaRemoveGroup

            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = "$($testADgroupNamePrefix) Wrong"
                Description    = 'Home drives 20GB'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f "$($testADgroupNamePrefix) 20GB"
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }
        }

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Failed converting the quota limit string*KB size*")
        }
        Should -Invoke Invoke-Command -Exactly 0
        Should -Invoke Get-ADUser -Exactly 0
        $Users | Should -BeNullOrEmpty
    } 
    Context 'a mandatory parameter is missing' {
        It "<_>" -Foreach @(
            'ScriptName' ,
            'ADGroupName',
            'MailTo' ,
            'ADGroupRemoveName' 
        ) {
            (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | Should -BeTrue
        }
    }
} -Tag test
Context 'not found' {
    It 'LogFolder' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.LogFolder = 'NotExisting'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Path*not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    It 'SetQuotaScriptFile' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.SetQuotaScriptFile = 'NotExisting.ps1'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Quota script file*not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    It 'ADGroupName' {
        Mock Get-ADGroup

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Couldn't find active directory groups starting with*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    It 'ADGroupRemoveName' {
        Mock Get-ADGroup {
            $testQuotaGroups[0]
        }

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Couldn't find the active directory group*for removing quotas*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
}
Context 'ThresholdFile' {
    AfterAll {
        $testThresholdData | ConvertTo-Json | Out-File $testParams.ThresholdFile
    }
    It 'not found' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.ThresholdFile = 'NotExisting'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Threshold file*not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    It 'invalid JSON data in ThresholdFile' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.ThresholdFile = './testThresholdCorrupt.json'
        'Incorrect data' | Out-File $testNewParams.ThresholdFile -Force
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Threshold file*does not contain valid (JSON) data*")
        }
        Should -Invoke Invoke-Command -Exactly 0
        Should -Invoke Get-ADUser -Exactly 0
        $Users | Should -BeNullOrEmpty
    } 
    It 'Percentage' {
        $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
        $testNewThreshold.Remove('Percentage')
        $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Threshold*Percentage*not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    It 'Color' {
        $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
        $testNewThreshold.Remove('Color')
        $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Color*not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    It 'Action' {
        $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
        $testNewThreshold.Remove('Action')
        $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*Action*not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
        Should -Invoke Invoke-Command -Exactly 0
    } 
    Context 'Action' {
        It 'MailTo' {
            $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
            $testNewThreshold.Action.Remove('MailTo')
            $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*MailTo*not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
            Should -Invoke Invoke-Command -Exactly 0
        } 
        It 'Subject' {
            $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
            $testNewThreshold.Action.Remove('Subject')
            $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Subject*not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
            Should -Invoke Invoke-Command -Exactly 0
        } 
        It 'Body' {
            $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
            $testNewThreshold.Action.Remove('Body')
            $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Body*not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
            Should -Invoke Invoke-Command -Exactly 0
        } 
        It 'Type' {
            $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
            $testNewThreshold.Action.Remove('Type')
            $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Type*not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
            Should -Invoke Invoke-Command -Exactly 0
        } 
        It 'RunLimitInterval' {
            $testNewThreshold = Copy-ObjectHC $testThresholdData[0]
            $testNewThreshold.Action.Remove('RunLimitInterval')
            $testNewThreshold | ConvertTo-Json | Out-File $testParams.ThresholdFile

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*RunLimitInterval*not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
            Should -Invoke Invoke-Command -Exactly 0
        } 
    }
}

Describe 'register an error in the Error worksheet when' {
    It 'the same user is member of multiple quota limit groups' {
        $WriteErrorCommand = Get-Command -Name Write-Error

        Mock Get-ADGroup {
            $testQuotaGroups[0]
            $testQuotaRemoveGroup
        }
        Mock Get-ADGroupMember {
            $testUsers[1]
        } -ParameterFilter {
            $testQuotaGroups[0].SamAccountName -eq $Identity
        }
        Mock Get-ADGroupMember {
            $testUsers[1]
        } -ParameterFilter {
            $testQuotaRemoveGroup.SamAccountName -eq $Identity
        }
        Mock Get-ADUser {
            $testUsers[1]
        } -ParameterFilter {
            $testUsers[1].SamAccountName -eq $Identity
        }
        Mock Write-Error -MockWith {
            & $WriteErrorCommand -Message 'User member of multiple'
        } -ParameterFilter {
            $Message -like '*member of multiple groups*'
        }

        .$testScript @testParams -EA SilentlyContinue

        Should -Invoke Write-Error -Exactly 1
        Should -Invoke Export-Excel -Exactly 1 -ParameterFilter {
            $WorksheetName -eq 'Errors'
        }
    } 
}

Context 'when the quota limit groups have no members' {
    BeforeAll {
        Mock Get-ADGroup {
            $testQuotaGroups[0]
            $testQuotaRemoveGroup
        }

        .$testScript @testParams
    }

    It 'a summary mail is sent to the user instead of an error mail' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            $MailParams.Subject -notMatch 'FAILURE'
        }
    } 
    It 'the script that applies the quotas is not called' {
        Should -Invoke Invoke-Command -Exactly 0 -Scope Context
    } 
}

Context 'when the quota limit groups have a member' {
    BeforeAll {
        Mock Get-ADGroup {
            $testQuotaGroups[0]
            $testQuotaRemoveGroup
        }
        Mock Get-ADGroupMember {
            $testUsers[1]
        } -ParameterFilter {
            $testQuotaGroups[0].SamAccountName -eq $Identity
        }
        Mock Get-ADUser {
            $testUsers[1]
        } -ParameterFilter {
            $testUsers[1].SamAccountName -eq $Identity
        }
        Mock Invoke-Command {
            [PSCustomObject]@{
                GroupName    = 'Quota 15GB'
                ComputerName = 'PC1'
            }
        } -ParameterFilter {
            $FilePath -eq (Get-Item $testParams.SetQuotaScriptFile)
        }

        .$testScript @testParams
    }

    It 'a summary mail is sent to the user instead of an error mail' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            $MailParams.Subject -notMatch 'FAILURE'
        }
    }
    It 'the script that applies the quotas is called' {
        Should -Invoke Invoke-Command -Exactly 1 -Scope Context -ParameterFilter {
            $FilePath -eq (Get-Item $testParams.SetQuotaScriptFile)
        }
    }
    It 'an Excel file is created containing a summary' {
        Should -Invoke Export-Excel -Exactly 1 -Scope Context -ParameterFilter {
            $WorksheetName -eq 'Quotas'
        }
    }
    It 'a mail is sent with the Excel file in attachment' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            $MailParams.Attachments -ne $null
        }
    }
}
