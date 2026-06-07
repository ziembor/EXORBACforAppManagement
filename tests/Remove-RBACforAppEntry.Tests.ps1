#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs so the module scope can resolve them and Pester can mock them on CI
    # (where Microsoft.Graph / ExchangeOnlineManagement are not installed). Parameters the
    # code passes must be declared so the mock can bind and filter on them.
    function global:Get-MgContext { [CmdletBinding()] param() }
    function global:Get-MgServicePrincipal { [CmdletBinding()] param([string]$Filter, [string]$ServicePrincipalId) }
    function global:Get-UnifiedGroup { [CmdletBinding()] param([string]$Identity) }
    function global:Get-ManagementRoleAssignment { [CmdletBinding()] param([string]$Role, [string]$Identity) }
    function global:Get-UnifiedGroupLinks { [CmdletBinding()] param([string]$Identity, [string]$LinkType) }
    function global:Remove-ManagementRoleAssignment { [CmdletBinding(SupportsShouldProcess)] param([string]$Identity) }
    function global:Remove-UnifiedGroup { [CmdletBinding(SupportsShouldProcess)] param([string]$Identity) }

    $script:Sp = [pscustomobject]@{ DisplayName = 'Contoso'; AppId = '11111111-1111-1111-1111-111111111111'; Id = '22222222-2222-2222-2222-222222222222' }
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgContext','Get-MgServicePrincipal','Get-UnifiedGroup','Get-ManagementRoleAssignment','Get-UnifiedGroupLinks','Remove-ManagementRoleAssignment','Remove-UnifiedGroup') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'Remove-RBACforAppEntry' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'admin@contoso.com' } }
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { $script:Sp }
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { [pscustomobject]@{ DisplayName = 'Um365RAo1-Contoso'; Identity = 'Um365RAo1-Contoso' } }
        # One own assignment scoped to the group.
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment {
            @([pscustomobject]@{
                Name                      = 'AppMailSend-Contoso'
                Role                      = 'Application Mail.Send'
                RoleAssigneeName          = 'Contoso_SP'
                RecipientWriteScope       = 'CustomRecipientScope'
                CustomRecipientWriteScope = 'Um365RAo1-Contoso'
            })
        }
        # Only the bootstrap placeholder member.
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroupLinks { @([pscustomobject]@{ PrimarySmtpAddress = $null; Name = 'GraphAPI-Dummy' }) }
        Mock -ModuleName EXORBACforAppManagement Remove-ManagementRoleAssignment { }
        Mock -ModuleName EXORBACforAppManagement Remove-UnifiedGroup { }
    }

    It 'removes own assignments and the group when clean' {
        $r = Remove-RBACforAppEntry -RegisteredAppName 'Contoso' -Confirm:$false

        $r.IsRemoved | Should -BeTrue
        $r.UnifiedGroupName | Should -Be 'Um365RAo1-Contoso'
        $r.OwnAssignments | Should -Be @('AppMailSend-Contoso')
        $r.ForeignAssignments | Should -BeNullOrEmpty
        $r.AssignmentsRemoved | Should -Be @('AppMailSend-Contoso')
        $r.GroupRemoved | Should -BeTrue
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-ManagementRoleAssignment -Times 1
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-UnifiedGroup -Times 1
    }

    It 'aborts when a foreign assignment is scoped to the group' {
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment {
            @(
                [pscustomobject]@{ Name = 'AppMailSend-Contoso'; Role = 'Application Mail.Send'; RoleAssigneeName = 'Contoso_SP'; RecipientWriteScope = 'CustomRecipientScope'; CustomRecipientWriteScope = 'Um365RAo1-Contoso' },
                [pscustomobject]@{ Name = 'AppMailRead-Other'; Role = 'Application Mail.Read'; RoleAssigneeName = 'OtherApp_SP'; RecipientWriteScope = 'CustomRecipientScope'; CustomRecipientWriteScope = 'Um365RAo1-Contoso' }
            )
        }

        $r = Remove-RBACforAppEntry -RegisteredAppName 'Contoso' -Confirm:$false

        $r.IsRemoved | Should -BeFalse
        $r.Reason | Should -Not -BeNullOrEmpty
        $r.ForeignAssignments | Should -Be @('AppMailRead-Other')
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-ManagementRoleAssignment -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-UnifiedGroup -Times 0
    }

    It 'aborts when the group has a real member beyond the bootstrap dummy' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroupLinks {
            @(
                [pscustomobject]@{ PrimarySmtpAddress = $null; Name = 'GraphAPI-Dummy' },
                [pscustomobject]@{ PrimarySmtpAddress = 'real@contoso.com'; Name = 'real' }
            )
        }

        $r = Remove-RBACforAppEntry -RegisteredAppName 'Contoso' -Confirm:$false

        $r.IsRemoved | Should -BeFalse
        $r.RealMembers | Should -Be @('real@contoso.com')
        $r.Reason | Should -Not -BeNullOrEmpty
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-UnifiedGroup -Times 0
    }

    It 'treats a group with only the bootstrap member as clean' {
        $r = Remove-RBACforAppEntry -RegisteredAppName 'Contoso' -Confirm:$false

        $r.RealMembers | Should -BeNullOrEmpty
        $r.IsRemoved | Should -BeTrue
    }

    It 'performs no removals under -WhatIf' {
        $r = Remove-RBACforAppEntry -RegisteredAppName 'Contoso' -WhatIf

        $r.IsRemoved | Should -BeFalse
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-ManagementRoleAssignment -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-UnifiedGroup -Times 0
    }

    It 'reports UnifiedGroupExisted false when the group is already gone' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { }

        $r = Remove-RBACforAppEntry -RegisteredAppName 'Contoso' -Confirm:$false

        $r.UnifiedGroupExisted | Should -BeFalse
        Should -Invoke -ModuleName EXORBACforAppManagement Remove-UnifiedGroup -Times 0
    }

    It 'records an error and is not removed when the service principal cannot be resolved' {
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { @() }

        $r = Remove-RBACforAppEntry -AppId '33333333-3333-3333-3333-333333333333' -Confirm:$false

        $r.IsRemoved | Should -BeFalse
        $r.Errors.Count | Should -BeGreaterThan 0
    }
}
