#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs for the external cmdlets the function calls, declaring the parameters
    # it filters on so Mock -ModuleName can bind and intercept them.
    function global:Get-ApplicationAccessPolicy { [CmdletBinding()] param([string]$AppId) }
    function global:Get-MgServicePrincipal { [CmdletBinding()] param([string]$Filter, [string]$ServicePrincipalId) }
    function global:Get-MgServicePrincipalAppRoleAssignment { [CmdletBinding()] param([string]$ServicePrincipalId) }
    function global:Get-DistributionGroupMember { [CmdletBinding()] param([string]$Identity) }
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-ApplicationAccessPolicy','Get-MgServicePrincipal','Get-MgServicePrincipalAppRoleAssignment','Get-DistributionGroupMember') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'Convert-ApplicationAccessPolicyToRBAC' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-ApplicationAccessPolicy {
            [pscustomobject]@{
                AppId       = '11111111-1111-1111-1111-111111111111'
                AccessRight = 'RestrictAccess'
                ScopeName   = 'EvenUsers'
                Description = 'Restrict app to EvenUsers'
            }
        }

        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal {
            # Pester only defines the parameter variable when it was bound on the call,
            # so Test-Path is the StrictMode-safe way to tell the two call shapes apart.
            if (Test-Path variable:ServicePrincipalId) {
                # Resource service principal (e.g. Microsoft Graph) AppRoles catalog.
                return [pscustomobject]@{
                    Id       = $ServicePrincipalId
                    AppRoles = @(
                        [pscustomobject]@{ Id = 'role-mailread'; Value = 'Mail.Read' }
                        [pscustomobject]@{ Id = 'role-ews';      Value = 'full_access_as_app' }
                        [pscustomobject]@{ Id = 'role-dir';      Value = 'Directory.Read.All' }
                    )
                }
            }
            # App service principal resolution by appId filter.
            [pscustomobject]@{
                DisplayName = 'Contoso'
                AppId       = '11111111-1111-1111-1111-111111111111'
                Id          = '22222222-2222-2222-2222-222222222222'
            }
        }

        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipalAppRoleAssignment {
            @(
                [pscustomobject]@{ AppRoleId = 'role-mailread'; ResourceId = '99999999-9999-9999-9999-999999999999' }
                [pscustomobject]@{ AppRoleId = 'role-ews';      ResourceId = '99999999-9999-9999-9999-999999999999' }
                [pscustomobject]@{ AppRoleId = 'role-dir';      ResourceId = '99999999-9999-9999-9999-999999999999' }
            )
        }

        Mock -ModuleName EXORBACforAppManagement Get-DistributionGroupMember {
            @(
                [pscustomobject]@{ PrimarySmtpAddress = 'user1@contoso.com' }
                [pscustomobject]@{ PrimarySmtpAddress = 'user2@contoso.com' }
            )
        }

        Mock -ModuleName EXORBACforAppManagement New-RBACforAppEntry {
            [pscustomobject]@{ Warnings = @(); Errors = @() }
        }
    }

    It 'derives RBAC roles from the app grants and delegates to New-RBACforAppEntry' {
        $r = Convert-ApplicationAccessPolicyToRBAC -Confirm:$false

        $r.AppId | Should -Be '11111111-1111-1111-1111-111111111111'
        $r.ResolvedDisplay | Should -Be 'Contoso'
        $r.DerivedRoles | Should -Contain 'Application Mail.Read'
        $r.MembersCopied | Should -Be @('user1@contoso.com', 'user2@contoso.com')

        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppEntry -Times 1 -ParameterFilter {
            ($Role -contains 'Application Mail.Read') -and ($Members -contains 'user1@contoso.com')
        }
    }

    It 'maps the legacy EWS scope full_access_as_app to Application EWS.AccessAsApp' {
        $r = Convert-ApplicationAccessPolicyToRBAC -Confirm:$false
        $r.DerivedRoles | Should -Contain 'Application EWS.AccessAsApp'
    }

    It 'records grants with no Application Access Policy equivalent as skipped' {
        $r = Convert-ApplicationAccessPolicyToRBAC -Confirm:$false
        $r.SkippedPermissions | Should -Contain 'Directory.Read.All'
        $r.DerivedRoles | Should -Not -Contain 'Application Directory.Read.All'
    }

    It 'skips DenyAccess policies with a warning and does not assign roles' {
        Mock -ModuleName EXORBACforAppManagement Get-ApplicationAccessPolicy {
            [pscustomobject]@{ AppId = '11111111-1111-1111-1111-111111111111'; AccessRight = 'DenyAccess'; ScopeName = 'EvenUsers' }
        }

        $r = Convert-ApplicationAccessPolicyToRBAC -WarningAction SilentlyContinue

        $r.Warnings | Should -Not -BeNullOrEmpty
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppEntry -Times 0
    }

    It 'uses an explicit -Role override instead of deriving from grants' {
        $r = Convert-ApplicationAccessPolicyToRBAC -Role 'Mail.Send' -Confirm:$false

        $r.DerivedRoles | Should -Be @('Application Mail.Send')
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Get-MgServicePrincipalAppRoleAssignment -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppEntry -Times 1 -ParameterFilter {
            $Role -contains 'Application Mail.Send'
        }
    }

    It 'does not delegate when -WhatIf is supplied' {
        $null = Convert-ApplicationAccessPolicyToRBAC -WhatIf
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppEntry -Times 0
    }
}

Describe 'Get-LegacyScopeRoleMap' {
    It 'maps every legacy scope to a valid Get-AppRoleMap role' {
        InModuleScope EXORBACforAppManagement {
            $appRoles = Get-AppRoleMap
            $legacy   = Get-LegacyScopeRoleMap
            foreach ($value in $legacy.Values) {
                $appRoles.Contains($value) | Should -BeTrue -Because "$value should be a key in Get-AppRoleMap"
            }
        }
    }
}
