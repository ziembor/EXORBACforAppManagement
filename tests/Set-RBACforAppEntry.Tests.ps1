#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs for the external Graph/EXO cmdlets so Pester can mock them in the module
    # scope without the real Microsoft.Graph / ExchangeOnlineManagement modules installed.
    # They must be global so the module's session state (a child of global) can resolve them.
    # The delegated module functions (New-RBACforAppUnifiedGroup / Register-EXOServicePrincipal)
    # are mocked directly with -ModuleName so their internals are not exercised here.
    function global:Get-MgContext { }
    function global:Get-MgServicePrincipal { param([string]$Filter, [string]$ServicePrincipalId) }
    function global:Get-UnifiedGroup { param([string]$Identity) }
    function global:Get-UnifiedGroupLinks { param([string]$Identity, [string]$LinkType) }
    function global:Add-UnifiedGroupLinks { param([string]$Identity, [string]$LinkType, [string]$Links) }
    function global:Get-ServicePrincipal { param([string]$Identity) }
    function global:Get-Recipient { param([string]$Identity) }
    function global:Get-ManagementRoleAssignment { param([string]$Identity, [string]$Role) }
    function global:New-ManagementRoleAssignment { param($App, $Role, $RecipientGroupScope, $Name) }
    function global:Remove-ManagementRoleAssignment { param([string]$Identity) }

    $script:Sp = [pscustomobject]@{
        DisplayName = 'Contoso'
        AppId       = '11111111-1111-1111-1111-111111111111'
        Id          = '22222222-2222-2222-2222-222222222222'
    }
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgContext','Get-MgServicePrincipal','Get-UnifiedGroup','Get-UnifiedGroupLinks','Add-UnifiedGroupLinks','Get-ServicePrincipal','Get-Recipient','Get-ManagementRoleAssignment','New-ManagementRoleAssignment','Remove-ManagementRoleAssignment') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'Set-RBACforAppEntry SP resolution' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'admin@contoso.com' } }
    }

    It 'records an error when no SP matches the AppId' {
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { @() }

        $r = Set-RBACforAppEntry -AppId '33333333-3333-3333-3333-333333333333' -WhatIf
        $r.Errors -join ';' | Should -Match 'No service principal found'
    }

    It 'records an error when the display name is ambiguous' {
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { @($script:Sp, $script:Sp) }

        $r = Set-RBACforAppEntry -RegisteredAppName 'dup' -WhatIf
        $r.Errors -join ';' | Should -Match 'Ambiguous'
    }
}

Describe 'Set-RBACforAppEntry reconcile' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'admin@contoso.com' } }
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { $script:Sp }
        # Everything present by default: group exists, EXO SP exists, assignment exists scoped to current group.
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { [pscustomobject]@{ DisplayName = 'g'; Identity = $Identity } }
        Mock -ModuleName EXORBACforAppManagement Get-ServicePrincipal { @([pscustomobject]@{ DisplayName = 'Contoso_SP'; AppId = '11111111-1111-1111-1111-111111111111' }) }
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroupLinks { @() }
        Mock -ModuleName EXORBACforAppManagement Get-Recipient { [pscustomobject]@{ PrimarySmtpAddress = $Identity; Name = $Identity } }
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment {
            [pscustomobject]@{
                Name                      = $Identity
                Role                      = 'Application Mail.Send'
                RecipientWriteScope       = 'Group'
                CustomRecipientWriteScope = 'Um365RAo1-Contoso'
            }
        }
        # Delegated module functions + mutating cmdlets.
        Mock -ModuleName EXORBACforAppManagement New-RBACforAppUnifiedGroup { [pscustomobject]@{ OwnerRequested = 'o'; OwnerAdded = 'o'; AlreadyExisted = $false; Group = [pscustomobject]@{ Identity = $Name } } }
        Mock -ModuleName EXORBACforAppManagement Register-EXOServicePrincipal { [pscustomobject]@{ DisplayName = $DisplayName } }
        Mock -ModuleName EXORBACforAppManagement Add-UnifiedGroupLinks { }
        Mock -ModuleName EXORBACforAppManagement New-ManagementRoleAssignment { }
        Mock -ModuleName EXORBACforAppManagement Remove-ManagementRoleAssignment { }
    }

    It 'leaves a fully-configured app on its group untouched and reports IsValid' {
        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send'

        $r.IsValid | Should -BeTrue
        $r.RoleAssignmentsUnchanged | Should -Contain 'AppMailSend-Contoso'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppUnifiedGroup -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Register-EXOServicePrincipal -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ManagementRoleAssignment -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Remove-ManagementRoleAssignment -Times 0
    }

    It 'creates only the missing role assignment, not the existing group' {
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment { }

        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -Confirm:$false

        $r.RoleAssignmentsCreated | Should -Contain 'AppMailSend-Contoso'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ManagementRoleAssignment -Times 1
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppUnifiedGroup -Times 0
    }

    It 'creates the Unified Group when it is missing' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { }

        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -Confirm:$false

        $r.UnifiedGroupCreated | Should -BeTrue
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppUnifiedGroup -Times 1
    }

    It 'creates the Exchange Online service principal when it is missing' {
        Mock -ModuleName EXORBACforAppManagement Get-ServicePrincipal { @() }

        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -Confirm:$false

        $r.ExoServicePrincipalCreated | Should -BeTrue
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Register-EXOServicePrincipal -Times 1
    }

    It 'adds a requested member that is not already in the group (additive)' {
        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -Members 'new@contoso.com' -Confirm:$false

        $r.MembersAdded | Should -Contain 'new@contoso.com'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Add-UnifiedGroupLinks -Times 1
    }

    It 'does not re-add a member already present in the group' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroupLinks { @([pscustomobject]@{ PrimarySmtpAddress = 'shared@contoso.com'; Name = 'shared' }) }

        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -Members 'shared@contoso.com'

        $r.MembersAlreadyPresent | Should -Contain 'shared@contoso.com'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Add-UnifiedGroupLinks -Times 0
    }

    It 're-scopes the assignment onto a new group when -NewGroupPrefix is supplied' {
        $r = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -NewGroupPrefix 'Um365Prod' -Confirm:$false

        $r.GroupChanged | Should -BeTrue
        $r.TargetGroupName | Should -Be 'Um365Prod-Contoso'
        $r.RoleAssignmentsRescoped | Should -Contain 'AppMailSend-Contoso'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Remove-ManagementRoleAssignment -Times 1
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ManagementRoleAssignment -Times 1
    }

    It 'makes no mutating calls under -WhatIf' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { }
        Mock -ModuleName EXORBACforAppManagement Get-ServicePrincipal { @() }
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment { }

        $null = Set-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -Members 'new@contoso.com' -WhatIf

        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-RBACforAppUnifiedGroup -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Register-EXOServicePrincipal -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Add-UnifiedGroupLinks -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ManagementRoleAssignment -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Remove-ManagementRoleAssignment -Times 0
    }
}
