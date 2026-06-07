#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs for the external Graph/EXO cmdlets so Pester can mock them in the module
    # scope without the real Microsoft.Graph / ExchangeOnlineManagement modules installed.
    # They must be global so the module's session state (a child of global) can resolve them.
    function global:Get-MgContext { }
    function global:Get-MgServicePrincipal { }
    function global:Get-UnifiedGroup { }
    function global:New-UnifiedGroup { }
    function global:Set-UnifiedGroup { }
    function global:Add-UnifiedGroupLinks { }
    function global:New-ServicePrincipal { }
    function global:Get-Recipient { }
    function global:New-ManagementRoleAssignment { }

    $script:Sp = [pscustomobject]@{
        DisplayName = 'Contoso'
        AppId       = '11111111-1111-1111-1111-111111111111'
        Id          = '22222222-2222-2222-2222-222222222222'
    }
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgContext','Get-MgServicePrincipal','Get-UnifiedGroup','New-UnifiedGroup','Set-UnifiedGroup','Add-UnifiedGroupLinks','New-ServicePrincipal','Get-Recipient','New-ManagementRoleAssignment') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'New-RBACforAppEntry SP resolution' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'admin@contoso.com' } }
    }

    It 'records an error when no SP matches the AppId' {
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { @() }

        $r = New-RBACforAppEntry -AppId '33333333-3333-3333-3333-333333333333' -WhatIf
        $r.Errors -join ';' | Should -Match 'No service principal found'
    }

    It 'records an error when the display name is ambiguous' {
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { @($script:Sp, $script:Sp) }

        $r = New-RBACforAppEntry -RegisteredAppName 'dup' -WhatIf
        $r.Errors -join ';' | Should -Match 'Ambiguous'
    }
}

Describe 'New-RBACforAppEntry -WhatIf' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'admin@contoso.com' } }
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal { $script:Sp }
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { [pscustomobject]@{ DisplayName = 'g'; Identity = 'g'; ManagedBy = @() } }
        Mock -ModuleName EXORBACforAppManagement Get-Recipient { [pscustomobject]@{ PrimarySmtpAddress = 'shared@contoso.com' } }
        Mock -ModuleName EXORBACforAppManagement New-UnifiedGroup { }
        Mock -ModuleName EXORBACforAppManagement Add-UnifiedGroupLinks { }
        Mock -ModuleName EXORBACforAppManagement New-ServicePrincipal { }
        Mock -ModuleName EXORBACforAppManagement New-ManagementRoleAssignment { }
        Mock -ModuleName EXORBACforAppManagement Export-Clixml { }
    }

    It 'normalizes the role and builds the assignment name' {
        $r = New-RBACforAppEntry -RegisteredAppName 'Contoso' -Members 'shared@contoso.com' -Role 'Mail.Send' -WhatIf
        $r.RolesNormalized | Should -Contain 'Application Mail.Send'
        $r.RoleAssignmentsName[0] | Should -BeLike 'AppMailSend-*'
    }

    It 'reports the requested and added Unified Group owner' {
        $r = New-RBACforAppEntry -RegisteredAppName 'Contoso' -ManagedBy 'owner@contoso.com' -Role 'Mail.Send' -WhatIf
        $r.PSObject.Properties.Name | Should -Contain 'OwnerRequested'
        $r.PSObject.Properties.Name | Should -Contain 'OwnerAdded'
        $r.OwnerRequested | Should -Be 'owner@contoso.com'
    }

    It 'does not perform any mutating EXO calls under -WhatIf' {
        $null = New-RBACforAppEntry -RegisteredAppName 'Contoso' -Role 'Mail.Send' -WhatIf
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ManagementRoleAssignment -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-UnifiedGroup -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Add-UnifiedGroupLinks -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ServicePrincipal -Times 0
    }
}
