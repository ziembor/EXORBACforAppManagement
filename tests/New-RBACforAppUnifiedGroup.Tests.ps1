#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs so the module scope can resolve them and Pester can mock them on CI.
    function global:Get-MgContext { }
    function global:Get-UnifiedGroup { }
    function global:New-UnifiedGroup { }
    function global:Set-UnifiedGroup { }
    function global:Get-Recipient { }
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgContext','Get-UnifiedGroup','New-UnifiedGroup','Set-UnifiedGroup','Get-Recipient') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'New-RBACforAppUnifiedGroup' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1'; Account = 'admin@contoso.com' } }
        Mock -ModuleName EXORBACforAppManagement Set-UnifiedGroup { }
        Mock -ModuleName EXORBACforAppManagement Get-Recipient { [pscustomobject]@{ PrimarySmtpAddress = 'owner@contoso.com' } }
    }

    It 'creates and configures the group when it does not exist' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { } # not found, then configured lookup
        Mock -ModuleName EXORBACforAppManagement New-UnifiedGroup { [pscustomobject]@{ DisplayName = 'g'; Alias = 'g'; AccessType = 'Private' } }

        $res = New-RBACforAppUnifiedGroup -Name 'Um365RAo1-Contoso' -ManagedBy 'owner@contoso.com' -Confirm:$false
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-UnifiedGroup -Times 1
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Set-UnifiedGroup -Times 1
        $res.AlreadyExisted | Should -BeFalse
        $res.OwnerRequested | Should -Be 'owner@contoso.com'
        $res.OwnerAdded     | Should -Be 'owner@contoso.com'
    }

    It 'warns and does not create when the group already exists' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { [pscustomobject]@{ DisplayName = 'g'; Identity = 'g'; ManagedBy = @('existing@contoso.com') } }
        Mock -ModuleName EXORBACforAppManagement New-UnifiedGroup { }

        $warn = $null
        $res = New-RBACforAppUnifiedGroup -Name 'Um365RAo1-Contoso' -ManagedBy 'owner@contoso.com' -Confirm:$false -WarningVariable warn
        ($warn.Message -join ';') | Should -Match 'already exists'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-UnifiedGroup -Times 0
        $res.AlreadyExisted | Should -BeTrue
        $res.OwnerRequested | Should -Be 'owner@contoso.com'
        $res.OwnerAdded     | Should -Be 'existing@contoso.com'
    }

    It 'does not create under -WhatIf' {
        Mock -ModuleName EXORBACforAppManagement Get-UnifiedGroup { }
        Mock -ModuleName EXORBACforAppManagement New-UnifiedGroup { }

        $null = New-RBACforAppUnifiedGroup -Name 'Um365RAo1-Contoso' -WhatIf
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-UnifiedGroup -Times 0
    }
}
