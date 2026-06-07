#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs so the module scope can resolve them and Pester can mock them on CI
    # (where Microsoft.Graph is not installed).
    function global:Get-MgContext { }
    function global:New-MgApplication { }
    function global:New-MgServicePrincipal { }
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgContext','New-MgApplication','New-MgServicePrincipal') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'New-RegisteredApp' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-1' } }
        Mock -ModuleName EXORBACforAppManagement New-MgApplication { [pscustomobject]@{ AppId = 'aaaa1111-1111-1111-1111-111111111111'; Id = 'bbbb2222-2222-2222-2222-222222222222' } }
        Mock -ModuleName EXORBACforAppManagement New-MgServicePrincipal { [pscustomobject]@{ Id = 'cccc3333-3333-3333-3333-333333333333' } }
    }

    It 'creates the application and service principal and returns their ids' {
        $r = New-RegisteredApp -DisplayName 'Contoso Mail App' -Confirm:$false
        $r.AppId | Should -Be 'aaaa1111-1111-1111-1111-111111111111'
        $r.AppObjectId | Should -Be 'bbbb2222-2222-2222-2222-222222222222'
        $r.ServicePrincipalId | Should -Be 'cccc3333-3333-3333-3333-333333333333'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-MgServicePrincipal -Times 1
    }

    It 'skips the service principal with -SkipServicePrincipal' {
        $r = New-RegisteredApp -DisplayName 'Contoso Mail App' -SkipServicePrincipal -Confirm:$false
        $r.ServicePrincipalId | Should -BeNullOrEmpty
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-MgServicePrincipal -Times 0
    }

    It 'creates nothing under -WhatIf' {
        $r = New-RegisteredApp -DisplayName 'Contoso Mail App' -WhatIf
        $r.AppId | Should -BeNullOrEmpty
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-MgApplication -Times 0
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-MgServicePrincipal -Times 0
    }
}
