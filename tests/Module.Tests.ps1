#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $script:ManifestPath = Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1'
    Import-Module $script:ManifestPath -Force
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
}

Describe 'EXORBACforAppManagement module' {
    It 'has a valid manifest' {
        { Test-ModuleManifest -Path $script:ManifestPath -ErrorAction Stop } | Should -Not -Throw
    }

    It 'exports exactly the ten public functions' {
        $exported = (Get-Command -Module EXORBACforAppManagement).Name | Sort-Object
        $exported | Should -Be @('Convert-ApplicationAccessPolicyToRBAC', 'Get-RBACforAppEntry', 'Get-RegisteredAppWithPermission', 'New-RBACforAppEntry', 'New-RBACforAppUnifiedGroup', 'New-RegisteredApp', 'Register-EXOServicePrincipal', 'Remove-RBACforAppEntry', 'Set-RBACforAppEntry', 'Test-RBACforAppEntry')
    }

    It 'does not export the private helpers' {
        foreach ($helper in 'Get-SafeName', 'Get-NormalizeRole', 'ConvertTo-AppRole', 'Get-AppRoleMap', 'Get-LegacyScopeRoleMap', 'Resolve-AppRolePermissionValue') {
            (Get-Command -Module EXORBACforAppManagement -Name $helper -ErrorAction SilentlyContinue) | Should -BeNullOrEmpty
        }
    }

    It 'makes the private helpers available inside the module scope' {
        InModuleScope EXORBACforAppManagement {
            (Get-Command Get-SafeName, Get-NormalizeRole, ConvertTo-AppRole, Get-AppRoleMap, Get-LegacyScopeRoleMap, Resolve-AppRolePermissionValue).Count | Should -Be 6
        }
    }
}
