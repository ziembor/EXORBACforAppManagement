#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stub so the module scope can resolve it and Pester can mock it on CI.
    function global:New-ServicePrincipal { [CmdletBinding()] param([string]$AppId, [string]$ObjectId, [string]$DisplayName) }

    $script:AppId    = '11111111-1111-1111-1111-111111111111'
    $script:ObjectId = '22222222-2222-2222-2222-222222222222'
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    Remove-Item 'Function:\global:New-ServicePrincipal' -ErrorAction SilentlyContinue
}

Describe 'Register-EXOServicePrincipal' {
    It 'invokes New-ServicePrincipal with the supplied identifiers' {
        Mock -ModuleName EXORBACforAppManagement New-ServicePrincipal { [pscustomobject]@{ DisplayName = $DisplayName } }

        $null = Register-EXOServicePrincipal -AppId $script:AppId -ObjectId $script:ObjectId -DisplayName 'Contoso_SP' -Confirm:$false
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ServicePrincipal -Times 1 -ParameterFilter {
            $AppId -eq $script:AppId -and $ObjectId -eq $script:ObjectId -and $DisplayName -eq 'Contoso_SP'
        }
    }

    It 'does not create under -WhatIf' {
        Mock -ModuleName EXORBACforAppManagement New-ServicePrincipal { }
        $null = Register-EXOServicePrincipal -AppId $script:AppId -ObjectId $script:ObjectId -DisplayName 'Contoso_SP' -WhatIf
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName New-ServicePrincipal -Times 0
    }
}
