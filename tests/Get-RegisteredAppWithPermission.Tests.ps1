#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    function global:Get-MgServicePrincipal { [CmdletBinding()] param([string]$Filter, [string]$ServicePrincipalId) }
    function global:Get-ManagementRoleAssignment { [CmdletBinding()] param([string]$Role, [string]$Identity) }

    $script:Assignments = @(
        [pscustomobject]@{ Name = 'AppMailSend-Contoso'; Role = 'Application Mail.Send'; RoleAssigneeName = 'Contoso_SP'; RoleAssigneeType = 'ServicePrincipal'; Enabled = $true }
        [pscustomobject]@{ Name = 'AppCldR-Contoso'; Role = 'Application Calendars.Read'; RoleAssigneeName = 'Contoso_SP'; RoleAssigneeType = 'ServicePrincipal'; Enabled = $true }
        [pscustomobject]@{ Name = 'AppMailSend-Fabrikam'; Role = 'Application Mail.Send'; RoleAssigneeName = 'Fabrikam_SP'; RoleAssigneeType = 'ServicePrincipal'; Enabled = $false }
        [pscustomobject]@{ Name = 'AppMailSend-Helpdesk'; Role = 'Application Mail.Send'; RoleAssigneeName = 'Helpdesk'; RoleAssigneeType = 'RoleGroup'; Enabled = $true }
    )
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgServicePrincipal','Get-ManagementRoleAssignment') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'Get-RegisteredAppWithPermission' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment {
            $script:Assignments | Where-Object { -not $Role -or $_.Role -eq $Role }
        }

        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal {
            switch ($Filter) {
                "displayName eq 'Contoso_SP'" { @() ; break }
                "displayName eq 'Contoso'" {
                    [pscustomobject]@{
                        DisplayName = 'Contoso'
                        AppId       = '11111111-1111-1111-1111-111111111111'
                        Id          = '22222222-2222-2222-2222-222222222222'
                    }
                    break
                }
                "displayName eq 'Fabrikam_SP'" { @() ; break }
                "displayName eq 'Fabrikam'" {
                    [pscustomobject]@{
                        DisplayName = 'Fabrikam'
                        AppId       = '33333333-3333-3333-3333-333333333333'
                        Id          = '44444444-4444-4444-4444-444444444444'
                    }
                    break
                }
                default { @() }
            }
        }
    }

    It 'returns one row per distinct registered application' {
        $r = Get-RegisteredAppWithPermission

        $r.Count | Should -Be 2
        ($r.DisplayName | Sort-Object) | Should -Be @('Contoso', 'Fabrikam')
        ($r | Where-Object DisplayName -eq 'Contoso').Roles | Should -Be @('Application Calendars.Read', 'Application Mail.Send')
    }

    It 'normalizes a short role filter before querying EXO' {
        $r = Get-RegisteredAppWithPermission -Role 'Mail.Send'

        $r.Count | Should -Be 2
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Get-ManagementRoleAssignment -Times 1 -ParameterFilter { $Role -eq 'Application Mail.Send' }
    }

    It 'filters by enabled state after collecting assignments' {
        $r = Get-RegisteredAppWithPermission -Enabled $false

        $r.Count | Should -Be 1
        $r.DisplayName | Should -Be 'Fabrikam'
        $r.DisabledAssignmentCount | Should -Be 1
    }
}
