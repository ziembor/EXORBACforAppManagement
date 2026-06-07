#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force

    # Global stubs so the module scope can resolve them and Pester can mock them on CI
    # (where Microsoft.Graph / ExchangeOnlineManagement are not installed). Parameters the
    # code passes (e.g. -Role) must be declared so the mock can bind and filter on them.
    function global:Get-MgServicePrincipal { [CmdletBinding()] param([string]$Filter, [string]$ServicePrincipalId) }
    function global:Get-ManagementRoleAssignment { [CmdletBinding()] param([string]$Role, [string]$Identity) }

    $script:Assignments = @(
        [pscustomobject]@{ Name = 'AppMailSend-Contoso'; Role = 'Application Mail.Send'; RoleAssigneeName = 'Contoso_SP'; RoleAssigneeType = 'ServicePrincipal'; CustomRecipientWriteScope = 'scope'; RecipientWriteScope = 'CustomRecipientScope'; Enabled = $true; Guid = [guid]::NewGuid(); Identity = 'AppMailSend-Contoso' }
        [pscustomobject]@{ Name = 'Mail Recipients-Admin'; Role = 'Mail Recipients'; RoleAssigneeName = 'Org Management'; RoleAssigneeType = 'RoleGroup'; CustomRecipientWriteScope = $null; RecipientWriteScope = 'Organization'; Enabled = $true; Guid = [guid]::NewGuid(); Identity = 'Mail Recipients-Admin' }
        [pscustomobject]@{ Name = 'AppCldR-Fabrikam'; Role = 'Application Calendars.Read'; RoleAssigneeName = 'Fabrikam_SP'; RoleAssigneeType = 'ServicePrincipal'; CustomRecipientWriteScope = 'scope2'; RecipientWriteScope = 'CustomRecipientScope'; Enabled = $false; Guid = [guid]::NewGuid(); Identity = 'AppCldR-Fabrikam' }
        [pscustomobject]@{ Name = 'AppMailR-Helpdesk'; Role = 'Application Mail.Read'; RoleAssigneeName = 'Helpdesk'; RoleAssigneeType = 'RoleGroup'; CustomRecipientWriteScope = 'scope3'; RecipientWriteScope = 'CustomRecipientScope'; Enabled = $true; Guid = [guid]::NewGuid(); Identity = 'AppMailR-Helpdesk' }
        [pscustomobject]@{ Name = 'AppMailboxSettings-Tailspin'; Role = 'Application MailboxSettings.Read'; RoleAssigneeName = 'Tailspin_SP'; RoleAssigneeType = 'ServicePrincipal'; CustomRecipientWriteScope = $null; RecipientWriteScope = 'Group'; Enabled = $true; Guid = [guid]::NewGuid(); Identity = 'AppMailboxSettings-Tailspin' }
        [pscustomobject]@{ Name = 'AppMailWide-Northwind'; Role = 'Application Mail.Send'; RoleAssigneeName = 'Northwind_SP'; RoleAssigneeType = 'ServicePrincipal'; CustomRecipientWriteScope = $null; RecipientWriteScope = 'Organization'; Enabled = $true; Guid = [guid]::NewGuid(); Identity = 'AppMailWide-Northwind' }
    )
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
    foreach ($n in 'Get-MgServicePrincipal','Get-ManagementRoleAssignment') {
        Remove-Item "Function:\global:$n" -ErrorAction SilentlyContinue
    }
}

Describe 'Get-RBACforAppEntry (no filter)' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment { $script:Assignments }
    }

    It 'returns only ServicePrincipal application-role assignments by default' {
        $r = Get-RBACforAppEntry
        $r.Count | Should -Be 3
        ($r.Role | Sort-Object -Unique) | Should -Be @('Application Calendars.Read', 'Application Mail.Send', 'Application MailboxSettings.Read')
        ($r.RoleAssigneeType | Sort-Object -Unique) | Should -Be @('ServicePrincipal')
        ($r.RecipientScope | Sort-Object -Unique) | Should -Be @('CustomRecipientScope', 'Group')
    }

    It 'excludes application assignments outside group or custom recipient scopes' {
        $r = Get-RBACforAppEntry
        $r.Name | Should -Not -Contain 'AppMailWide-Northwind'
        $r.RecipientScope | Should -Not -Contain 'Organization'
    }

    It 'projects the expected shape' {
        $r = Get-RBACforAppEntry | Select-Object -First 1
        $r.PSObject.Properties.Name | Should -Be @('Name','Role','RoleAssigneeName','RoleAssigneeType','Scope','RecipientScope','Enabled','Guid','Identity')
    }
}

Describe 'Get-RBACforAppEntry -Role' {
    It 'normalizes a short role and queries EXO with the full role name' {
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment {
            $script:Assignments | Where-Object { $_.Role -eq $Role }
        } -ParameterFilter { $Role -eq 'Application Mail.Send' }

        $r = Get-RBACforAppEntry -Role 'Mail.Send'
        $r.Role | Should -Be 'Application Mail.Send'
        Should -Invoke -ModuleName EXORBACforAppManagement -CommandName Get-ManagementRoleAssignment -Times 1 -ParameterFilter { $Role -eq 'Application Mail.Send' }
    }
}

Describe 'Get-RBACforAppEntry -RoleAssigneeType' {
    BeforeEach {
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment { $script:Assignments }
    }

    It 'returns every assignee type with -RoleAssigneeType All' {
        $r = Get-RBACforAppEntry -RoleAssigneeType All
        $r.Count | Should -Be 4
        ($r.RoleAssigneeType | Sort-Object -Unique) | Should -Be @('RoleGroup', 'ServicePrincipal')
        ($r.RecipientScope | Sort-Object -Unique) | Should -Be @('CustomRecipientScope', 'Group')
    }

    It 'filters to a specific assignee type' {
        $r = Get-RBACforAppEntry -RoleAssigneeType RoleGroup
        $r.Count | Should -Be 1
        $r.RoleAssigneeName | Should -Be 'Helpdesk'
    }
}

Describe 'Get-RBACforAppEntry application filter' {
    It 'keeps only assignments matching the resolved service principal' {
        Mock -ModuleName EXORBACforAppManagement Get-MgServicePrincipal {
            [pscustomobject]@{ DisplayName = 'Contoso'; AppId = '11111111-1111-1111-1111-111111111111'; Id = '22222222-2222-2222-2222-222222222222' }
        }
        Mock -ModuleName EXORBACforAppManagement Get-ManagementRoleAssignment { $script:Assignments }

        $r = Get-RBACforAppEntry -RegisteredAppName 'Contoso'
        $r.Count | Should -Be 1
        $r.RoleAssigneeName | Should -Be 'Contoso_SP'
    }
}
