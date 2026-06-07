#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'EXORBACforAppManagement' 'EXORBACforAppManagement.psd1') -Force
}

AfterAll {
    Remove-Module EXORBACforAppManagement -Force -ErrorAction SilentlyContinue
}

Describe 'Get-SafeName' {
    It 'strips characters that are not letters, digits, or dash' {
        InModuleScope EXORBACforAppManagement {
            Get-SafeName -s 'Um365RAo1-Contoso Mail App!@#' | Should -Be 'Um365RAo1-ContosoMailApp'
        }
    }

    It 'caps the result at the requested length' {
        InModuleScope EXORBACforAppManagement {
            $out = Get-SafeName -s ('a' * 100) -max 63
            $out.Length | Should -Be 63
        }
    }

    It 'defaults the cap to 63' {
        InModuleScope EXORBACforAppManagement {
            (Get-SafeName -s ('b' * 200)).Length | Should -Be 63
        }
    }
}

Describe 'Get-NormalizeRole' {
    It 'maps a known short name to the Application role' {
        InModuleScope EXORBACforAppManagement {
            Get-NormalizeRole 'Mail.Send' | Should -Be 'Application Mail.Send'
        }
    }

    It 'passes through an already-normalized role unchanged' {
        InModuleScope EXORBACforAppManagement {
            Get-NormalizeRole 'Application Calendars.Read' | Should -Be 'Application Calendars.Read'
        }
    }

    It 'passes through an unknown role unchanged' {
        InModuleScope EXORBACforAppManagement {
            Get-NormalizeRole 'Totally.Unknown' | Should -Be 'Totally.Unknown'
        }
    }
}

Describe 'ConvertTo-AppRole' {
    It 'prefixes a short name with Application' {
        InModuleScope EXORBACforAppManagement {
            ConvertTo-AppRole 'Mail.Send' | Should -Be 'Application Mail.Send'
        }
    }

    It 'leaves an already-prefixed role unchanged' {
        InModuleScope EXORBACforAppManagement {
            ConvertTo-AppRole 'Application Mail.Send' | Should -Be 'Application Mail.Send'
        }
    }
}
