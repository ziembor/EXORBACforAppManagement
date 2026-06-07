<#
.SYNOPSIS
Build script for the EXORBACforAppManagement module.

.DESCRIPTION
Runs the module's quality gates and packaging steps. Tasks:
  Init    - ensure Pester (>=5) and PSScriptAnalyzer are available
  Clean   - remove build output and test results
  Analyze - run PSScriptAnalyzer over src (fails on Error severity)
  Test    - run the Pester suite (emits testResults.xml in NUnit format)
  Build   - copy the module to ./output and validate the manifest
  All     - Init, Clean, Analyze, Test, Build (default)

.EXAMPLE
./build.ps1
Runs the full pipeline.

.EXAMPLE
./build.ps1 -Task Test
Runs only the tests.
#>
[CmdletBinding()]
param(
    [ValidateSet('All', 'Init', 'Clean', 'Analyze', 'Test', 'Build')]
    [string[]] $Task = 'All'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$RepoRoot     = $PSScriptRoot
$ModuleName   = 'EXORBACforAppManagement'
$SrcPath      = Join-Path $RepoRoot 'src' $ModuleName
$ManifestPath = Join-Path $SrcPath "$ModuleName.psd1"
$TestsPath    = Join-Path $RepoRoot 'tests'
$OutputPath   = Join-Path $RepoRoot 'output'
$TestResults  = Join-Path $RepoRoot 'testResults.xml'
$SettingsPath = Join-Path $RepoRoot 'PSScriptAnalyzerSettings.psd1'

function Resolve-TaskList {
    param([string[]] $Requested)
    if ($Requested -contains 'All') {
        return @('Init', 'Clean', 'Analyze', 'Test', 'Build')
    }
    return $Requested
}

function Invoke-Init {
    Write-Host '==> Init' -ForegroundColor Cyan

    $needPester = -not (Get-Module Pester -ListAvailable | Where-Object Version -ge ([version]'5.0.0'))
    $needPSSA   = -not (Get-Module PSScriptAnalyzer -ListAvailable)
    if (-not ($needPester -or $needPSSA)) { return }

    # Prefer PSResourceGet: it talks to PSGallery directly and does NOT need the legacy NuGet
    # package provider, whose bootstrap (Install-PackageProvider) fails on hosts with multiple
    # duplicate PackageManagement/PowerShellGet versions ("Collection was modified ...").
    $resourceGet = Get-Module Microsoft.PowerShell.PSResourceGet -ListAvailable |
        Sort-Object Version -Descending | Select-Object -First 1
    if ($resourceGet) {
        Import-Module $resourceGet -Force
        if ($needPester) {
            Write-Host 'Installing Pester 5 (PSResourceGet)...'
            Install-PSResource Pester -Version '[5.0.0,)' -Repository PSGallery -TrustRepository -Scope CurrentUser
        }
        if ($needPSSA) {
            Write-Host 'Installing PSScriptAnalyzer (PSResourceGet)...'
            Install-PSResource PSScriptAnalyzer -Repository PSGallery -TrustRepository -Scope CurrentUser
        }
        return
    }

    # Fallback for hosts without PSResourceGet: ensure the NuGet provider exists and PSGallery is
    # trusted so Install-Module runs non-interactively (avoids the untrusted-repository prompt).
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Write-Host 'Installing NuGet package provider...'
        Install-PackageProvider -Name NuGet -MinimumVersion '2.8.5.201' -Force -Scope CurrentUser | Out-Null
    }
    $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    if ($needPester) {
        Write-Host 'Installing Pester 5...'
        Install-Module Pester -MinimumVersion 5.0.0 -Force -Scope CurrentUser -SkipPublisherCheck -Repository PSGallery
    }
    if ($needPSSA) {
        Write-Host 'Installing PSScriptAnalyzer...'
        Install-Module PSScriptAnalyzer -Force -Scope CurrentUser -Repository PSGallery
    }
}

function Invoke-Clean {
    Write-Host '==> Clean' -ForegroundColor Cyan
    if (Test-Path $OutputPath)  { Remove-Item $OutputPath -Recurse -Force }
    if (Test-Path $TestResults) { Remove-Item $TestResults -Force }
}

function Invoke-Analyze {
    Write-Host '==> Analyze' -ForegroundColor Cyan
    Import-Module PSScriptAnalyzer -Force
    $findings = Invoke-ScriptAnalyzer -Path $SrcPath -Recurse -Settings $SettingsPath
    if ($findings) {
        $findings | Format-Table -AutoSize | Out-String | Write-Host
    }
    $errors = @($findings | Where-Object Severity -eq 'Error')
    if ($errors.Count -gt 0) {
        throw "PSScriptAnalyzer found $($errors.Count) error(s)."
    }
    Write-Host 'Analyze passed (no errors).' -ForegroundColor Green
}

function Invoke-Test {
    Write-Host '==> Test' -ForegroundColor Cyan
    Import-Module Pester -MinimumVersion 5.0.0 -Force
    $config = New-PesterConfiguration
    $config.Run.Path = $TestsPath
    $config.Run.PassThru = $true
    $config.Output.Verbosity = 'Detailed'
    $config.TestResult.Enabled = $true
    $config.TestResult.OutputFormat = 'NUnitXml'
    $config.TestResult.OutputPath = $TestResults
    $result = Invoke-Pester -Configuration $config
    if ($result.FailedCount -gt 0) {
        throw "$($result.FailedCount) test(s) failed."
    }
    Write-Host "Tests passed ($($result.PassedCount))." -ForegroundColor Green
}

function Invoke-Build {
    Write-Host '==> Build' -ForegroundColor Cyan
    $dest = Join-Path $OutputPath $ModuleName
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item -Path (Join-Path $SrcPath '*') -Destination $dest -Recurse -Force
    $null = Test-ModuleManifest -Path (Join-Path $dest "$ModuleName.psd1")
    Write-Host "Built module at $dest" -ForegroundColor Green
}

foreach ($t in (Resolve-TaskList -Requested $Task)) {
    switch ($t) {
        'Init'    { Invoke-Init }
        'Clean'   { Invoke-Clean }
        'Analyze' { Invoke-Analyze }
        'Test'    { Invoke-Test }
        'Build'   { Invoke-Build }
    }
}

Write-Host 'Done.' -ForegroundColor Green
