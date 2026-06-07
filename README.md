# EXORBACforAppManagement

[![CI](https://github.com/ziembor/new-RBACforAppEntry/actions/workflows/ci.yml/badge.svg)](https://github.com/ziembor/new-RBACforAppEntry/actions/workflows/ci.yml)

A PowerShell module for governing **Exchange Online (EXO) Role-Based Access Control for Entra
(Azure AD) applications**. It registers applications, assigns resource-scoped EXO *Application*
role permissions (e.g. `Application Mail.Send`) scoped to a Microsoft 365 group, and reads those
assignments back.

The three functions form a **create → assign → read** flow and share the same
`ByName` / `ByAppId` / `BySpObjectId` parameter conventions, so they compose over the pipeline.

| Function | Step | What it does |
| --- | --- | --- |
| [`New-RegisteredApp`](#new-registeredapp) | create | Creates an Entra app registration and (by default) its service principal. |
| [`New-RBACforAppEntry`](#new-rbacforappentry) | assign | Creates a scoped Unified Group, ensures the EXO service principal, and assigns EXO Application roles scoped to that group. |
| [`Get-RBACforAppEntry`](#get-rbacforappentry) | read | Lists the EXO Application-role assignments. |
| [`Get-RegisteredAppWithPermission`](#get-registeredappwithpermission) | inventory | Lists distinct registered applications that currently hold supported EXO Application roles. |
| [`Test-RBACforAppEntry`](#test-rbacforappentry) | validate | Checks an application has every component `New-RBACforAppEntry` creates (SP, Unified Group, EXO service principal, role assignments) and returns an `IsValid` summary. |
| [`Convert-ApplicationAccessPolicyToRBAC`](#convert-applicationaccesspolicytorbac) | migrate | Migrates legacy Application Access Policies to RBAC for Applications, delegating to `New-RBACforAppEntry`. |
| [`New-RBACforAppUnifiedGroup`](#new-rbacforappunifiedgroup) | helper | Ensures/creates and configures the scoped Unified Group (used by `New-RBACforAppEntry`). |
| [`Register-EXOServicePrincipal`](#register-exoserviceprincipal) | helper | Creates the EXO service principal pointer for an Entra app (used by `New-RBACforAppEntry`). |

> Background: [Microsoft Learn — Role Based Access Control for Applications in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac).

## Requirements

- **PowerShell 5.1+** (developed/tested on PowerShell 7).
- The following modules installed and **connected** at runtime (they are intentionally *not*
  declared as `RequiredModules`, so the module imports without them for unit testing):
  - **Microsoft Graph** — `Connect-MgGraph`
    (used by `Get-MgServicePrincipal`, `Get-MgContext`, `New-MgApplication`, `New-MgServicePrincipal`).
    `New-RegisteredApp` needs the `Application.ReadWrite.All` scope.
  - **Exchange Online** — `Connect-ExchangeOnline`
    (used by `Get-UnifiedGroup`, `New-UnifiedGroup`, `Set-UnifiedGroup`, `Add-UnifiedGroupLinks`,
    `New-ServicePrincipal`, `Get-Recipient`, `New-ManagementRoleAssignment`, `Get-ManagementRoleAssignment`).

## Install / import

The module is run from source (it is not published to the PowerShell Gallery):

```powershell
git clone https://github.com/ziembor/new-RBACforAppEntry.git
Import-Module ./new-RBACforAppEntry/src/EXORBACforAppManagement/EXORBACforAppManagement.psd1 -Force
```

Then connect your sessions:

```powershell
Connect-MgGraph -Scopes 'Application.ReadWrite.All'
Connect-ExchangeOnline
```

> **Always preview with `-WhatIf` first.** `New-RBACforAppEntry` (`ConfirmImpact='High'`) and
> `New-RegisteredApp` (`ConfirmImpact='Medium'`) gate every mutating step behind `ShouldProcess`.

## Usage

### End-to-end

```powershell
# 1. Register the app + service principal, then 2. scope EXO RBAC for it (pipeline):
New-RegisteredApp -DisplayName 'Contoso Mail App' |
    New-RBACforAppEntry -Members 'shared@contoso.com' -Role 'Mail.Send' -WhatIf -Verbose

# 3. Read the resulting assignments:
Get-RBACforAppEntry -RegisteredAppName 'Contoso Mail App'

# Inventory distinct registered applications that already hold supported EXO app permissions:
Get-RegisteredAppWithPermission
```

### New-RegisteredApp

Creates an Entra application registration and, unless `-SkipServicePrincipal`, its service
principal. Emits `AppId` / `ServicePrincipalId` so it can pipe into `New-RBACforAppEntry`.

```powershell
New-RegisteredApp -DisplayName 'Contoso Mail App' -WhatIf -Verbose
```

### New-RBACforAppEntry

Resolves the service principal (by name, AppId, or SP object id), creates a scoped Unified Group
named `"{GroupPrefix}-{DisplayName}"`, adds members, ensures the EXO service principal, and creates
one role assignment per role — each scoped to the group via `-RecipientGroupScope`. Short role
names such as `Mail.Send` are normalized to `Application Mail.Send`.

```powershell
# By AppId, assigning a single role to a shared mailbox:
New-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' `
    -Members 'sharedmailbox@contoso.com' -Role 'Mail.Send' -Verbose

# By SP object id, multiple roles, custom group prefix:
New-RBACforAppEntry -SpObjectId '11111111-2222-3333-4444-555555555555' `
    -Role 'Application Calendars.Read','Application Contacts.Read' -GroupPrefix 'Um365Prod'
```

Returns a summary `[pscustomobject]` (resolved identity, group name, normalized roles, assignment
names, `Warnings`, `Errors`) and also exports it to `$env:TEMP\<name>_<timestamp>.clixml`.

### Get-RBACforAppEntry

Returns EXO management role assignments for Application roles (`Application *`). With no arguments
it returns all of them; filter by application and/or role, plus optional `-Enabled`.

```powershell
Get-RBACforAppEntry                                            # every application-role assignment
Get-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -Role 'Mail.Send'
Get-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' | Format-Table Name,Role,Scope
```

> `Get-ManagementRoleAssignment` has no `-App` parameter, so role filtering uses native `-Role`
> while the application filter is applied client-side against each assignment's `RoleAssigneeName`
> and `Name`.

### Get-RegisteredAppWithPermission

Returns one row per distinct registered application that already holds one or more Exchange Online
Application-role assignments. By default it inventories the full set of roles supported by
`New-RBACforAppEntry`; you can narrow it with `-Role`.

```powershell
Get-RegisteredAppWithPermission
Get-RegisteredAppWithPermission -Role 'Mail.Send'
```

### Test-RBACforAppEntry

Read-only check that an application has every component `New-RBACforAppEntry` provisions: the
resolvable service principal, the scoped Unified Group, the Exchange Online service principal
pointer, and one role assignment per role (looked up by the deterministic assignment name). It
mirrors `New-RBACforAppEntry`'s `-Role` / `-GroupPrefix` defaults and optionally verifies `-Members`
against the group's membership. Returns a summary `[pscustomobject]` with per-component flags
(`ServicePrincipalExists`, `UnifiedGroupExists`, `ExoServicePrincipalExists`), the
expected/found/missing role assignments, a `Missing` list, and an overall `IsValid`.

```powershell
Test-RBACforAppEntry -RegisteredAppName 'Contoso Mail App'
Test-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' -Role 'Mail.Send','Calendars.Read' -Members 'shared@contoso.com'
```

### Convert-ApplicationAccessPolicyToRBAC

Migrates legacy Exchange Online Application Access Policies to RBAC for Applications: derives the
roles from the app's granted Microsoft Graph application permissions, copies the original scope
group's members, and delegates to `New-RBACforAppEntry`. `DenyAccess` policies are skipped.

```powershell
Convert-ApplicationAccessPolicyToRBAC -WhatIf -Verbose
```

### New-RBACforAppUnifiedGroup

Ensures the scoped, private/hidden Unified Group exists and is configured (subscription, address
list, connectors disabled). Returns a summary object (`OwnerRequested`, `OwnerAdded`,
`AlreadyExisted`, and the underlying `Group`). `New-RBACforAppEntry` delegates to it, but it can be
used on its own.

```powershell
New-RBACforAppUnifiedGroup -Name 'Um365RAo1-ContosoMailApp' -WhatIf -Verbose
```

### Register-EXOServicePrincipal

Creates the Exchange Online service principal pointer that links an Entra application into EXO so it
can receive application RBAC assignments.

```powershell
Register-EXOServicePrincipal -AppId '1111...' -ObjectId '2222...' -DisplayName 'Contoso_SP' -WhatIf
```

## Project layout

```
src/EXORBACforAppManagement/
  EXORBACforAppManagement.psd1          # manifest
  EXORBACforAppManagement.psm1          # loader: dot-sources Private + Public, exports Public only
  Public/                        # 8 exported functions (see table above)
  Private/                       # Get-SafeName, Get-NormalizeRole, ConvertTo-AppRole
tests/                           # Pester v5 tests
build.ps1                        # Init / Clean / Analyze / Test / Build
PSScriptAnalyzerSettings.psd1    # analyzer config (build fails only on Error severity)
.github/workflows/ci.yml         # CI: ./build.ps1 -Task All on ubuntu-latest
.github/workflows/release.yml    # Release on v* tag -> GitHub release with module zip
CHANGELOG.md                     # version history
```

## Build & test

`build.ps1` is the entry point for all quality gates (CI runs `./build.ps1 -Task All`):

```powershell
./build.ps1                 # All: Init, Clean, Analyze, Test, Build
./build.ps1 -Task Test      # run the Pester suite only
./build.ps1 -Task Analyze   # run PSScriptAnalyzer only
```

- **Init** installs Pester (>= 5) and PSScriptAnalyzer if missing.
- **Analyze** runs PSScriptAnalyzer over `src`; fails only on Error-severity findings.
- **Test** runs Pester and writes `testResults.xml` (NUnit).
- **Build** assembles the module into `output/EXORBACforAppManagement` and validates the manifest.

Tests mock the Graph/EXO cmdlets, so the suite runs without those modules installed (this is what
CI does on `ubuntu-latest`).

## Releasing

Bump `ModuleVersion` in the manifest, add a [`CHANGELOG.md`](CHANGELOG.md) entry, then push a
matching tag:

```powershell
git tag v0.2.0
git push origin v0.2.0
```

The `release.yml` workflow builds/tests the module, packages it, and publishes a GitHub release with
the module zip attached.

## Contributing

Work on feature branches and open a PR into `main`; CI must be green. When adding a role, update
both role tables (`$roleMap` in `Private/Get-NormalizeRole.ps1` and `$shortRoleMap` in
`Public/New-RBACforAppEntry.ps1`). See [`AGENTS.md`](AGENTS.md) for deeper architecture notes.

## License

Licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
