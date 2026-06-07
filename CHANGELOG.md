# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `Set-RBACforAppEntry` — reconcile/"make it so" companion to `Test-RBACforAppEntry` and
  `New-RBACforAppEntry`. Resolves the application and brings its Exchange Online RBAC components to the
  desired state, changing only what is needed: creates the scoped Unified Group and the Exchange
  Online service principal pointer if missing, adds any requested `-Members` not already in the group
  (additive — never removes members), and ensures one role assignment per role scoped to the target
  group (creating a missing one, or re-scoping one that points elsewhere). An optional
  `-NewGroupPrefix`/`-NewGroupName` moves the role assignments onto a different scoping group (created
  if needed; the old group is left in place). Each change is gated by `SupportsShouldProcess`
  (`ConfirmImpact='High'`) so it is confirmed interactively; under `-WhatIf` nothing is changed.
  Returns a `[pscustomobject]` with the current/target group names, a `GroupChanged` flag, which
  components were created, members added/already-present, role assignments created/re-scoped/unchanged,
  and an overall `IsValid` flag.
- `Remove-RBACforAppEntry` — safe teardown counterpart to `New-RBACforAppEntry`. Resolves the
  application, derives the scoped Unified Group name, and removes this app's Exchange Online role
  assignments and the Unified Group — but only after confirming the group is no longer in use (no
  foreign role assignments scoped to it and no members beyond the `-BootstrapMember` placeholder). On
  an unsafe condition it aborts and removes nothing, returning a `[pscustomobject]` summary with a
  `Reason`, the offending foreign assignments / real members, and an `IsRemoved` flag. Leaves the
  shared Exchange Online service principal pointer in place and supports `-WhatIf`/`-Confirm`.
- `Test-RBACforAppEntry` — read-only validator that confirms a registered application has every
  component `New-RBACforAppEntry` creates: the resolvable service principal, the scoped Unified
  Group, the Exchange Online service principal pointer, and one role assignment per role (matched by
  the deterministic assignment name). Mirrors `New-RBACforAppEntry`'s `-Role`/`-GroupPrefix` defaults
  and optionally verifies `-Members` against the group. Returns a `[pscustomobject]` with
  per-component flags, a `Missing` list, and an overall `IsValid`.

## [0.4.1] - 2026-06-07

### Added
- `New-RBACforAppEntry` and `New-RBACforAppUnifiedGroup` now report the Unified Group owner as
  `OwnerRequested` (the `-ManagedBy` input) and `OwnerAdded` (the owner actually applied/in place),
  mirroring the existing `MembersRequested`/`MembersAdded` reporting. The owner is resolved via
  `Get-Recipient` (like members); if it cannot be resolved the requested value is used as-is and a
  warning is emitted.

### Changed
- Renamed the module from `RBACforAppGovern` to **`EXORBACforAppManagement`**. This is a module
  identity rename only (directory, manifest/loader files, and all `-ModuleName` / `Import-Module` /
  `InModuleScope` references); the public function names and the manifest `GUID` are unchanged.
- `New-RBACforAppUnifiedGroup` now returns a summary `[pscustomobject]`
  (`Name`, `DisplayName`, `OwnerRequested`, `OwnerAdded`, `AlreadyExisted`, `Group`) instead of the
  raw Exchange Online group object. The underlying group object remains available via `.Group`.
- `New-RBACforAppUnifiedGroup` now creates the Unified Group with the smallest set of essential
  attributes (DisplayName/Name/Alias, AccessType Private, the creation-only HiddenGroupMembershipEnabled,
  owner, and bootstrap member), then applies the remaining settings (member edit, auto-subscribe,
  calendar subscribe, language, subscription, address-list visibility, connectors) via a single
  follow-up `Set-UnifiedGroup` call.

## [0.4.0] - 2026-06-07

### Added
- `Convert-ApplicationAccessPolicyToRBAC` — migrates legacy Exchange Online Application Access
  Policies to RBAC for Applications. For each `RestrictAccess` policy it resolves the service
  principal, derives the application roles from the app's granted Microsoft Graph application
  permissions (the set Application Access Policies supported, mapped to their App RBAC role names),
  copies the original scope group's members, and delegates to `New-RBACforAppEntry`. `-Role`
  overrides the auto-derived roles; `DenyAccess` policies are skipped (no additive RBAC equivalent).
- Private helper `Get-LegacyScopeRoleMap` — maps each legacy Application Access Policy permission
  scope (e.g. `Mail.Read`, EWS `full_access_as_app`) to its App RBAC role name.

## [0.3.0] - 2026-06-07

### Added
- `Get-RBACforAppEntry -RoleAssigneeType` — filters by the assignment's assignee type. Defaults to
  `ServicePrincipal` so only application (service-principal) assignments are returned; pass `All` for
  every type, or a specific type (`User`, `RoleGroup`, etc.) to narrow.
- `Get-RegisteredAppWithPermission` — inventories distinct registered applications that currently
  hold supported Exchange Online application-role assignments, with optional `-Role` and `-Enabled`
  filtering.

## [0.2.1] - 2026-06-07

### Changed

- `-ManagedBy` now defaults to `GraphAPI-Dummy-owner` (in `New-RBACforAppEntry` and
  `New-RBACforAppUnifiedGroup`). A Unified Group owner must be a valid owner account, distinct from a
  plain member, so the previous `GraphAPI-Dummy` owner default caused `New-UnifiedGroup` to fail.
  `Members` / `BootstrapMember` still default to `GraphAPI-Dummy`.

## [0.2.0] - 2026-06-07

### Added

- `New-RBACforAppUnifiedGroup` — public function that ensures/creates and configures the scoped
  Microsoft 365 Unified Group (extracted from `New-RBACforAppEntry`).
- `Register-EXOServicePrincipal` — public function that creates the Exchange Online service
  principal pointer for an Entra application (extracted from `New-RBACforAppEntry`).
- `.github/workflows/release.yml` — builds, tests, and publishes a GitHub release (with the packaged
  module zip) when a `v*` tag is pushed.

### Changed

- `New-RBACforAppEntry` now delegates Unified Group creation and EXO service principal registration
  to the two new functions instead of inlining that logic. Observable behavior is unchanged.

## [0.1.0] - 2026-06-07

### Added

- Initial `EXORBACforAppManagement` module packaging the `New-RBACforAppEntry`, `New-RegisteredApp`, and
  `Get-RBACforAppEntry` functions, with Pester tests, a `build.ps1` pipeline, and CI.
