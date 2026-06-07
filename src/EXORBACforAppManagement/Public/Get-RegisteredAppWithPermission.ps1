<#
.SYNOPSIS
Lists registered applications that hold Exchange Online application RBAC permissions.

.DESCRIPTION
Get-RegisteredAppWithPermission returns one object per distinct Entra service principal
that has one or more Exchange Online management role assignments for application roles.
It is the app-centric (one row per application) inventory counterpart to Get-RBACforAppEntry,
which is assignment-centric. Unlike the New-/Get-RBACforAppEntry functions it has no
ByName/ByAppId/BySpObjectId parameter sets: it performs a tenant-wide sweep rather than a
single-application lookup.

Processing steps:

  1. Decide which roles to query. With -Role, the supplied short names (e.g. Mail.Send) are
     normalized to their full names (Application Mail.Send) via Get-NormalizeRole and deduped.
     Without -Role, every role supported by New-RBACforAppEntry is queried (the keys of the
     shared Get-AppRoleMap table).
  2. Query Exchange Online once per role via Get-ManagementRoleAssignment -Role (EXO does the
     role filtering). Roles with no assignments simply contribute nothing.
  3. Keep only application assignments to service principals (Role -like 'Application *' and
     RoleAssigneeType -eq 'ServicePrincipal').
  4. Apply the optional -Enabled filter.
  5. Group the surviving assignments by RoleAssigneeName (this collapses many assignments into
     one row per application) and resolve each distinct assignee back to its Entra service
     principal through Microsoft Graph. Because the EXO assignee is usually the service principal
     pointer named "<DisplayName>_SP", resolution is attempted by display name against both the
     raw assignee and the "_SP"-stripped variant: exactly one Graph match wins; more than one
     match writes an error (re-run with a narrower role filter or resolve the Entra duplicates);
     no match falls through to the next candidate.
  6. Emit one object per application. When Graph resolution fails the row is still returned with
     EXO-only details (DisplayName falls back to the assignee with "_SP" stripped; AppId and
     ServicePrincipalId are null) and a warning is written.

.PARAMETER Role
One or more application roles to query. Short names such as Mail.Send are accepted and
normalized to Application Mail.Send. When omitted, every role supported by
New-RBACforAppEntry is queried.

.PARAMETER Enabled
Return only enabled ($true) or only disabled ($false) assignments. Omit to return both.

.EXAMPLE
Get-RegisteredAppWithPermission

Returns every registered application that currently has one or more supported
application-role assignments.

.EXAMPLE
Get-RegisteredAppWithPermission -Role 'Mail.Send'

Returns the registered applications that hold the Application Mail.Send role.

.OUTPUTS
PSCustomObject

One object per distinct registered application, with the following properties:

  DisplayName             - Graph display name (or the "_SP"-stripped assignee when unresolved).
  AppId                   - Application (client) id from Graph (null when unresolved).
  ServicePrincipalId      - Service principal object id from Graph (null when unresolved).
  ExoServicePrincipal     - The raw Exchange Online assignee name (e.g. Contoso_SP).
  Roles                   - Sorted, unique application roles the app holds.
  RoleAssignmentNames     - Sorted, unique management role assignment names.
  AssignmentCount         - Total matched assignments for the app.
  EnabledAssignmentCount  - Count of those that are enabled.
  DisabledAssignmentCount - Count of those that are disabled.

.NOTES
Requires a connected Exchange Online session (Get-ManagementRoleAssignment) and a
connected Microsoft Graph session (Get-MgServicePrincipal).

Performance / behavior notes:
- The function issues one Get-ManagementRoleAssignment query per role, so cost scales with the
  number of roles requested (all supported roles by default); EXO does the filtering and results
  are grouped client-side.
- Graph reverse-resolution is by display name (not AppId), which is why duplicate display names
  produce the ambiguity error and why the "_SP"-stripped variant is also tried.
- Unlike Get-RBACforAppEntry, this function does not filter on recipient scope: any
  'Application *' assignment to a service principal is counted.
#>
function Get-RegisteredAppWithPermission {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]] $Role,

        [Parameter()]
        [bool] $Enabled
    )

    process {
        $supportedRoleMap = Get-AppRoleMap
        $rolesNormalized =
            if ($PSBoundParameters.ContainsKey('Role')) {
                @($Role | ForEach-Object { Get-NormalizeRole $_ } | Select-Object -Unique)
            }
            else {
                @($supportedRoleMap.Keys)
            }

        $assignments = foreach ($roleItem in $rolesNormalized) {
            Get-ManagementRoleAssignment -Role $roleItem -ErrorAction SilentlyContinue
        }

        $assignments = @(
            $assignments |
                Where-Object {
                    $_ -and
                    $_.Role -like 'Application *' -and
                    $_.RoleAssigneeType -eq 'ServicePrincipal'
                }
        )

        if ($PSBoundParameters.ContainsKey('Enabled')) {
            $assignments = @($assignments | Where-Object { $_.Enabled -eq $Enabled })
        }

        foreach ($assignmentGroup in ($assignments | Group-Object RoleAssigneeName | Sort-Object Name)) {
            $assigneeName = [string]$assignmentGroup.Name
            $resolvedSp = $null
            $lookupNames = @($assigneeName)

            if ($assigneeName -match '_SP$') {
                $lookupNames += ($assigneeName -replace '_SP$', '')
            }

            foreach ($lookupName in ($lookupNames | Select-Object -Unique)) {
                $matchesRes = @(Get-MgServicePrincipal -Filter "displayName eq '$lookupName'" -ErrorAction Stop)
                if ($matchesRes.Count -eq 1) {
                    $resolvedSp = $matchesRes[0]
                    break
                }

                if ($matchesRes.Count -gt 1) {
                    Write-Error -Message ("Ambiguous service principal resolution for assignee '{0}' using displayName '{1}'. Re-run with a narrower role filter or resolve the duplicates in Entra." -f $assigneeName, $lookupName)
                    break
                }
            }

            if (-not $resolvedSp) {
                Write-Warning -Message ("Could not resolve EXO assignee '{0}' to a single Graph service principal; returning EXO-only details." -f $assigneeName)
            }

            $rolesForApp = @($assignmentGroup.Group.Role | Sort-Object -Unique)
            $assignmentNames = @($assignmentGroup.Group.Name | Sort-Object -Unique)

            [pscustomobject][ordered]@{
                DisplayName           = if ($resolvedSp) { $resolvedSp.DisplayName } else { ($assigneeName -replace '_SP$', '') }
                AppId                 = if ($resolvedSp) { $resolvedSp.AppId } else { $null }
                ServicePrincipalId    = if ($resolvedSp) { $resolvedSp.Id } else { $null }
                ExoServicePrincipal   = $assigneeName
                Roles                 = $rolesForApp
                RoleAssignmentNames   = $assignmentNames
                AssignmentCount       = $assignmentGroup.Count
                EnabledAssignmentCount = @($assignmentGroup.Group | Where-Object { $_.Enabled }).Count
                DisabledAssignmentCount = @($assignmentGroup.Group | Where-Object { -not $_.Enabled }).Count
            }
        }
    }
}
