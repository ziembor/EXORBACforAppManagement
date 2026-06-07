<#
.SYNOPSIS
Gets Exchange Online RBAC role assignments for application roles (the roles that
New-RBACforAppEntry creates).

.DESCRIPTION
Get-RBACforAppEntry retrieves Exchange Online management role assignments whose role
is an application role (named "Application <permission>", e.g. "Application Mail.Send").
It returns only assignments scoped the way this module creates them: group-scoped or
custom-recipient-scoped entries. These are the assignments produced by New-RBACforAppEntry.

By default service-principal application role assignments are returned, limited to
RecipientWriteScope values of Group or CustomRecipientScope. You can narrow the results
to a specific application (by display name, AppId, or service principal object id)
and/or to specific roles. Filtering by application is done client-side because
Get-ManagementRoleAssignment has no -App parameter; the resolved identity is matched
against each assignment's RoleAssigneeName and Name.

.PARAMETER RegisteredAppName
Display name of the application / service principal to filter assignments by. This is
matched (via the resolved service principal) against the assignment's assignee and name.

.PARAMETER AppId
Application (client) id of the application to filter by. GUID-validated.

.PARAMETER SpObjectId
Object id of the service principal to filter by. GUID-validated.

.PARAMETER Role
One or more application roles to filter by. Short names such as Mail.Send are accepted
and normalized to "Application Mail.Send".

.PARAMETER Enabled
Return only enabled ($true) or only disabled ($false) assignments. Omit to return both.

.PARAMETER RoleAssigneeType
Filter by the assignment's RoleAssigneeType. Defaults to 'ServicePrincipal' (application
assignments). Use 'All' to return every assignee type within the supported Group /
CustomRecipientScope recipient scopes, or pass a specific type to narrow.

.EXAMPLE
Get-RBACforAppEntry

Returns service-principal application role assignments that are group-scoped or use a
custom recipient scope (the default behavior).

.EXAMPLE
Get-RBACforAppEntry -RoleAssigneeType All

Returns application role assignments regardless of assignee type, still limited to
Group / CustomRecipientScope recipient scopes.

.EXAMPLE
Get-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -Role 'Mail.Send'

Returns the Application Mail.Send assignments scoped to the resolved Contoso Mail App.

.EXAMPLE
Get-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' | Format-Table Name,Role,Scope

Filters by AppId and formats the key columns.

.OUTPUTS
PSCustomObject

One object per matching assignment with Name, Role, assignee, scope, enabled state,
and identity fields.

.NOTES
Requires a connected Exchange Online session (Connect-ExchangeOnline) for
Get-ManagementRoleAssignment. When an application filter is supplied, a connected
Microsoft Graph session (Connect-MgGraph) is also required to resolve the service
principal. Companion to New-RBACforAppEntry.
#>
function Get-RBACforAppEntry {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ByName')]
        [Alias('DisplayName','Name')]
        [ValidateNotNullOrEmpty()]
        [string] $RegisteredAppName,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ByAppId')]
        [Alias('ClientId','ApplicationId')]
        [ValidatePattern('^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}$')]
        [string] $AppId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'BySpObjectId')]
        [Alias('Id','ObjectId','ServicePrincipalId')]
        [ValidatePattern('^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}$')]
        [string] $SpObjectId,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]] $Role,

        [Parameter()]
        [bool] $Enabled,

        [Parameter()]
        [ValidateSet('All','ServicePrincipal','User','SecurityGroup','RoleGroup','RoleAssignmentPolicy','ForeignSecurityPrincipal','LinkedRoleGroup','Computer')]
        [string] $RoleAssigneeType = 'ServicePrincipal'
    )

    process {
        # --- Resolve the service principal when an application filter is requested.
        $sp = $null
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'BySpObjectId' {
                    $sp = Get-MgServicePrincipal -ServicePrincipalId $SpObjectId -ErrorAction Stop
                }
                'ByAppId' {
                    $matchesRes = @(Get-MgServicePrincipal -Filter "appId eq `'$AppId`'" -ErrorAction Stop)
                    if ($matchesRes.Count -eq 0) { throw "No service principal found for AppId '$AppId'." }
                    $sp = $matchesRes[0]
                }
                'ByName' {
                    $matchesRes = @(Get-MgServicePrincipal -Filter "displayName eq `'$RegisteredAppName`'" -ErrorAction Stop)
                    if ($matchesRes.Count -eq 0) { throw "No service principal found for displayName '$RegisteredAppName'." }
                    if ($matchesRes.Count -gt 1) {
                        $ids = ($matchesRes | Select-Object -First 10 -ExpandProperty Id) -join ', '
                        throw "Ambiguous displayName '$RegisteredAppName' matched $($matchesRes.Count) service principals. Re-run with -AppId or -SpObjectId. Example SP objectIds: $ids"
                    }
                    $sp = $matchesRes[0]
                }
            }
        }
        catch {
            Write-Error -Message $_.Exception.Message
            return
        }

        # --- Retrieve assignments, filtered by role where possible.
        $rolesNormalized = if ($PSBoundParameters.ContainsKey('Role')) { @($Role | ForEach-Object { ConvertTo-AppRole $_ }) } else { @() }

        $assignments =
            if ($rolesNormalized.Count -gt 0) {
                # Query per requested role so EXO does the role filtering for us.
                foreach ($r in $rolesNormalized) {
                    Get-ManagementRoleAssignment -Role $r -ErrorAction SilentlyContinue
                }
            }
            else {
                # No role filter: get everything and keep only application roles.
                Get-ManagementRoleAssignment -ErrorAction SilentlyContinue |
                    Where-Object { $_.Role -like 'Application *' }
            }

        # --- Optional enabled-state filter.
        if ($PSBoundParameters.ContainsKey('Enabled')) {
            $assignments = $assignments | Where-Object { $_.Enabled -eq $Enabled }
        }

        # --- Keep only the recipient scopes this module creates/uses.
        $assignments = $assignments | Where-Object {
            $_.RecipientWriteScope -in @('Group', 'CustomRecipientScope')
        }

        # --- Assignee-type filter (default: ServicePrincipal; 'All' returns every type).
        if ($RoleAssigneeType -ne 'All') {
            $assignments = $assignments | Where-Object { $_.RoleAssigneeType -eq $RoleAssigneeType }
        }

        # --- Optional application filter (client-side: Get-ManagementRoleAssignment has no -App).
        if ($sp) {
            $needles = @($sp.DisplayName, ("{0}_SP" -f $sp.DisplayName), $sp.AppId, $sp.Id) | Where-Object { $_ }
            $assignments = $assignments | Where-Object {
                $assignee = [string]$_.RoleAssigneeName
                $name     = [string]$_.Name
                $hit = $false
                foreach ($n in $needles) {
                    if (($assignee -and $assignee -like "*$n*") -or ($name -and $name -like "*$n*")) { $hit = $true; break }
                }
                $hit
            }
        }

        foreach ($a in $assignments) {
            if (-not $a) { continue }
            [pscustomobject][ordered]@{
                Name             = $a.Name
                Role             = $a.Role
                RoleAssigneeName = $a.RoleAssigneeName
                RoleAssigneeType = $a.RoleAssigneeType
                Scope            = $a.CustomRecipientWriteScope
                RecipientScope   = $a.RecipientWriteScope
                Enabled          = $a.Enabled
                Guid             = $a.Guid
                Identity         = $a.Identity
            }
        }
    }
}
