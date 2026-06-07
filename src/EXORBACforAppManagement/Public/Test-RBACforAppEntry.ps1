<#
.SYNOPSIS
Validates that a registered application has all the Exchange Online RBAC components that
New-RBACforAppEntry creates.

.DESCRIPTION
Test-RBACforAppEntry resolves an Entra application / service principal (by display name, AppId, or
service principal object id) and checks that every component New-RBACforAppEntry provisions is in
place:

  1. the Entra service principal is resolvable (Microsoft Graph),
  2. the scoped Unified Group "{GroupPrefix}-{DisplayName}" (sanitized via the same Get-SafeName
     rule) exists,
  3. the Exchange Online service principal pointer "{DisplayName}_SP" exists, and
  4. one Exchange Online management role assignment exists per requested role, named the same way
     New-RBACforAppEntry names them ("{ShortRoleToken}-{DisplayName}") and bound to the expected
     role.

When -Members is supplied, the requested recipients are also verified against the Unified Group's
membership. The function is read-only: it makes no changes and does not support -WhatIf.

The defaults for -Role and -GroupPrefix mirror New-RBACforAppEntry, so a plain
"Test-RBACforAppEntry -RegisteredAppName <app>" validates the default creation.

.PARAMETER RegisteredAppName
Display name of the registered application or service principal. Default parameter set; must resolve
to exactly one service principal.

.PARAMETER AppId
Application (client) id of the registered application. GUID-validated.

.PARAMETER SpObjectId
Object id of the target service principal. GUID-validated.

.PARAMETER Role
Exchange Online application roles expected to be assigned. Short names such as Mail.Send are
normalized to Application Mail.Send. Defaults to 'Application Mail.Send' (matching New-RBACforAppEntry).

.PARAMETER Members
Optional recipients expected to be members of the Unified Group scope. When supplied, each is
resolved through Get-Recipient and checked against the group's membership. Omit to skip the
membership check.

.PARAMETER GroupPrefix
Prefix used when building the Unified Group name. Defaults to 'Um365RAo1' (matching
New-RBACforAppEntry).

.EXAMPLE
Test-RBACforAppEntry -RegisteredAppName 'Contoso Mail App'

Validates the default Application Mail.Send setup for the resolved application and returns a summary
with an IsValid flag.

.EXAMPLE
Test-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' -Role 'Mail.Send','Calendars.Read' -Members 'shared@contoso.com'

Checks the service principal, Unified Group, Exchange Online service principal, both role
assignments, and that shared@contoso.com is a group member.

.EXAMPLE
New-RBACforAppEntry -RegisteredAppName 'Contoso' -WhatIf; Test-RBACforAppEntry -RegisteredAppName 'Contoso'

Reports which components are still missing before/after a run.

.OUTPUTS
PSCustomObject

A summary object with the resolved identity, per-component existence flags
(ServicePrincipalExists, UnifiedGroupExists, ExoServicePrincipalExists), the expected/found/missing
role assignments, optional membership results, an overall IsValid flag, a Missing list, and any
Warnings/Errors.

.NOTES
Requires a connected Microsoft Graph session (Get-MgServicePrincipal, Get-MgContext) and a connected
Exchange Online session (Get-UnifiedGroup, Get-ServicePrincipal, Get-ManagementRoleAssignment, and,
when -Members is supplied, Get-Recipient and Get-UnifiedGroupLinks). Read-only companion to
New-RBACforAppEntry.
#>
function Test-RBACforAppEntry {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
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

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Role = @('Application Mail.Send'),

        [Parameter(Position = 2)]
        [string[]] $Members,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $GroupPrefix = 'Um365RAo1'
    )

    begin {
        $shortRoleMap = Get-AppRoleMap

        $tenantid = $null
        try { $tenantid = Get-MgContext | Select-Object -ExpandProperty TenantId }
        catch { Write-Verbose -Message "Could not read tenant id from Get-MgContext: $($_.Exception.Message)" }
    }

    process {
        $result = [ordered]@{
            ParameterSet            = $PSCmdlet.ParameterSetName
            IdentityInput           = $RegisteredAppName
            ResolvedDisplay         = $null
            AppId                   = $null
            SpObjectId              = $null
            TenantId                = $tenantid
            ServicePrincipalExists  = $false
            UnifiedGroupName        = $null
            UnifiedGroupExists      = $false
            ExoServicePrincipalName = $null
            ExoServicePrincipalExists = $false
            RolesExpected           = @()
            RoleAssignmentsExpected = @()
            RoleAssignmentsFound    = @()
            RoleAssignmentsMissing  = @()
            MembersExpected         = @()
            MembersPresent          = @()
            MembersMissing          = @()
            IsValid                 = $false
            Missing                 = @()
            Warnings                = @()
            Errors                  = @()
        }

        try {
            # --- Resolve the service principal depending on parameter set.
            $sp = $null
            switch ($PSCmdlet.ParameterSetName) {
                'BySpObjectId' {
                    $result.IdentityInput = $SpObjectId
                    $sp = Get-MgServicePrincipal -ServicePrincipalId $SpObjectId -ErrorAction Stop
                }
                'ByAppId' {
                    $result.IdentityInput = $AppId
                    $matchesRes = @(Get-MgServicePrincipal -Filter "appId eq `'$AppId`'" -ErrorAction Stop)
                    if ($matchesRes.Count -eq 0) { throw "No service principal found for AppId '$AppId'." }
                    if ($matchesRes.Count -gt 1) { throw "Unexpected: multiple service principals for AppId '$AppId'." }
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

            $result.ServicePrincipalExists = $true
            $result.ResolvedDisplay        = $sp.DisplayName
            $result.AppId                  = $sp.AppId
            $result.SpObjectId             = $sp.Id

            # --- Unified Group existence (same name rule as New-RBACforAppEntry).
            $umGroupName = Get-SafeName -s ("{0}-{1}" -f $GroupPrefix, $sp.DisplayName)
            $result.UnifiedGroupName = $umGroupName
            $group = Get-UnifiedGroup -Identity $umGroupName -ErrorAction SilentlyContinue
            if ($group) {
                $result.UnifiedGroupExists = $true
            }
            else {
                $result.Missing += "Unified Group '$umGroupName'"
            }

            # --- Exchange Online service principal pointer existence (matched by AppId, then name).
            $exoSpDisplay = "{0}_SP" -f $sp.DisplayName
            $result.ExoServicePrincipalName = $exoSpDisplay
            $exoSp = @(Get-ServicePrincipal -ErrorAction SilentlyContinue) |
                Where-Object { $_ -and (($_.AppId -eq $sp.AppId) -or ($_.DisplayName -eq $exoSpDisplay)) } |
                Select-Object -First 1
            if ($exoSp) {
                $result.ExoServicePrincipalExists = $true
            }
            else {
                $result.Missing += "Exchange Online service principal '$exoSpDisplay'"
            }

            # --- Role assignment existence, by the deterministic name New-RBACforAppEntry builds.
            $rolesNormalized = foreach ($r in @($Role)) { Get-NormalizeRole $r }
            $result.RolesExpected = @($rolesNormalized)

            foreach ($roleItem in $rolesNormalized) {
                $shortRoleName = $shortRoleMap[$roleItem]
                if (-not $shortRoleName) {
                    $result.Warnings += "Role '$roleItem' is not a recognized application role; skipping its assignment check."
                    continue
                }

                $expectedName = Get-SafeName -s ("{0}-{1}" -f $shortRoleName, $sp.DisplayName) -max 63
                $result.RoleAssignmentsExpected += $expectedName

                $assignment = Get-ManagementRoleAssignment -Identity $expectedName -ErrorAction SilentlyContinue
                if ($assignment -and ([string]$assignment.Role -eq $roleItem)) {
                    $result.RoleAssignmentsFound += $expectedName
                }
                else {
                    $result.RoleAssignmentsMissing += $expectedName
                    $result.Missing += "Role assignment '$expectedName' ($roleItem)"
                }
            }

            # --- Optional membership check.
            if ($PSBoundParameters.ContainsKey('Members')) {
                $requested = @($Members | Where-Object { $_ })
                $result.MembersExpected = $requested

                if ($result.UnifiedGroupExists) {
                    $links = @(Get-UnifiedGroupLinks -Identity $umGroupName -LinkType Members -ErrorAction SilentlyContinue)
                    $linkAddresses = @($links | ForEach-Object { [string]$_.PrimarySmtpAddress; [string]$_.Name } | Where-Object { $_ })

                    foreach ($member in $requested) {
                        $rec = Get-Recipient -Identity $member -ErrorAction SilentlyContinue
                        $needle = if ($rec) { [string]$rec.PrimarySmtpAddress } else { [string]$member }
                        if ($linkAddresses -contains $needle -or ($rec -and ($linkAddresses -contains [string]$rec.Name))) {
                            $result.MembersPresent += $needle
                        }
                        else {
                            $result.MembersMissing += $needle
                            $result.Missing += "Group member '$needle'"
                        }
                    }
                }
                else {
                    foreach ($member in $requested) {
                        $result.MembersMissing += [string]$member
                        $result.Missing += "Group member '$member' (group missing)"
                    }
                }
            }

            $result.IsValid = $result.ServicePrincipalExists -and
                              $result.UnifiedGroupExists -and
                              $result.ExoServicePrincipalExists -and
                              ($result.RoleAssignmentsMissing.Count -eq 0) -and
                              ($result.MembersMissing.Count -eq 0)

            [pscustomobject]$result
        }
        catch {
            $result.Errors += $_.Exception.Message
            [pscustomobject]$result
        }
    }
}
