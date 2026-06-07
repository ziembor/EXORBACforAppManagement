<#
.SYNOPSIS
Safely removes the Exchange Online RBAC scoping that New-RBACforAppEntry creates for an Entra
application.

.DESCRIPTION
Remove-RBACforAppEntry is the teardown counterpart to New-RBACforAppEntry. It resolves an Entra
application / service principal (by display name, AppId, or service principal object id), derives the
scoped Unified Group name the same way New-RBACforAppEntry does ("{GroupPrefix}-{DisplayName}",
sanitized via Get-SafeName), and removes:

  1. the Exchange Online management role assignments scoped to that group that belong to this
     application, and
  2. the scoped Unified Group itself.

Before removing anything the function runs two safety checks and ABORTS (removing nothing) if either
fails:

  * Foreign role assignments - if any management role assignment scoped to the group resolves to a
    DIFFERENT identity than this service principal, the group is still in use and removal is refused.
  * Real members - if the group has any member other than the bootstrap placeholder
    (-BootstrapMember, default 'GraphAPI-Dummy'), the group is still in use and removal is refused.

The Exchange Online service principal pointer ("{DisplayName}_SP") is intentionally left in place
because it may be shared by other scoping. There is no -Role parameter: the safety question is about
the group as a whole, so the teardown operates at group granularity and removes all of this app's
assignments scoped to the group or nothing at all.

The function supports -WhatIf and -Confirm through SupportsShouldProcess; under -WhatIf no removals
are performed and IsRemoved is reported as false.

.PARAMETER RegisteredAppName
Display name of the registered application or service principal. Default parameter set; must resolve
to exactly one service principal.

.PARAMETER AppId
Application (client) id of the registered application. GUID-validated.

.PARAMETER SpObjectId
Object id of the target service principal. GUID-validated.

.PARAMETER GroupPrefix
Prefix used when building the Unified Group name. Defaults to 'Um365RAo1' (matching
New-RBACforAppEntry).

.PARAMETER BootstrapMember
Bootstrap placeholder member to ignore when deciding whether the group has real members. Defaults to
'GraphAPI-Dummy' (matching New-RBACforAppEntry).

.EXAMPLE
Remove-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -WhatIf -Verbose

Shows the role assignments and Unified Group that would be removed for the resolved application,
without making changes.

.EXAMPLE
Remove-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555'

Removes this application's role assignments and the scoped Unified Group, but only if no foreign
assignments and no real members are present.

.OUTPUTS
PSCustomObject

A summary object with the resolved identity, the Unified Group name and whether it existed, the
assignments scoped to the group partitioned into own/foreign, the real members found, what was
removed (AssignmentsRemoved, GroupRemoved), an overall IsRemoved flag, a Reason when removal was
refused or skipped, and any Warnings/Errors.

.NOTES
Requires a connected Microsoft Graph session (Get-MgServicePrincipal, Get-MgContext) and a connected
Exchange Online session (Get-UnifiedGroup, Get-UnifiedGroupLinks, Get-ManagementRoleAssignment,
Remove-ManagementRoleAssignment, Remove-UnifiedGroup). Inverse of New-RBACforAppEntry; the safe
companion to Test-RBACforAppEntry.
#>
function Remove-RBACforAppEntry {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'ByName')]
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
        [string] $GroupPrefix = 'Um365RAo1',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $BootstrapMember = 'GraphAPI-Dummy'
    )

    begin {
        $tenantid = $null
        try { $tenantid = Get-MgContext | Select-Object -ExpandProperty TenantId }
        catch { Write-Verbose -Message "Could not read tenant id from Get-MgContext: $($_.Exception.Message)" }
    }

    process {
        $result = [ordered]@{
            ParameterSet        = $PSCmdlet.ParameterSetName
            IdentityInput       = $RegisteredAppName
            ResolvedDisplay     = $null
            AppId               = $null
            SpObjectId          = $null
            TenantId            = $tenantid
            UnifiedGroupName    = $null
            UnifiedGroupExisted = $false
            AssignmentsScoped   = @()
            OwnAssignments      = @()
            ForeignAssignments  = @()
            RealMembers         = @()
            AssignmentsRemoved  = @()
            GroupRemoved        = $false
            IsRemoved           = $false
            Reason              = $null
            Warnings            = @()
            Errors              = @()
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

            $result.ResolvedDisplay = $sp.DisplayName
            $result.AppId           = $sp.AppId
            $result.SpObjectId      = $sp.Id

            # --- Unified Group name (same rule as New-RBACforAppEntry).
            $umGroupName = Get-SafeName -s ("{0}-{1}" -f $GroupPrefix, $sp.DisplayName)
            $result.UnifiedGroupName = $umGroupName

            $group = Get-UnifiedGroup -Identity $umGroupName -ErrorAction SilentlyContinue
            $result.UnifiedGroupExisted = [bool]$group

            # --- Assignments scoped to the group (client-side filter: no -App on the EXO cmdlet).
            $scoped = @(Get-ManagementRoleAssignment -ErrorAction SilentlyContinue) | Where-Object {
                $_ -and
                ([string]$_.RecipientWriteScope -in @('Group','CustomRecipientScope')) -and
                ([string]$_.CustomRecipientWriteScope -eq $umGroupName)
            }
            $result.AssignmentsScoped = @($scoped | ForEach-Object { [string]$_.Name })

            # --- Partition own vs foreign by assignee (same needle set as Get-RBACforAppEntry).
            $needles = @($sp.DisplayName, ("{0}_SP" -f $sp.DisplayName), $sp.AppId, $sp.Id) | Where-Object { $_ }
            $own = @()
            $foreign = @()
            foreach ($a in $scoped) {
                $assignee = [string]$a.RoleAssigneeName
                $isOwn = $false
                foreach ($n in $needles) {
                    if ($assignee -and $assignee -like "*$n*") { $isOwn = $true; break }
                }
                if ($isOwn) { $own += $a } else { $foreign += $a }
            }
            $result.OwnAssignments     = @($own | ForEach-Object { [string]$_.Name })
            $result.ForeignAssignments = @($foreign | ForEach-Object { [string]$_.Name })

            # --- Real members (ignore the bootstrap placeholder).
            $realMembers = @()
            if ($group) {
                $links = @(Get-UnifiedGroupLinks -Identity $umGroupName -LinkType Members -ErrorAction SilentlyContinue)
                foreach ($l in $links) {
                    if (-not $l) { continue }
                    $smtp = [string]$l.PrimarySmtpAddress
                    $name = [string]$l.Name
                    if (($smtp -and $smtp -eq $BootstrapMember) -or ($name -and $name -eq $BootstrapMember)) { continue }
                    $realMembers += if ($smtp) { $smtp } else { $name }
                }
            }
            $result.RealMembers = @($realMembers)

            # --- Safety gate: abort (remove nothing) when the group is still in use.
            if ($foreign.Count -gt 0 -or $realMembers.Count -gt 0) {
                $parts = @()
                if ($foreign.Count -gt 0)     { $parts += "$($foreign.Count) foreign role assignment(s) scoped to the group" }
                if ($realMembers.Count -gt 0) { $parts += "$($realMembers.Count) member(s) beyond the '$BootstrapMember' placeholder" }
                $result.Reason = "Refusing to remove '$umGroupName': it is still in use ($($parts -join '; '))."
                $result.IsRemoved = $false
                Write-Warning -Message $result.Reason
                return [pscustomobject]$result
            }

            # --- Safe path: remove this app's assignments, then the group.
            $removedAny = $false
            foreach ($a in $own) {
                $name = [string]$a.Name
                if ($PSCmdlet.ShouldProcess($name, 'Remove-ManagementRoleAssignment')) {
                    try {
                        Remove-ManagementRoleAssignment -Identity $name -Confirm:$false -ErrorAction Stop
                        $result.AssignmentsRemoved += $name
                        $removedAny = $true
                    }
                    catch {
                        $result.Errors += "Failed to remove role assignment '$name': $($_.Exception.Message)"
                    }
                }
            }

            if (-not $group) {
                $result.Warnings += "Unified Group '$umGroupName' did not exist; only role assignments (if any) were processed."
            }
            elseif ($PSCmdlet.ShouldProcess($umGroupName, 'Remove-UnifiedGroup')) {
                try {
                    Remove-UnifiedGroup -Identity $umGroupName -Confirm:$false -ErrorAction Stop
                    $result.GroupRemoved = $true
                    $removedAny = $true
                }
                catch {
                    $result.Errors += "Failed to remove Unified Group '$umGroupName': $($_.Exception.Message)"
                }
            }

            # IsRemoved is true only when the group is gone (or never existed) and no errors occurred.
            $result.IsRemoved = ($result.Errors.Count -eq 0) -and
                                ($result.GroupRemoved -or -not $group) -and
                                $removedAny
            if (-not $result.IsRemoved -and -not $result.Reason -and $WhatIfPreference) {
                $result.Reason = 'WhatIf: no changes were made.'
            }

            [pscustomobject]$result
        }
        catch {
            $result.Errors += $_.Exception.Message
            [pscustomobject]$result
        }
    }
}
