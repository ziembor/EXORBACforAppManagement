<#
.SYNOPSIS
Reconciles a registered application's Exchange Online RBAC scoping to the desired state: repairs
missing components, tops up the Unified Group membership, and can move the role assignments to a
different scoping group.

.DESCRIPTION
Set-RBACforAppEntry is the "make it so" companion to Test-RBACforAppEntry (which only reports) and
New-RBACforAppEntry (which creates everything from scratch). It resolves an Entra application /
service principal (by display name, AppId, or service principal object id) and brings the components
New-RBACforAppEntry provisions into the desired state, changing only what is needed:

  1. the scoped Unified Group ("{GroupPrefix}-{DisplayName}", sanitized via Get-SafeName) is created
     if it is missing (delegated to New-RBACforAppUnifiedGroup),
  2. the Exchange Online service principal pointer ("{DisplayName}_SP") is created if it is missing
     (delegated to Register-EXOServicePrincipal),
  3. the requested -Members are added to the group if they are not already members (additive only -
     existing members are never removed), and
  4. one Exchange Online management role assignment exists per requested role, named the same way
     New-RBACforAppEntry names them ("{ShortRoleToken}-{DisplayName}") and scoped to the target group.
     A missing assignment is created; an assignment that exists but is scoped to a different group is
     re-scoped (removed and recreated with the same name) - this is how the scoping group is changed.

Changing the scoping group: supply -NewGroupName (an explicit group name) or -NewGroupPrefix (builds
"{NewGroupPrefix}-{DisplayName}"). The target group is ensured to exist and the role assignments are
re-scoped onto it. The old group is left in place (use Remove-RBACforAppEntry to tear it down once it
is no longer in use); members are not migrated.

The function supports -WhatIf and -Confirm through SupportsShouldProcess (ConfirmImpact High), so each
change is confirmed interactively unless -Confirm:$false / -Force-style suppression is used. Under
-WhatIf no changes are made and IsValid reflects the actual (unchanged) state.

.PARAMETER RegisteredAppName
Display name of the registered application or service principal. Default parameter set; must resolve
to exactly one service principal.

.PARAMETER AppId
Application (client) id of the registered application. GUID-validated.

.PARAMETER SpObjectId
Object id of the target service principal. GUID-validated.

.PARAMETER Role
Exchange Online application roles to ensure are assigned. Short names such as Mail.Send are normalized
to Application Mail.Send. Defaults to 'Application Mail.Send' (matching New-RBACforAppEntry).

.PARAMETER Members
Optional recipients to ensure are members of the scoped Unified Group. Each is resolved through
Get-Recipient and added if absent. Membership is additive: members already present are left as-is and
nothing is removed.

.PARAMETER ManagedBy
Recipient assigned as the Unified Group owner when the group must be created. Defaults to
'GraphAPI-Dummy-owner' (matching New-RBACforAppEntry).

.PARAMETER GroupPrefix
Prefix used when building the current Unified Group name. Defaults to 'Um365RAo1' (matching
New-RBACforAppEntry).

.PARAMETER BootstrapMember
Optional initial member passed to New-RBACforAppUnifiedGroup when the group must be created. Defaults
to 'GraphAPI-Dummy'.

.PARAMETER NewGroupPrefix
Optional. Switch the role assignments to a group named "{NewGroupPrefix}-{DisplayName}" (sanitized).
Ignored when -NewGroupName is supplied.

.PARAMETER NewGroupName
Optional. Switch the role assignments to this explicit group name (sanitized via Get-SafeName). Takes
precedence over -NewGroupPrefix.

.EXAMPLE
Set-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -WhatIf -Verbose

Shows which missing components would be (re)created for the default Application Mail.Send setup,
without making changes.

.EXAMPLE
Set-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -Members 'shared@contoso.com'

Ensures every component exists and adds shared@contoso.com to the scoped group if it is not already a
member.

.EXAMPLE
Set-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' -NewGroupPrefix 'Um365Prod'

Re-scopes the application's role assignments onto the 'Um365Prod-...' group (creating it if needed).

.OUTPUTS
PSCustomObject

A summary object with the resolved identity, the current and target group names, a GroupChanged flag,
which components were created, the requested/added members, and the role assignments partitioned into
created / re-scoped / unchanged, an overall IsValid flag, and any Warnings/Errors.

.NOTES
Requires a connected Microsoft Graph session (Get-MgServicePrincipal, Get-MgContext) and a connected
Exchange Online session (Get-UnifiedGroup, Get-UnifiedGroupLinks, Get-ServicePrincipal, Get-Recipient,
Get-ManagementRoleAssignment, New-ManagementRoleAssignment, Remove-ManagementRoleAssignment, plus the
cmdlets used by the delegated New-RBACforAppUnifiedGroup / Register-EXOServicePrincipal). Reconcile
companion to Test-RBACforAppEntry and New-RBACforAppEntry.
#>
function Set-RBACforAppEntry {
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

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Role = @('Application Mail.Send'),

        [Parameter(Position = 2)]
        [string[]] $Members,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $ManagedBy = 'GraphAPI-Dummy-owner',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $GroupPrefix = 'Um365RAo1',

        [Parameter()]
        [string] $BootstrapMember = 'GraphAPI-Dummy',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $NewGroupPrefix,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $NewGroupName
    )

    begin {
        $shortRoleMap = Get-AppRoleMap

        $tenantid = $null
        try { $tenantid = Get-MgContext | Select-Object -ExpandProperty TenantId }
        catch { Write-Verbose -Message "Could not read tenant id from Get-MgContext: $($_.Exception.Message)" }
    }

    process {
        $result = [ordered]@{
            ParameterSet              = $PSCmdlet.ParameterSetName
            IdentityInput             = $RegisteredAppName
            ResolvedDisplay           = $null
            AppId                     = $null
            SpObjectId                = $null
            TenantId                  = $tenantid
            CurrentGroupName          = $null
            TargetGroupName           = $null
            GroupChanged              = $false
            UnifiedGroupExisted       = $false
            UnifiedGroupCreated       = $false
            ExoServicePrincipalName   = $null
            ExoServicePrincipalExisted = $false
            ExoServicePrincipalCreated = $false
            MembersRequested          = @()
            MembersAdded              = @()
            MembersAlreadyPresent     = @()
            FilteredMembers           = @()
            RolesNormalized           = @()
            RoleAssignmentsCreated    = @()
            RoleAssignmentsRescoped   = @()
            RoleAssignmentsUnchanged  = @()
            IsValid                   = $false
            Warnings                  = @()
            Errors                    = @()
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

            # --- Determine the current and target Unified Group names (same name rule as New-RBACforAppEntry).
            $currentGroup = Get-SafeName -s ("{0}-{1}" -f $GroupPrefix, $sp.DisplayName)
            if ($PSBoundParameters.ContainsKey('NewGroupName')) {
                $targetGroup = Get-SafeName -s $NewGroupName
            }
            elseif ($PSBoundParameters.ContainsKey('NewGroupPrefix')) {
                $targetGroup = Get-SafeName -s ("{0}-{1}" -f $NewGroupPrefix, $sp.DisplayName)
            }
            else {
                $targetGroup = $currentGroup
            }
            $result.CurrentGroupName = $currentGroup
            $result.TargetGroupName  = $targetGroup
            $result.GroupChanged     = ($targetGroup -ne $currentGroup)

            # --- Ensure the target Unified Group exists (delegated to New-RBACforAppUnifiedGroup).
            $group = Get-UnifiedGroup -Identity $targetGroup -ErrorAction SilentlyContinue
            $result.UnifiedGroupExisted = [bool]$group
            if (-not $group) {
                if ($PSCmdlet.ShouldProcess($targetGroup, 'Create Unified Group')) {
                    $ugResult = New-RBACforAppUnifiedGroup -Name $targetGroup -ManagedBy $ManagedBy -BootstrapMember $BootstrapMember -WarningVariable ugWarnings
                    foreach ($w in $ugWarnings) {
                        if ([string]$w.Message -like '*already exists*') { $result.Warnings += [string]$w.Message }
                    }
                    if ($ugResult) {
                        $result.UnifiedGroupCreated = $true
                        $group = $ugResult.Group
                    }
                }
            }

            # --- Ensure the Exchange Online service principal pointer (matched by AppId, then name).
            $exoSpDisplay = "{0}_SP" -f $sp.DisplayName
            $result.ExoServicePrincipalName = $exoSpDisplay
            $exoSp = @(Get-ServicePrincipal -ErrorAction SilentlyContinue) |
                Where-Object { $_ -and (($_.AppId -eq $sp.AppId) -or ($_.DisplayName -eq $exoSpDisplay)) } |
                Select-Object -First 1
            if ($exoSp) {
                $result.ExoServicePrincipalExisted = $true
            }
            elseif ($PSCmdlet.ShouldProcess($exoSpDisplay, 'Create Exchange Online service principal')) {
                $null = Register-EXOServicePrincipal -AppId $sp.AppId -ObjectId $sp.Id -DisplayName $exoSpDisplay
                $result.ExoServicePrincipalCreated = $true
            }

            # --- Members (additive): add any requested member not already in the target group.
            if ($PSBoundParameters.ContainsKey('Members')) {
                $requested = @($Members | Where-Object { $_ })
                $result.MembersRequested = $requested

                $currentUserUpn = $null
                try { $currentUserUpn = (Get-MgContext -ErrorAction Stop).Account }
                catch {
                    $result.Warnings += "Could not retrieve current connection user via Get-MgContext; connection user filtering will be skipped. Error: $($_.Exception.Message)"
                }

                $links = @(Get-UnifiedGroupLinks -Identity $targetGroup -LinkType Members -ErrorAction SilentlyContinue)
                $linkAddresses = @($links | ForEach-Object { [string]$_.PrimarySmtpAddress; [string]$_.Name } | Where-Object { $_ })

                foreach ($member in $requested) {
                    if ($currentUserUpn -and ($member -ieq $currentUserUpn)) {
                        $result.FilteredMembers += [string]$member
                        $filterWarning = "Current connection user '$currentUserUpn' was found in the members list and has been filtered out."
                        $result.Warnings += $filterWarning
                        Write-Warning -Message $filterWarning
                        continue
                    }

                    $rec = Get-Recipient -Identity $member -ErrorAction SilentlyContinue
                    if (-not $rec) {
                        $result.Warnings += "Recipient not found for '$member' (skipped)."
                        continue
                    }
                    $needle = [string]$rec.PrimarySmtpAddress

                    if (($linkAddresses -contains $needle) -or ($linkAddresses -contains [string]$rec.Name)) {
                        $result.MembersAlreadyPresent += $needle
                        continue
                    }

                    if ($PSCmdlet.ShouldProcess("UnifiedGroup $targetGroup", "Add member $needle")) {
                        Add-UnifiedGroupLinks -Identity $targetGroup -LinkType Members -Links $needle -ErrorAction Stop
                        $result.MembersAdded += $needle
                    }
                }
            }

            # --- Role assignments: ensure one per role, scoped to the target group.
            $rolesNormalized = foreach ($r in @($Role)) { Get-NormalizeRole $r }
            $result.RolesNormalized = @($rolesNormalized)

            $rolesAllSatisfied = $true
            foreach ($roleItem in $rolesNormalized) {
                $shortRoleName = $shortRoleMap[$roleItem]
                if (-not $shortRoleName) {
                    $result.Warnings += "Role '$roleItem' is not a recognized application role; skipping its assignment."
                    $rolesAllSatisfied = $false
                    continue
                }

                $rbacName = Get-SafeName -s ("{0}-{1}" -f $shortRoleName, $sp.DisplayName) -max 63

                try {
                    $existing = Get-ManagementRoleAssignment -Identity $rbacName -ErrorAction SilentlyContinue
                    $scopedToTarget = $existing -and
                        ([string]$existing.RecipientWriteScope -in @('Group','CustomRecipientScope')) -and
                        ([string]$existing.CustomRecipientWriteScope -eq $targetGroup)

                    if ($existing -and $scopedToTarget) {
                        # Already in desired state.
                        $result.RoleAssignmentsUnchanged += $rbacName
                    }
                    elseif ($existing) {
                        # Exists but scoped elsewhere (repair / group change): re-scope by recreating.
                        $action = "Re-scope role '$roleItem' for App '$($sp.DisplayName)' to '$targetGroup'"
                        if ($PSCmdlet.ShouldProcess($rbacName, $action)) {
                            Remove-ManagementRoleAssignment -Identity $rbacName -Confirm:$false -ErrorAction Stop
                            $null = New-ManagementRoleAssignment -App $sp.Id -Role $roleItem -RecipientGroupScope $targetGroup -Name $rbacName -ErrorAction Stop
                            $result.RoleAssignmentsRescoped += $rbacName
                        }
                        else {
                            $rolesAllSatisfied = $false
                        }
                    }
                    else {
                        # Missing: create.
                        $action = "Assign '$roleItem' to App '$($sp.DisplayName)' scoped to '$targetGroup'"
                        if ($PSCmdlet.ShouldProcess($rbacName, $action)) {
                            $null = New-ManagementRoleAssignment -App $sp.Id -Role $roleItem -RecipientGroupScope $targetGroup -Name $rbacName -ErrorAction Stop
                            $result.RoleAssignmentsCreated += $rbacName
                        }
                        else {
                            $rolesAllSatisfied = $false
                        }
                    }
                }
                catch {
                    $result.Errors += "Failed to reconcile role assignment '$rbacName' ($roleItem): $($_.Exception.Message)"
                    $rolesAllSatisfied = $false
                }
            }

            # IsValid: every component is present after this run (true unchanged state, or actually applied).
            $groupOk = $result.UnifiedGroupExisted -or $result.UnifiedGroupCreated
            $exoOk   = $result.ExoServicePrincipalExisted -or $result.ExoServicePrincipalCreated
            $result.IsValid = ($result.Errors.Count -eq 0) -and $groupOk -and $exoOk -and $rolesAllSatisfied

            [pscustomobject]$result
        }
        catch {
            $result.Errors += $_.Exception.Message
            [pscustomobject]$result
        }
    }
}
