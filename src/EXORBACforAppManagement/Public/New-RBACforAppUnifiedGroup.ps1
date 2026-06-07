<#
.SYNOPSIS
Ensures the scoped Microsoft 365 Unified Group used by New-RBACforAppEntry exists and is configured.

.DESCRIPTION
New-RBACforAppUnifiedGroup creates a private, hidden Unified Group (when it does not already exist)
to act as the recipient scope for Exchange Online application RBAC. The group is first created with
the smallest set of essential attributes (DisplayName/Name/Alias, AccessType Private, the
creation-only HiddenGroupMembershipEnabled, and the owner/bootstrap member); the remaining settings
(member edit, auto-subscribe, calendar subscribe, language, subscription, address-list visibility,
and connectors) are then applied via Set-UnifiedGroup. If the group already exists it is left in
place and a warning is emitted. A summary object describing the resolved group is returned.

The function supports -WhatIf and -Confirm through SupportsShouldProcess.

.PARAMETER Name
Name and Alias of the Unified Group. Expected to already be a safe value (<= 63 chars,
alphanumeric/dash); callers such as New-RBACforAppEntry sanitize it with Get-SafeName first.

.PARAMETER DisplayName
Display name for the group. Defaults to "{Name} - RBAC for APP".

.PARAMETER ManagedBy
Recipient assigned as the group owner. Defaults to the GraphAPI-Dummy-owner placeholder.

.PARAMETER BootstrapMember
Optional initial member passed during group creation. Defaults to the GraphAPI-Dummy placeholder.

.EXAMPLE
New-RBACforAppUnifiedGroup -Name 'Um365RAo1-ContosoMailApp' -WhatIf -Verbose

Shows the planned Unified Group creation without making changes.

.OUTPUTS
PSCustomObject

A summary object describing the group: Name, DisplayName, OwnerRequested (the -ManagedBy input),
OwnerAdded (the owner actually applied/in place), AlreadyExisted, and Group (the underlying Exchange
Online Unified Group object, existing or newly created).

.NOTES
Requires a connected Exchange Online session (Get-UnifiedGroup, New-UnifiedGroup, Set-UnifiedGroup,
Get-Recipient) and a connected Microsoft Graph session for the debug calling-context snapshot.
Companion to New-RBACforAppEntry.
#>
function New-RBACforAppUnifiedGroup {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([object])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string] $Name,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $DisplayName,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $ManagedBy = 'GraphAPI-Dummy-owner',

        [Parameter()]
        [string] $BootstrapMember = 'GraphAPI-Dummy'
    )

    process {
        if (-not $DisplayName) { $DisplayName = '{0} - RBAC for APP' -f $Name }

        Write-Verbose -Message ("Checking Unified Group '{0}'." -f $Name)
        $existingGroup = Get-UnifiedGroup -Identity $Name -ErrorAction SilentlyContinue
        if ($existingGroup) {
            $existingOwner = ($existingGroup.ManagedBy | Where-Object { $_ }) -join ', '
            Write-Warning -Message ("UnifiedGroup '{0}' already exists; will only add missing members / assignments." -f $Name)
            Write-Verbose -Message ("Unified Group '{0}' already exists; skipping creation." -f $Name)
            Write-Debug -Message ("Existing Unified Group details: DisplayName='{0}'; Identity='{1}'; ManagedBy='{2}'" -f $existingGroup.DisplayName, $existingGroup.Identity, $existingOwner)
            return [pscustomobject]@{
                Name           = $Name
                DisplayName    = $existingGroup.DisplayName
                OwnerRequested = $ManagedBy
                OwnerAdded     = $existingOwner
                AlreadyExisted = $true
                Group          = $existingGroup
            }
        }

        Write-Warning -Message ('{0} do not yet exists' -f $Name)

        # Resolve the requested owner (like members are resolved via Get-Recipient). Unlike a member,
        # an owner cannot be skipped: the group must be created with a ManagedBy, so if the recipient
        # cannot be resolved we warn and fall back to the raw -ManagedBy value.
        $resolvedOwner = $ManagedBy
        $ownerRecipient = Get-Recipient -Identity $ManagedBy -ErrorAction SilentlyContinue
        if ($ownerRecipient) {
            $resolvedOwner = [string]$ownerRecipient.PrimarySmtpAddress
            Write-Verbose -Message ("Owner '{0}' resolved to '{1}'." -f $ManagedBy, $resolvedOwner)
        }
        else {
            Write-Warning -Message ("Owner recipient '{0}' could not be resolved; using the requested value as-is." -f $ManagedBy)
        }

        $initialMembers = @()
        if ($BootstrapMember) { $initialMembers += $BootstrapMember }
        Write-Verbose -Message ("Unified Group '{0}' not found. Creating new group." -f $Name)

        # ---- Pre-call debug snapshot (minimal/essential creation parameters only) ----
        Write-Debug -Message ("[New-UnifiedGroup] PRE-CALL parameter snapshot (essential attributes):")
        Write-Debug -Message ("  -DisplayName     : '{0}'" -f $DisplayName)
        Write-Debug -Message ("  -Name            : '{0}'" -f $Name)
        Write-Debug -Message ("  -Alias           : '{0}'" -f $Name)
        Write-Debug -Message ("  -AccessType      : Private")
        Write-Debug -Message ("  -HiddenGroupMembershipEnabled : True (creation-only)")
        Write-Debug -Message ("  -ManagedBy       : '{0}' (requested: '{1}')" -f $resolvedOwner, $ManagedBy)
        Write-Debug -Message ("  -Members (count) : {0}  Values: [{1}]" -f $initialMembers.Count, (($initialMembers | Where-Object { $_ }) -join ', '))
        Write-Debug -Message ("  Calling context  : TenantId='{0}'; CallerAccount='{1}'" -f (Get-MgContext | Select-Object -ExpandProperty TenantId), (Get-MgContext | Select-Object -ExpandProperty Account))

        if (-not $PSCmdlet.ShouldProcess($Name, 'Create')) { return }

        Write-Debug -Message ("[New-UnifiedGroup] ShouldProcess approved - invoking New-UnifiedGroup for '{0}'." -f $Name)
        $nugInvokeStart = [datetime]::UtcNow
        try {
            $nug = New-UnifiedGroup `
                -DisplayName $DisplayName `
                -Name $Name `
                -Alias $Name `
                -AccessType Private `
                -HiddenGroupMembershipEnabled:$true `
                -ManagedBy $resolvedOwner `
                -Members $initialMembers `
                -ErrorAction Stop
            $nugElapsed = ([datetime]::UtcNow - $nugInvokeStart).TotalSeconds
            Write-Debug -Message ("[New-UnifiedGroup] Cmdlet returned after {0:N2} seconds. Raw return type: '{1}'." -f $nugElapsed, $(if ($null -ne $nug) { $nug.GetType().FullName } else { '<null>' }))
        }
        catch {
            $nugElapsed = ([datetime]::UtcNow - $nugInvokeStart).TotalSeconds
            Write-Debug -Message ("[New-UnifiedGroup] EXCEPTION after {0:N2} seconds." -f $nugElapsed)
            Write-Debug -Message ("[New-UnifiedGroup] Failed for '{0}'." -f $Name)
            Write-Debug -Message ("  ManagedBy       : '{0}' (requested: '{1}')" -f $resolvedOwner, $ManagedBy)
            Write-Debug -Message ("  BootstrapMembers: '{0}'" -f (($initialMembers | Where-Object { $_ }) -join ', '))
            Write-Debug -Message ("  Exception type  : '{0}'" -f $_.Exception.GetType().FullName)
            Write-Debug -Message ("  Exception msg   : '{0}'" -f $_.Exception.Message)
            Write-Debug -Message ("  ScriptStackTrace: {0}" -f $_.ScriptStackTrace)
            throw
        }

        if ($null -ne $nug) {
            Write-Verbose -Message ("Unified Group '{0}' created successfully." -f $Name)
            Write-Debug -Message ("[New-UnifiedGroup] Created object properties:")
            Write-Debug -Message ("  DisplayName                  : '{0}'" -f $nug.DisplayName)
            Write-Debug -Message ("  ExternalDirectoryObjectId    : '{0}'" -f $nug.ExternalDirectoryObjectId)
            Write-Debug -Message ("  PrimarySmtpAddress           : '{0}'" -f $nug.PrimarySmtpAddress)
            Write-Debug -Message ("  Alias                        : '{0}'" -f $nug.Alias)
            Write-Debug -Message ("  AccessType                   : '{0}'" -f $nug.AccessType)
            Write-Debug -Message ("  HiddenGroupMembershipEnabled : '{0}'" -f $nug.HiddenGroupMembershipEnabled)
            Write-Debug -Message ("  WhenCreatedUTC               : '{0}'" -f $nug.WhenCreatedUTC)

            Write-Verbose -Message ("Applying post-creation settings to Unified Group '{0}'." -f $Name)
            Write-Debug -Message ("[Set-UnifiedGroup] Applying settings to '{0}':" -f $Name)
            Write-Debug -Message ("  -IsMemberAllowedToEditContent           : False")
            Write-Debug -Message ("  -AutoSubscribeNewMembers                : False")
            Write-Debug -Message ("  -AlwaysSubscribeMembersToCalendarEvents : False")
            Write-Debug -Message ("  -Language                               : en-us")
            Write-Debug -Message ("  -SubscriptionEnabled                    : False")
            Write-Debug -Message ("  -HiddenFromAddressListsEnabled          : True")
            Write-Debug -Message ("  -ConnectorsEnabled                      : False")
            Set-UnifiedGroup -Identity $Name `
                -IsMemberAllowedToEditContent $false `
                -AutoSubscribeNewMembers:$false `
                -AlwaysSubscribeMembersToCalendarEvents:$false `
                -Language en-us `
                -SubscriptionEnabled:$false `
                -HiddenFromAddressListsEnabled:$true `
                -ConnectorsEnabled:$false `
                -ErrorAction Stop

            $configuredGroup = Get-UnifiedGroup -Identity $Name -ErrorAction SilentlyContinue
            if ($configuredGroup) {
                Write-Debug -Message ("[New-UnifiedGroup] Post-creation Set-UnifiedGroup verification:")
                Write-Debug -Message ("  Identity                     : '{0}'" -f $configuredGroup.Identity)
                Write-Debug -Message ("  HiddenFromAddressListsEnabled: '{0}'" -f $configuredGroup.HiddenFromAddressListsEnabled)
                Write-Debug -Message ("  AccessType                   : '{0}'" -f $configuredGroup.AccessType)
                Write-Debug -Message ("  HiddenGroupMembershipEnabled : '{0}'" -f $configuredGroup.HiddenGroupMembershipEnabled)
                Write-Debug -Message ("  SubscriptionEnabled          : '{0}'" -f $configuredGroup.SubscriptionEnabled)
                Write-Debug -Message ("  ConnectorsEnabled            : '{0}'" -f $configuredGroup.ConnectorsEnabled)
                Write-Debug -Message ("  ManagedBy                    : '{0}'" -f (($configuredGroup.ManagedBy | Where-Object { $_ }) -join ', '))
                Write-Debug -Message ("  MemberCount                  : '{0}'" -f $configuredGroup.GroupMemberCount)
                return [pscustomobject]@{
                    Name           = $Name
                    DisplayName    = $configuredGroup.DisplayName
                    OwnerRequested = $ManagedBy
                    OwnerAdded     = $resolvedOwner
                    AlreadyExisted = $false
                    Group          = $configuredGroup
                }
            }
            else {
                Write-Debug -Message ("[New-UnifiedGroup] WARNING: Get-UnifiedGroup returned null after Set-UnifiedGroup for '{0}'. Post-create verification skipped." -f $Name)
            }

            return [pscustomobject]@{
                Name           = $Name
                DisplayName    = $nug.DisplayName
                OwnerRequested = $ManagedBy
                OwnerAdded     = $resolvedOwner
                AlreadyExisted = $false
                Group          = $nug
            }
        }
    }
}
