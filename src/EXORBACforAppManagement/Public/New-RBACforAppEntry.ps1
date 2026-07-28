<#
.SYNOPSIS
Creates or updates Exchange Online RBAC scoping for an Entra application service principal.

.DESCRIPTION
New-RBACforAppEntry resolves an Entra service principal by display name, AppId, or
service principal object id, creates the scoped Unified Group when needed, adds
requested members, ensures the Exchange Online service principal exists, and creates
Exchange Online RBAC role assignments for the application.

Unified Group creation/configuration is delegated to New-RBACforAppUnifiedGroup and the
Exchange Online service principal step to Register-EXOServicePrincipal.

The function supports -WhatIf and -Confirm through SupportsShouldProcess.

.PARAMETER RegisteredAppName
Display name of the registered application or service principal. This is the default
parameter set and must resolve to exactly one service principal.

.PARAMETER AppId
Application (client) id of the registered application.

.PARAMETER SpObjectId
Object id of the target service principal.

.PARAMETER Members
Recipients to add to the Unified Group scope. Values must resolve through
Get-Recipient. The default placeholder value is GraphAPI-Dummy.

.PARAMETER Role
Exchange Online application roles to assign. Short names such as Mail.Send are
normalized to Application Mail.Send where supported.

.PARAMETER ManagedBy
Recipient that will be assigned as the Unified Group owner.

.PARAMETER GroupPrefix
Prefix used when building the Unified Group name.

.PARAMETER AccessGroupName
Explicit Unified Group name to use for RBAC scoping instead of generating a name from
GroupPrefix and the resolved service principal display name. Cannot be combined with
an explicit GroupPrefix value.

.PARAMETER BootstrapMember
Optional initial member passed during Unified Group creation.

.EXAMPLE
New-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -Verbose -WhatIf

Shows the planned service principal resolution, Unified Group creation, and RBAC
assignment actions without making changes.

.EXAMPLE
New-RBACforAppEntry -AppId '11111111-2222-3333-4444-555555555555' -Members 'sharedmailbox@contoso.com' -Role 'Mail.Send' -Verbose

Resolves the application by AppId, ensures the scoped Unified Group exists, adds the
recipient, and creates the Application Mail.Send role assignment.

.EXAMPLE
New-RBACforAppEntry -SpObjectId '11111111-2222-3333-4444-555555555555' -Role 'Application Calendars.Read','Application Contacts.Read' -GroupPrefix 'Um365Prod'

Uses the service principal object id directly and creates multiple application role
assignments scoped to the generated Unified Group.

.EXAMPLE
New-RBACforAppEntry -RegisteredAppName 'Contoso Mail App' -AccessGroupName 'RBAC-AppScope-ContosoMail' -Role 'Mail.Send'

Uses an explicit Unified Group name for scoping and assigns the requested RBAC role to
the application against that group.

.OUTPUTS
PSCustomObject

Returns a summary object with resolved identity, Unified Group name, normalized roles,
assignment names, warnings, and errors.

.NOTES
Requires Microsoft Graph and Exchange Online cmdlets used by Get-MgServicePrincipal,
New-ServicePrincipal, Get-UnifiedGroup, Set-UnifiedGroup, Add-UnifiedGroupLinks, and
New-ManagementRoleAssignment.
#>
function New-RBACforAppEntry {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'ByName')]
    param(
        # Default: resolve by displayName (can be non-unique; will error if ambiguous)
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ByName')]
        [Alias('DisplayName','Name')]
        [ValidateNotNullOrEmpty()]
        [string] $RegisteredAppName,

        # Alternative: resolve by AppId (GUID)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ByAppId')]
        [Alias('ClientId','ApplicationId')]
        [ValidatePattern('^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}$')]
        [string] $AppId,

        # Alternative: resolve by Service Principal ObjectId (GUID)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'BySpObjectId')]
        [Alias('Id','ObjectId','ServicePrincipalId')]
        [ValidatePattern('^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}$')]
        [string] $SpObjectId,

        [Parameter(Position = 1)]
        [string[]] $Members = @("GraphAPI-Dummy"),

        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Role = @("Application Mail.Send"),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $ManagedBy = "GraphAPI-Dummy-owner",

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $GroupPrefix = "Um365RAo1",

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $AccessGroupName,

        # Optional placeholder member (dont validate as email)
        [Parameter()]
        [string] $BootstrapMember = "GraphAPI-Dummy"
    )

    begin {
        $shortRoleMap = Get-AppRoleMap

        $tenantid = Get-MgContext | Select-Object -ExpandProperty TenantId
    }

    process {
        if ($PSBoundParameters.ContainsKey('AccessGroupName') -and $PSBoundParameters.ContainsKey('GroupPrefix')) {
            throw "Parameters -AccessGroupName and -GroupPrefix cannot be used together."
        }

        $result = [ordered]@{
            ParameterSet      = $PSCmdlet.ParameterSetName
            IdentityInput     = $RegisteredAppName
            ResolvedDisplay   = $null
            AppId             = $null
            SpObjectId        = $null
            TenantId          = $tenantid
            UnifiedGroupName  = $null
            OwnerRequested    = $ManagedBy
            OwnerAdded        = $null
            MembersRequested  = @($Members)
            MembersAdded      = @()
            FilteredMembers   = @()
            RolesNormalized   = @()
            RoleAssignments   = @()
            RoleAssignmentsName = @()
            Warnings          = @()
            Errors            = @()
        }

        try {
            # --- Resolve service principal depending on parameter set
            $sp = $null
            switch ($PSCmdlet.ParameterSetName) {
                'BySpObjectId' {
                    $sp = Get-MgServicePrincipal -ServicePrincipalId $SpObjectId -ErrorAction Stop
                }

                'ByAppId' {
                    $filter = "appId eq `'$AppId`'"
                    $matchesRes = @(Get-MgServicePrincipal -Filter $filter -ErrorAction Stop)
                    if ($matchesRes.Count -eq 0) { throw "No service principal found for AppId '$AppId'." }
                    if ($matchesRes.Count -gt 1) { throw "Unexpected: multiple service principals for AppId '$AppId'." }
                    $sp = $matchesRes[0]
                }

                'ByName' {
                    $filter = "displayName eq `'$RegisteredAppName`'"
                    $matchesRes = @(Get-MgServicePrincipal -Filter $filter -ErrorAction Stop)
                    if ($matchesRes.Count -eq 0) { throw "No service principal found for displayName '$RegisteredAppName'." }
                    if ($matchesRes.Count -gt 1) {
                        $names = ($matchesRes | Select-Object -First 10 -ExpandProperty Id) -join ', '
                        throw "Ambiguous displayName '$RegisteredAppName' matched $($matchesRes.Count) service principals. Re-run with -AppId or -SpObjectId. Example SP objectIds: $names"
                    }
                    $sp = $matchesRes[0]
                }
            }

            $result.ResolvedDisplay = $sp.DisplayName
            $result.AppId           = $sp.AppId
            $result.SpObjectId      = $sp.Id

            # --- Ensure Unified Group for scoping (delegated to New-RBACforAppUnifiedGroup)
            if ($PSBoundParameters.ContainsKey('AccessGroupName')) {
                $umGroupName = $AccessGroupName
            }
            else {
                $umGroupName = "{0}-{1}" -f $GroupPrefix, $sp.DisplayName
                $umGroupName = Get-SafeName($umGroupName)
            }
            $result.UnifiedGroupName = $umGroupName

            Write-Verbose -Message ("Checking Unified Group '{0}' for service principal '{1}' ({2})." -f $umGroupName, $sp.DisplayName, $sp.Id)
            $ugResult = New-RBACforAppUnifiedGroup -Name $umGroupName -ManagedBy $ManagedBy -BootstrapMember $BootstrapMember -WarningVariable ugWarnings
            foreach ($w in $ugWarnings) {
                if ([string]$w.Message -like '*already exists*') { $result.Warnings += [string]$w.Message }
            }
            if ($ugResult) {
                $result.OwnerRequested = $ugResult.OwnerRequested
                $result.OwnerAdded    = $ugResult.OwnerAdded
            }

            # --- Add members
            $currentUserUpn = $null
            try {
                $connectionInfo = Get-MgContext -ErrorAction Stop
                $currentUserUpn = $connectionInfo.Account
            }
            catch {
                $result.Warnings += "Could not retrieve current connection user via get-mgContext; connection user filtering will be skipped. Error: $($_.Exception.Message)"
                Write-Warning -Message "Could not retrieve current connection user via get-mgContext; connection user filtering will be skipped."
            }

            foreach ($member in $Members) {
                if (-not $member) { continue }

                $rec = Get-Recipient -Identity $member -ErrorAction SilentlyContinue
                if (-not $rec) {
                    $result.Warnings += "Recipient not found for '$member' (skipped)."
                    continue
                }
                if ($currentUserUpn -and ($member -ieq $currentUserUpn)) {
                # if ($currentUserUpn -and ($rec.PrimarySmtpAddress -ieq $currentUserUpn -or $member -ieq $currentUserUpn)) {
                    $result.FilteredMembers += [string]$rec.PrimarySmtpAddress
                    $filterWarning = "Current connection user '$currentUserUpn' was found in the members list and has been filtered out."
                    $result.Warnings += $filterWarning
                    Write-Warning -Message $filterWarning
                    continue
                }

                if ($PSCmdlet.ShouldProcess("UnifiedGroup $umGroupName", "Add member $($rec.PrimarySmtpAddress)")) {
                    Add-UnifiedGroupLinks -Identity $umGroupName -LinkType Members -Links $rec.PrimarySmtpAddress -ErrorAction Stop
                }
                $result.MembersAdded += [string]$rec.PrimarySmtpAddress
            }

            # --- Ensure EXO ServicePrincipal extension (delegated to Register-EXOServicePrincipal)
            $exoSpDisplay = "{0}_SP" -f $sp.DisplayName
            $null = Register-EXOServicePrincipal -AppId $sp.AppId -ObjectId $sp.Id -DisplayName $exoSpDisplay

            # --- Role assignments
            $rolesNormalized = foreach ($r in @($Role)) { Get-NormalizeRole $r }
            $result.RolesNormalized = @($rolesNormalized)

            foreach ($roleItem in $rolesNormalized) {
                $ShortRoleName = $shortRoleMap[$roleItem]

                $rbacNameBase = Get-SafeName -s ("{0}-{1}" -f $ShortRoleName,$sp.DisplayName) -max 63
                $result.RoleAssignmentsName += $rbacNameBase

                if ($PSCmdlet.ShouldProcess("RBAC role assignment", "Assign '$roleItem' to App '$($sp.DisplayName)' scoped to '$umGroupName'")) {
                    $assignment = New-ManagementRoleAssignment `
                        -App $sp.Id `
                        -Role $roleItem `
                        -RecipientGroupScope $umGroupName `
                        -Name $rbacNameBase `
                        -ErrorAction Stop
                    $result.RoleAssignments += $assignment
                }
            }

            [pscustomobject]$result
            [pscustomobject]$result | Export-Clixml ('{0}/{1}_{2}.clixml' -f $env:TEMP,$rbacNameBase,(get-date -format s).Replace(':','')) -Verbose
        }
        catch {
            $result.Errors += $_.Exception.Message
            [pscustomobject]$result
        }
    }
}
