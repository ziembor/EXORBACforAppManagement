function Resolve-AppRolePermissionValue {
    # Resolve a Microsoft Graph app-role assignment (an application permission grant) to
    # its permission value string (e.g. 'Mail.Read'). The grant carries the resource
    # service principal id and the AppRoleId GUID; the human-readable value lives in the
    # resource SP's AppRoles collection. Resource SPs are cached by ResourceId in the
    # supplied hashtable so each resource (e.g. Microsoft Graph) is fetched only once.
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [object] $Grant,

        [Parameter(Mandatory)]
        [hashtable] $Cache
    )

    if (-not $Grant.AppRoleId -or -not $Grant.ResourceId) { return $null }

    $resourceId = [string]$Grant.ResourceId
    if (-not $Cache.ContainsKey($resourceId)) {
        $Cache[$resourceId] = Get-MgServicePrincipal -ServicePrincipalId $resourceId -ErrorAction SilentlyContinue
    }
    $resourceSp = $Cache[$resourceId]
    if (-not $resourceSp) { return $null }

    $appRole = @($resourceSp.AppRoles) | Where-Object { [string]$_.Id -eq [string]$Grant.AppRoleId } | Select-Object -First 1
    if (-not $appRole) { return $null }

    return [string]$appRole.Value
}
