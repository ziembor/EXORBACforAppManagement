function Get-NormalizeRole {
    # Normalize a short application-role name (e.g. Mail.Send) to its full EXO role
    # name (Application Mail.Send). Unknown / already-normalized values pass through.
    param(
        [string] $r
    )

    if (-not $r) { return $r }

    $roleMap = Get-AppRoleMap
    if ($roleMap.Contains($r)) { return $r }

    $applicationRole = 'Application {0}' -f $r
    if ($roleMap.Contains($applicationRole)) { return $applicationRole }

    return $r
}
