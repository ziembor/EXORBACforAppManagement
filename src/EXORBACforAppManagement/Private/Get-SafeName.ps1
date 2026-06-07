function Get-SafeName {
    # Strip everything except letters, digits, and dash, then cap the length.
    # Used to build Unified Group names and RBAC assignment names.
    param(
        [string] $s,
        [int] $max = 63
    )

    $clean = ($s -replace '[^A-Za-z0-9\-]', '')
    if ($clean.Length -le $max) { return $clean }
    $clean.Substring(0, $max)
}
