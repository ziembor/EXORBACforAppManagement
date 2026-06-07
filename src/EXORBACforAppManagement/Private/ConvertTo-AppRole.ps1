function ConvertTo-AppRole {
    # All EXO application roles are named "Application <permission>", so normalizing a
    # short name just means prefixing "Application " when it is not already present.
    param(
        [string] $r
    )

    if ($r -match '^Application\s+') { return $r }
    return "Application $r"
}
