#requires -Version 5.1

# Module loader: dot-source every Private helper, then every Public function, and
# export only the Public functions. File names are expected to match function names.

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue)
$public  = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public')  -Filter '*.ps1' -ErrorAction SilentlyContinue)

foreach ($file in @($private + $public)) {
    try {
        . $file.FullName
    }
    catch {
        throw "Failed to import function file '$($file.FullName)': $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function $public.BaseName
