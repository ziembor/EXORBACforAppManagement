<#
.SYNOPSIS
Ensures the Exchange Online service principal pointer for an Entra application exists.

.DESCRIPTION
Register-EXOServicePrincipal creates the Exchange Online service principal (via New-ServicePrincipal)
that links an Entra (Azure AD) application registration into Exchange Online so it can receive
application RBAC role assignments. Exchange treats this object as a pointer to the existing Entra
service principal.

The function supports -WhatIf and -Confirm through SupportsShouldProcess.

.PARAMETER AppId
Application (client) id of the Entra application. GUID-validated.

.PARAMETER ObjectId
Object id of the Entra service principal (Enterprise Application object id). GUID-validated.

.PARAMETER DisplayName
Display name for the Exchange Online service principal (e.g. "<App>_SP").

.EXAMPLE
Register-EXOServicePrincipal -AppId '1111...' -ObjectId '2222...' -DisplayName 'Contoso_SP' -WhatIf

Shows the planned Exchange Online service principal creation without making changes.

.OUTPUTS
The object returned by New-ServicePrincipal (when created).

.NOTES
Requires a connected Exchange Online session (New-ServicePrincipal). Companion to
New-RBACforAppEntry.
#>
function Register-EXOServicePrincipal {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([object])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('ClientId','ApplicationId')]
        [ValidatePattern('^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}$')]
        [string] $AppId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('SpObjectId','Id','ServicePrincipalId')]
        [ValidatePattern('^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}$')]
        [string] $ObjectId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string] $DisplayName
    )

    process {
        # Typical usage expects AppId + Enterprise AppObjectId (SP object id).
        if ($PSCmdlet.ShouldProcess("EXO ServicePrincipal for AppId $AppId", "Create/Ensure '$DisplayName'")) {
            $resultNewSP = New-ServicePrincipal -AppId $AppId -ObjectId $ObjectId -DisplayName $DisplayName -ErrorAction 0
            Write-Warning -Message "resultNewSP: $resultNewSP"
            return $resultNewSP
        }
    }
}
