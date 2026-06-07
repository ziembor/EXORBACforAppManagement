<#
.SYNOPSIS
Creates an Entra (Azure AD) application registration and, by default, its service principal.

.DESCRIPTION
New-RegisteredApp creates an application registration via Microsoft Graph
(New-MgApplication) and, unless -SkipServicePrincipal is supplied, the matching
enterprise application / service principal (New-MgServicePrincipal). It is intended
as a companion to New-RBACforAppEntry: the emitted object exposes AppId and
ServicePrincipalId so the new application can be piped straight into that function.

The function supports -WhatIf and -Confirm through SupportsShouldProcess.

.PARAMETER DisplayName
Display name of the application registration to create.

.PARAMETER SignInAudience
Account types allowed to sign in. Defaults to AzureADMyOrg (single tenant).

.PARAMETER RedirectUri
Optional web redirect URIs to configure on the application.

.PARAMETER Description
Optional description (notes) stored on the application registration.

.PARAMETER RequiredResourceAccess
Optional API permission definitions passed straight through to
New-MgApplication -RequiredResourceAccess.

.PARAMETER SkipServicePrincipal
Create only the application registration; do not create the service principal.

.EXAMPLE
New-RegisteredApp -DisplayName 'Contoso Mail App' -Verbose -WhatIf

Shows the planned application and service principal creation without making changes.

.EXAMPLE
New-RegisteredApp -DisplayName 'Contoso Mail App' | New-RBACforAppEntry -Members 'shared@contoso.com' -Role 'Mail.Send'

Creates the application and service principal, then pipes the new AppId into
New-RBACforAppEntry to scope Exchange Online RBAC for it.

.OUTPUTS
PSCustomObject

Returns a summary object with the display name, AppId, application object id,
service principal object id (when created), sign-in audience, warnings, and errors.

.NOTES
Requires a connected Microsoft Graph session (Connect-MgGraph) with permission to
create applications and service principals (e.g. Application.ReadWrite.All). Uses
New-MgApplication and New-MgServicePrincipal from the Microsoft.Graph modules.
#>
function New-RegisteredApp {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name','RegisteredAppName')]
        [ValidateNotNullOrEmpty()]
        [string] $DisplayName,

        [Parameter()]
        [ValidateSet('AzureADMyOrg','AzureADMultipleOrgs','AzureADandPersonalMicrosoftAccount','PersonalMicrosoftAccount')]
        [string] $SignInAudience = 'AzureADMyOrg',

        [Parameter()]
        [string[]] $RedirectUri,

        [Parameter()]
        [string] $Description,

        [Parameter()]
        [object[]] $RequiredResourceAccess,

        [Parameter()]
        [switch] $SkipServicePrincipal
    )

    begin {
        $tenantId = Get-MgContext | Select-Object -ExpandProperty TenantId
    }

    process {
        $result = [ordered]@{
            DisplayName          = $DisplayName
            AppId                = $null
            AppObjectId          = $null
            ServicePrincipalId   = $null
            SignInAudience       = $SignInAudience
            TenantId             = $tenantId
            Warnings             = @()
            Errors               = @()
        }

        try {
            # --- Build New-MgApplication parameters (only include optional ones when supplied)
            $appParams = @{
                DisplayName    = $DisplayName
                SignInAudience = $SignInAudience
                ErrorAction    = 'Stop'
            }
            if ($Description)            { $appParams['Description'] = $Description }
            if ($RequiredResourceAccess) { $appParams['RequiredResourceAccess'] = $RequiredResourceAccess }
            if ($RedirectUri)            { $appParams['Web'] = @{ RedirectUris = @($RedirectUri) } }

            # --- Create the application registration
            $app = $null
            if ($PSCmdlet.ShouldProcess($DisplayName, 'Create application registration')) {
                Write-Verbose -Message ("Creating application registration '{0}' (audience '{1}')." -f $DisplayName, $SignInAudience)
                $app = New-MgApplication @appParams
                $result.AppId       = $app.AppId
                $result.AppObjectId = $app.Id
                Write-Verbose -Message ("Application created. AppId='{0}'; ObjectId='{1}'." -f $app.AppId, $app.Id)
            }

            # --- Create the matching service principal (enterprise application)
            if (-not $SkipServicePrincipal) {
                if ($null -eq $app) {
                    # -WhatIf path: application was not actually created, so nothing to attach an SP to.
                    Write-Verbose -Message 'Skipping service principal creation because the application was not created (WhatIf).'
                }
                elseif ($PSCmdlet.ShouldProcess($DisplayName, 'Create service principal')) {
                    Write-Verbose -Message ("Creating service principal for AppId '{0}'." -f $app.AppId)
                    $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
                    $result.ServicePrincipalId = $sp.Id
                    Write-Verbose -Message ("Service principal created. ObjectId='{0}'." -f $sp.Id)
                }
            }
            else {
                $result.Warnings += 'Service principal creation skipped (-SkipServicePrincipal).'
            }

            [pscustomobject]$result
        }
        catch {
            $result.Errors += $_.Exception.Message
            [pscustomobject]$result
        }
    }
}
