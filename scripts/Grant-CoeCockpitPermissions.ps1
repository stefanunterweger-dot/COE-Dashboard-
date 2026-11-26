<#
.SYNOPSIS
  Gibt den SharePoint-Webparts (SPFx) die delegierten Berechtigungen, die das CoE Cockpit zusätzlich braucht.

.DESCRIPTION
  Normalerweise genehmigt man die Anforderungen des Pakets im SharePoint Admin Center unter „Erweitert → API-Zugriff“.
  Erscheinen dort keine Anforderungen, setzt dieses Skript dieselben Berechtigungen direkt (Microsoft Graph PowerShell):

    -MailSend          Microsoft Graph „Mail.Send“ – automatischer E-Mail-Versand an Maker (optional)
    -PowerPlatformApi  Power Platform API „Licensing.Allocations.Read“, „Licensing.BillingPolicies.Read“,
                       „EnvironmentManagement.Environments.Read“ – Ansichten „Copilot-Credits“ und „Abrechnungspläne“

  Ohne Schalter werden beide gesetzt. Bestehende Berechtigungen bleiben erhalten; es wird nur ergänzt.
  Fehlt im Tenant der Dienstprinzipal „Power Platform API“, legt das Skript ihn an.

  Wichtig: Diese Berechtigungen gelten – wie eine Genehmigung unter „API-Zugriff“ – für ALLE SharePoint-Webparts im Tenant.

  Ausführen als Globaler Administrator oder Administrator für privilegierte Rollen
  (für die Power Platform API reicht auch Anwendungsadministrator / Cloudanwendungsadministrator).

.EXAMPLE
  .\Grant-CoeCockpitPermissions.ps1 -PowerPlatformApi
#>
param(
  [switch]$MailSend,
  [switch]$PowerPlatformApi
)
$ErrorActionPreference = 'Stop'
if (-not $MailSend -and -not $PowerPlatformApi) { $MailSend = $true; $PowerPlatformApi = $true }

# Install-Module Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns -Scope CurrentUser   # nur falls noch nicht installiert
Connect-MgGraph -Scopes 'Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All' | Out-Null

# Über diese Microsoft-Apps holen sich SharePoint-Webparts ihre Tokens (neu und – falls noch vorhanden – alt)
$spfx = @('SharePoint Online Web Client Extensibility', 'SharePoint Online Client Extensibility Web Application Principal') |
  ForEach-Object { Get-MgServicePrincipal -Filter "displayName eq '$_'" } | Where-Object { $_ }
if (-not $spfx) { throw 'Der SharePoint-Dienstprinzipal für Webparts wurde nicht gefunden.' }
$spfx | Format-Table DisplayName, AppId, Id

function Grant-Scopes($resource, [string[]]$scopes) {
  if (-not $scopes) { return }
  foreach ($sp in $spfx) {
    $grant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" -All |
      Where-Object { $_.ResourceId -eq $resource.Id -and $_.ConsentType -eq 'AllPrincipals' } | Select-Object -First 1
    if ($grant) {
      $have = @($grant.Scope -split ' ' | Where-Object { $_ })
      $all = @($have + $scopes | Select-Object -Unique)
      if ($all.Count -ne $have.Count) {
        Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -Scope ($all -join ' ')
      }
    } else {
      New-MgOauth2PermissionGrant -BodyParameter @{
        clientId = $sp.Id; consentType = 'AllPrincipals'; resourceId = $resource.Id; scope = ($scopes -join ' ')
      } | Out-Null
    }
    Write-Host "OK  $($sp.DisplayName) -> $($resource.DisplayName): $($scopes -join ', ')"
  }
}

if ($MailSend) {
  $graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
  Grant-Scopes $graph @('Mail.Send')
}

if ($PowerPlatformApi) {
  $ppAppId = '8578e004-a5c6-46e7-913e-12f58912df43'   # „Power Platform API“ (Microsoft)
  $pp = Get-MgServicePrincipal -Filter "appId eq '$ppAppId'"
  if (-not $pp) {
    $pp = New-MgServicePrincipal -AppId $ppAppId
    Write-Host 'Dienstprinzipal „Power Platform API“ im Tenant angelegt.'
    Start-Sleep -Seconds 10
    $pp = Get-MgServicePrincipal -Filter "appId eq '$ppAppId'"
  }
  $wanted = @('Licensing.Allocations.Read', 'Licensing.BillingPolicies.Read', 'EnvironmentManagement.Environments.Read')
  $offered = @($pp.Oauth2PermissionScopes | ForEach-Object { $_.Value })
  $missing = $wanted | Where-Object { $offered -notcontains $_ }
  if ($missing) { Write-Warning "Diese Berechtigungen bietet die Power Platform API nicht an und werden übersprungen: $($missing -join ', ')" }
  Grant-Scopes $pp @($wanted | Where-Object { $offered -contains $_ })
}

Write-Host ''
Write-Host 'Fertig. Tokens werden zwischengespeichert: bis zu 1 Stunde warten oder ab- und wieder anmelden.'
