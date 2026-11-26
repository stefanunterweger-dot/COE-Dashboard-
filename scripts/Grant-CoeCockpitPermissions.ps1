<#
.SYNOPSIS
  Gibt den SharePoint-Webparts (SPFx) die delegierten Berechtigungen, die das CoE Cockpit zusätzlich braucht.

.DESCRIPTION
  Normalerweise genehmigt man die Anforderungen des Pakets im SharePoint Admin Center unter „Erweitert → API-Zugriff“.
  Erscheinen dort keine Anforderungen, setzt dieses Skript dieselben Berechtigungen direkt (Microsoft Graph PowerShell):

    -Dataverse         Dynamics CRM „user_impersonation“ – CoE-Daten lesen (Pflicht für das Cockpit)
    -MailSend          Microsoft Graph „Mail.Send“ – automatischer E-Mail-Versand an Maker (optional)
    -Licenses          Microsoft Graph „User.Read.All“, „GroupMember.Read.All“ und „Organization.Read.All“ – Lizenzprüfung in
                       „Betroffen ab Februar 2027“ (Lizenzen der Berechtigten, Mitglieder freigegebener Gruppen) und
                       „Lizenzen je Benutzer“ (gekaufte/freie Lizenzen aus „subscribedSkus“)
    -PowerPlatformApi  Power Platform API „Licensing.Allocations.Read“, „Licensing.BillingPolicies.Read“,
                       „EnvironmentManagement.Environments.Read“ – Ansichten „Copilot-Credits“ und „Abrechnungspläne“
    -Actions           Aktionen direkt im Webpart (ohne Flow, im Namen der angemeldeten Person):
                         Power Platform API   „EnvironmentManagement.Write“, „EnvironmentManagement.Settings.ReadWrite“
                                              (Environment auf Managed umstellen, deaktivieren, aktivieren)
                         PowerApps Service    „User“ (Apps: Quarantäne, Besitzer, Freigaben, Löschen)
                         Microsoft Flow Service „User“, „Flows.Read.All“, „Flows.Manage.All“
                                              (Flows: ein/aus, Mitbesitzer, Freigaben, Löschen)
                         Microsoft Graph      „User.Read.All“ (Person über die E-Mail-Adresse finden)

  Ohne Schalter werden alle gesetzt. Bestehende Berechtigungen bleiben erhalten; es wird nur ergänzt.
  Fehlt im Tenant der Dienstprinzipal einer Microsoft-App (z. B. „Power Platform API“), legt das Skript ihn an.

  Wichtig:
   - Diese Berechtigungen gelten – wie eine Genehmigung unter „API-Zugriff“ – für ALLE SharePoint-Webparts im Tenant.
   - Sie geben niemandem zusätzliche Rechte: Microsoft prüft bei jedem Aufruf die Rolle der angemeldeten Person.
     Ausführen kann eine Aktion nur, wer die Rolle „Power Platform Administrator“ (oder „Globaler Administrator“) aktiv hat.

  Ausführen als Globaler Administrator oder Administrator für privilegierte Rollen
  (für Power Platform API, PowerApps Service und Flow Service reicht auch Anwendungsadministrator / Cloudanwendungsadministrator).

.EXAMPLE
  .\Grant-CoeCockpitPermissions.ps1
.EXAMPLE
  .\Grant-CoeCockpitPermissions.ps1 -Actions
#>
param(
  [switch]$Dataverse,
  [switch]$MailSend,
  [switch]$Licenses,
  [switch]$PowerPlatformApi,
  [switch]$Actions
)
$ErrorActionPreference = 'Stop'
if (-not $Dataverse -and -not $MailSend -and -not $Licenses -and -not $PowerPlatformApi -and -not $Actions) {
  $Dataverse = $true; $MailSend = $true; $Licenses = $true; $PowerPlatformApi = $true; $Actions = $true
}

# Install-Module Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns -Scope CurrentUser   # nur falls noch nicht installiert
Connect-MgGraph -Scopes 'Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All' | Out-Null

# Über diese Microsoft-Apps holen sich SharePoint-Webparts ihre Tokens (neu und – falls noch vorhanden – alt)
$spfx = @('SharePoint Online Web Client Extensibility', 'SharePoint Online Client Extensibility Web Application Principal') |
  ForEach-Object { Get-MgServicePrincipal -Filter "displayName eq '$_'" } | Where-Object { $_ }
if (-not $spfx) { throw 'Der SharePoint-Dienstprinzipal für Webparts wurde nicht gefunden.' }
$spfx | Format-Table DisplayName, AppId, Id

# Dienstprinzipal einer Microsoft-App holen (bei Bedarf im Tenant anlegen)
function Get-ResourceSp([string]$appId, [string]$name) {
  $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'"
  if (-not $sp) {
    $sp = New-MgServicePrincipal -AppId $appId
    Write-Host "Dienstprinzipal „$name“ im Tenant angelegt."
    Start-Sleep -Seconds 10
    $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'"
  }
  return $sp
}

function Grant-Scopes($resource, [string[]]$scopes) {
  # nur Berechtigungen, die die App auch anbietet
  $offered = @($resource.Oauth2PermissionScopes | ForEach-Object { $_.Value })
  $missing = @($scopes | Where-Object { $offered -notcontains $_ })
  if ($missing) { Write-Warning "„$($resource.DisplayName)“ bietet diese Berechtigungen nicht an (übersprungen): $($missing -join ', ')" }
  $scopes = @($scopes | Where-Object { $offered -contains $_ })
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

if ($Dataverse) {
  $crm = Get-ResourceSp '00000007-0000-0000-c000-000000000000' 'Dynamics CRM'
  Grant-Scopes $crm @('user_impersonation')
}

$graph = Get-ResourceSp '00000003-0000-0000-c000-000000000000' 'Microsoft Graph'
if ($MailSend) {
  Grant-Scopes $graph @('Mail.Send')
}
if ($Licenses) {
  Grant-Scopes $graph @('User.Read.All', 'GroupMember.Read.All', 'Organization.Read.All')
}

if ($PowerPlatformApi -or $Actions) {
  $pp = Get-ResourceSp '8578e004-a5c6-46e7-913e-12f58912df43' 'Power Platform API'
  if ($PowerPlatformApi) {
    Grant-Scopes $pp @('Licensing.Allocations.Read', 'Licensing.BillingPolicies.Read', 'EnvironmentManagement.Environments.Read')
  }
  if ($Actions) {
    Grant-Scopes $pp @('EnvironmentManagement.Write', 'EnvironmentManagement.Settings.ReadWrite')
  }
}

if ($Actions) {
  $apps = Get-ResourceSp '475226c6-020e-4fb2-8a90-7a972cbfc1d4' 'PowerApps Service'
  Grant-Scopes $apps @('User')
  $flow = Get-ResourceSp '7df0a125-d3be-4c96-aa54-591f83ff541c' 'Microsoft Flow Service'
  Grant-Scopes $flow @('User', 'Flows.Read.All', 'Flows.Manage.All')
  Grant-Scopes $graph @('User.Read.All')
}

Write-Host ''
Write-Host 'Fertig. Tokens werden zwischengespeichert: bis zu 1 Stunde warten oder ab- und wieder anmelden.'
