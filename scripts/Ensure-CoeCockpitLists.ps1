<#
.SYNOPSIS
  Prüft die SharePoint-Listen des CoE Cockpits – Listen, Spalten, Ansichten und Berechtigungen – und legt Fehlendes an.

.DESCRIPTION
  Das CoE Cockpit verwendet auf der Website, auf der es eingebunden ist, zwei Listen:

    „CoE Cockpit Aktionen“          Aktionsprotokoll (Quarantäne, Löschen, Besitzer, Freigaben, Flows, Environments,
                                    DLP-Richtlinien) und Aufträge für den Flow „CoE Cockpit | Aktion ausführen“.
                                    Eigene Berechtigungen: SharePoint-Gruppe (Standard „CoE Cockpit Admins“) = Mitwirken,
                                    Websitebesitzer = Vollzugriff, sonst niemand.
    „CoE Cockpit Benachrichtigungen“ Versandprotokoll (welches Environment wann von wem verständigt wurde).
                                    Erbt die Berechtigungen der Website.

  Das Skript prüft je Liste:
    - ob die Liste vorhanden ist (sonst anlegen),
    - jede Spalte: vorhanden, Typ, Anzeigename, Länge/Format (sonst anlegen bzw. korrigieren; falscher Typ = Fehler),
    - die Indizes, nach denen die Ansichten filtern/sortieren (sonst brechen sie ab 5.000 Einträgen ab),
    - jede Ansicht: vorhanden, Spalten, Filter/Sortierung/Gruppierung, Zeilenzahl, Standardansicht (sonst anlegen bzw. korrigieren),
    - bei „CoE Cockpit Aktionen“ den Versionsverlauf und die Berechtigungen (eigene Berechtigungen, Gruppe mit für alle
      sichtbarer Mitgliedschaft, Websitebesitzer, optional Flow-Dienstkonto).
  Am Ende steht ein Bericht (OK / ANGELEGT / KORRIGIERT / FEHLT / WARNUNG / FEHLER). Mit -CheckOnly wird nur geprüft.

  Die Definition entspricht dem Webpart (src/webparts/coeCockpit/services/ActionService.ts und NotifyLogService.ts):
  Ändern sich dort Listen oder Spalten, wird dieses Skript mit angepasst.

  Anmeldung: interaktiv mit PnP.PowerShell über die eigene App-Registrierung (Client-ID). Die App braucht:
    - Plattform „Mobile and desktop applications“ mit Umleitungs-URI http://localhost (öffentlicher Client),
    - delegierte SharePoint-Berechtigung „AllSites.FullControl“ (für Listen und Berechtigungen) mit Administratorzustimmung.
  Ausgeführt wird alles mit den Rechten der angemeldeten Person (Besitzer der Website bzw. Websitesammlungsadministrator).

  Voraussetzung: PowerShell 7.4 oder neuer und PnP.PowerShell
    Install-Module PnP.PowerShell -Scope CurrentUser

.PARAMETER SiteUrl
  Website, auf der das CoE Cockpit eingebunden ist, z. B. https://contoso.sharepoint.com/sites/powerplatform (wird abgefragt, wenn leer).

.PARAMETER ClientId
  Client-ID (Anwendungs-ID) der eigenen App-Registrierung für PnP.PowerShell (wird abgefragt, wenn leer).

.PARAMETER GroupName
  SharePoint-Gruppe für die Aktionen – derselbe Name wie in den Webpart-Einstellungen (Standard „CoE Cockpit Admins“).

.PARAMETER FlowServiceAccount
  Optional: E-Mail-Adresse des Dienstkontos, unter dem der Flow „CoE Cockpit | Aktion ausführen“ läuft – bekommt „Mitwirken“ auf die Aktionsliste.

.PARAMETER CheckOnly
  Nur prüfen und berichten, nichts anlegen oder ändern.

.EXAMPLE
  .\Ensure-CoeCockpitLists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/powerplatform -ClientId 00000000-0000-0000-0000-000000000000 -CheckOnly
.EXAMPLE
  .\Ensure-CoeCockpitLists.ps1
  (fragt Website und Client-ID ab, prüft und legt Fehlendes an)
.EXAMPLE
  .\Ensure-CoeCockpitLists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/powerplatform -FlowServiceAccount svc-coe@contoso.com
#>
[CmdletBinding()]
param(
  [string]$SiteUrl,
  [string]$ClientId,
  [string]$GroupName = 'CoE Cockpit Admins',
  [string]$FlowServiceAccount,
  [switch]$CheckOnly
)
$ErrorActionPreference = 'Stop'

if (-not $SiteUrl) { $SiteUrl = Read-Host 'URL der SharePoint-Website mit dem CoE Cockpit (https://…/sites/…)' }
if (-not $ClientId) { $ClientId = Read-Host 'Client-ID der App-Registrierung (PnP.PowerShell, interaktive Anmeldung)' }
$SiteUrl = $SiteUrl.Trim().TrimEnd('/')
$ClientId = $ClientId.Trim()
if ($SiteUrl -notmatch '^https://[^/]+\.sharepoint\.com/') { throw "Keine gültige SharePoint-URL: $SiteUrl" }
if ($ClientId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { throw "Keine gültige Client-ID (GUID): $ClientId" }

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  throw 'PnP.PowerShell ist nicht installiert: Install-Module PnP.PowerShell -Scope CurrentUser'
}
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ---------------------------------------------------------------------------------------------------------------
# Definition (wie im Webpart). Spalten: interner Name = Name im Webpart. Ansichten: Spalten in dieser Reihenfolge.
# ---------------------------------------------------------------------------------------------------------------
$OrderNewest = '<OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy>'
$Lists = @(
  [ordered]@{
    Title          = 'CoE Cockpit Aktionen'
    Description    = 'Aktionsprotokoll des CoE Cockpits (Quarantäne, Löschen, Besitzer, Freigaben, Flows ein/aus, Environments, Solutions, DLP-Richtlinien). Das Cockpit führt die meisten Aktionen direkt aus; „Solution löschen“ übernimmt der Flow im CoE-Environment.'
    TitleField     = 'Eintrag'
    UniquePermissions = $true
    Versioning     = $true
    # Indizes: die Ansichten filtern/sortieren/gruppieren danach – ohne Index brechen sie ab 5.000 Einträgen ab
    Indexes        = @('Created', 'Status', 'Action', 'Author')
    Fields = @(
      '<Field Type="Text" DisplayName="Action" Name="Action" StaticName="Action" MaxLength="40" Indexed="TRUE" />'
      '<Field Type="Text" DisplayName="ResourceType" Name="ResourceType" StaticName="ResourceType" MaxLength="20" />'
      '<Field Type="Text" DisplayName="ResourceId" Name="ResourceId" StaticName="ResourceId" MaxLength="80" />'
      '<Field Type="Text" DisplayName="CoeRecordId" Name="CoeRecordId" StaticName="CoeRecordId" MaxLength="80" />'
      '<Field Type="Text" DisplayName="ResourceName" Name="ResourceName" StaticName="ResourceName" MaxLength="255" />'
      '<Field Type="Text" DisplayName="EnvironmentId" Name="EnvironmentId" StaticName="EnvironmentId" MaxLength="120" />'
      '<Field Type="Text" DisplayName="TargetUser" Name="TargetUser" StaticName="TargetUser" MaxLength="255" />'
      '<Field Type="Text" DisplayName="TargetName" Name="TargetName" StaticName="TargetName" MaxLength="255" />'
      '<Field Type="Boolean" DisplayName="TargetIsGroup" Name="TargetIsGroup" StaticName="TargetIsGroup"><Default>0</Default></Field>'
      '<Field Type="Note" DisplayName="Reason" Name="Reason" StaticName="Reason" NumLines="4" RichText="FALSE" />'
      '<Field Type="Text" DisplayName="Status" Name="Status" StaticName="Status" MaxLength="40" Indexed="TRUE" />'
      '<Field Type="Note" DisplayName="Result" Name="Result" StaticName="Result" NumLines="4" RichText="FALSE" />'
      '<Field Type="DateTime" DisplayName="CompletedAt" Name="CompletedAt" StaticName="CompletedAt" Format="DateTime" />'
      '<Field Type="Text" DisplayName="RequestedBy" Name="RequestedBy" StaticName="RequestedBy" MaxLength="255" />'
    )
    Views = @(
      [ordered]@{ Title = 'Alle Aktionen'; Default = $true; RowLimit = 100
        Fields = @('Created', 'Action', 'ResourceType', 'ResourceName', 'Status', 'RequestedBy', 'TargetName', 'Reason', 'Result', 'CompletedAt')
        Query = $OrderNewest }
      [ordered]@{ Title = 'Offen'; RowLimit = 100
        Fields = @('Created', 'Action', 'ResourceType', 'ResourceName', 'Status', 'RequestedBy', 'Reason')
        Query = '<Where><Or><Eq><FieldRef Name="Status" /><Value Type="Text">Angefordert</Value></Eq><Eq><FieldRef Name="Status" /><Value Type="Text">In Arbeit</Value></Eq></Or></Where>' + $OrderNewest }
      [ordered]@{ Title = 'Fehler und Abgelehnt'; RowLimit = 100
        Fields = @('Created', 'Action', 'ResourceType', 'ResourceName', 'Status', 'RequestedBy', 'Result', 'CompletedAt')
        Query = '<Where><Or><Eq><FieldRef Name="Status" /><Value Type="Text">Fehler</Value></Eq><Eq><FieldRef Name="Status" /><Value Type="Text">Abgelehnt</Value></Eq></Or></Where>' + $OrderNewest }
      [ordered]@{ Title = 'Letzte 30 Tage'; RowLimit = 100
        Fields = @('Created', 'Action', 'ResourceType', 'ResourceName', 'Status', 'RequestedBy', 'Result')
        Query = '<Where><Geq><FieldRef Name="Created" /><Value Type="DateTime"><Today OffsetDays="-30" /></Value></Geq></Where>' + $OrderNewest }
      [ordered]@{ Title = 'Meine Aktionen'; RowLimit = 100
        Fields = @('Created', 'Action', 'ResourceType', 'ResourceName', 'Status', 'Reason', 'Result')
        Query = '<Where><Eq><FieldRef Name="Author" /><Value Type="Integer"><UserID /></Value></Eq></Where>' + $OrderNewest }
      [ordered]@{ Title = 'Nach Aktion'; RowLimit = 100
        Fields = @('Created', 'ResourceType', 'ResourceName', 'Status', 'RequestedBy', 'Result')
        Query = '<GroupBy Collapse="TRUE" GroupLimit="50"><FieldRef Name="Action" /></GroupBy>' + $OrderNewest }
    )
  }
  [ordered]@{
    Title          = 'CoE Cockpit Benachrichtigungen'
    Description    = 'Protokoll der Benachrichtigungen aus dem CoE Cockpit (welches Environment wann von wem verständigt wurde).'
    TitleField     = 'Environment'
    UniquePermissions = $false
    Versioning     = $false
    Indexes        = @('SentAt', 'Title', 'Author')
    Fields = @(
      '<Field Type="Text" DisplayName="EnvironmentId" Name="EnvironmentId" StaticName="EnvironmentId" MaxLength="80" />'
      '<Field Type="DateTime" DisplayName="SentAt" Name="SentAt" StaticName="SentAt" Format="DateTime" Indexed="TRUE" />'
      '<Field Type="Text" DisplayName="SentBy" Name="SentBy" StaticName="SentBy" MaxLength="255" />'
      '<Field Type="Text" DisplayName="SentByEmail" Name="SentByEmail" StaticName="SentByEmail" MaxLength="255" />'
      '<Field Type="Text" DisplayName="Method" Name="Method" StaticName="Method" MaxLength="80" />'
      '<Field Type="Note" DisplayName="Recipients" Name="Recipients" StaticName="Recipients" NumLines="4" RichText="FALSE" />'
      '<Field Type="Text" DisplayName="Subject" Name="Subject" StaticName="Subject" MaxLength="255" />'
    )
    Views = @(
      [ordered]@{ Title = 'Alle Benachrichtigungen'; Default = $true; RowLimit = 100
        Fields = @('LinkTitle', 'EnvironmentId', 'SentAt', 'SentBy', 'Method', 'Subject', 'Recipients')
        Query = '<OrderBy><FieldRef Name="SentAt" Ascending="FALSE" /></OrderBy>' }
      [ordered]@{ Title = 'Letzte 30 Tage'; RowLimit = 100
        Fields = @('LinkTitle', 'SentAt', 'SentBy', 'Method', 'Subject')
        Query = '<Where><Geq><FieldRef Name="SentAt" /><Value Type="DateTime"><Today OffsetDays="-30" /></Value></Geq></Where><OrderBy><FieldRef Name="SentAt" Ascending="FALSE" /></OrderBy>' }
      [ordered]@{ Title = 'Nach Environment'; RowLimit = 100
        Fields = @('SentAt', 'SentBy', 'Method', 'Subject', 'Recipients')
        Query = '<GroupBy Collapse="TRUE" GroupLimit="100"><FieldRef Name="Title" /></GroupBy><OrderBy><FieldRef Name="SentAt" Ascending="FALSE" /></OrderBy>' }
      [ordered]@{ Title = 'Meine Benachrichtigungen'; RowLimit = 100
        Fields = @('LinkTitle', 'SentAt', 'Method', 'Subject', 'Recipients')
        Query = '<Where><Eq><FieldRef Name="Author" /><Value Type="Integer"><UserID /></Value></Eq></Where><OrderBy><FieldRef Name="SentAt" Ascending="FALSE" /></OrderBy>' }
    )
  }
)

# ---------------------------------------------------------------------------------------------------------------
# Bericht
# ---------------------------------------------------------------------------------------------------------------
$Report = [System.Collections.Generic.List[object]]::new()
function Add-Result([string]$List, [string]$Element, [string]$Status, [string]$Detail = '') {
  $Report.Add([pscustomobject]@{ Liste = $List; Element = $Element; Status = $Status; Detail = $Detail })
  $color = switch ($Status) { 'OK' { 'Green' } 'ANGELEGT' { 'Cyan' } 'KORRIGIERT' { 'Cyan' } 'FEHLT' { 'Yellow' } 'WARNUNG' { 'Yellow' } default { 'Red' } }
  $text = '{0,-11} {1} – {2}' -f $Status, $List, $Element
  if ($Detail) { $text += " ($Detail)" }
  Write-Host $text -ForegroundColor $color
}

# Abfrage, die bei „nicht gefunden“ $null liefert statt abzubrechen
function Get-OrNull([scriptblock]$Query) {
  try { return (& $Query) } catch { return $null }
}

# Leerzeichen/Zeilenumbrüche für den Vergleich von CAML ignorieren
function Format-Caml([string]$Caml) {
  if (-not $Caml) { return '' }
  return (($Caml -replace '>\s+<', '><') -replace '\s+', ' ').Trim()
}

# ---------------------------------------------------------------------------------------------------------------
# Liste
# ---------------------------------------------------------------------------------------------------------------
function Confirm-List($Def) {
  $title = $Def.Title
  $list = Get-OrNull { Get-PnPList -Identity $title -ThrowExceptionIfListNotFound }
  if (-not $list) {
    if ($CheckOnly) { Add-Result $title 'Liste' 'FEHLT'; return $null }
    $list = New-PnPList -Title $title -Template GenericList -OnQuickLaunch:$false
    Set-PnPList -Identity $title -Description $Def.Description -EnableContentTypes $false | Out-Null
    Add-Result $title 'Liste' 'ANGELEGT'
  } else {
    Add-Result $title 'Liste' 'OK'
  }
  $list = Get-PnPList -Identity $title -Includes HasUniqueRoleAssignments, ContentTypesEnabled, EnableVersioning
  if ($Def.Versioning) {
    # Versionsverlauf: Änderungen an Protokolleinträgen bleiben nachvollziehbar
    if ($list.EnableVersioning) { Add-Result $title 'Versionsverlauf' 'OK' }
    elseif ($CheckOnly) { Add-Result $title 'Versionsverlauf' 'FEHLT' 'Änderungen an Einträgen wären nicht nachvollziehbar' }
    else {
      Set-PnPList -Identity $title -EnableVersioning $true | Out-Null
      Add-Result $title 'Versionsverlauf' 'KORRIGIERT' 'eingeschaltet'
    }
  }
  return $list
}

# ---------------------------------------------------------------------------------------------------------------
# Spalten
# ---------------------------------------------------------------------------------------------------------------
function Confirm-Fields($Def) {
  $title = $Def.Title
  # Titelspalte: Anzeigename (intern bleibt „Title“)
  $titleField = Get-PnPField -List $title -Identity 'Title'
  if ($titleField.Title -ne $Def.TitleField) {
    if ($CheckOnly) { Add-Result $title 'Spalte Title' 'WARNUNG' "Anzeigename '$($titleField.Title)' statt '$($Def.TitleField)'" }
    else {
      Set-PnPField -List $title -Identity 'Title' -Values @{ Title = $Def.TitleField } | Out-Null
      Add-Result $title 'Spalte Title' 'KORRIGIERT' "Anzeigename '$($Def.TitleField)'"
    }
  } else { Add-Result $title 'Spalte Title' 'OK' }

  foreach ($xml in $Def.Fields) {
    $want = ([xml]$xml).DocumentElement
    $name = $want.GetAttribute('Name')
    $type = $want.GetAttribute('Type')
    $field = Get-OrNull { Get-PnPField -List $title -Identity $name }
    if (-not $field) {
      if ($CheckOnly) { Add-Result $title "Spalte $name" 'FEHLT' $type; continue }
      Add-PnPFieldFromXml -List $title -FieldXml $xml | Out-Null
      Add-Result $title "Spalte $name" 'ANGELEGT' $type
      continue
    }
    if ($field.TypeAsString -ne $type) {
      Add-Result $title "Spalte $name" 'FEHLER' "Typ '$($field.TypeAsString)' statt '$type' – bitte von Hand prüfen (Daten!)"
      continue
    }
    # Attribute vergleichen (Anzeigename, Länge, Format, Rich Text, Zeilen)
    $have = ([xml]$field.SchemaXml).DocumentElement
    $diff = @()
    foreach ($attr in 'DisplayName', 'MaxLength', 'Format', 'RichText', 'NumLines') {
      $w = $want.GetAttribute($attr)
      if (-not $w) { continue }
      $h = $have.GetAttribute($attr)
      if ([string]$h -ne [string]$w) { $diff += "$attr=$w (ist: $(if ($h) { $h } else { 'leer' }))"; $have.SetAttribute($attr, $w) }
    }
    if (-not $diff) { Add-Result $title "Spalte $name" 'OK'; continue }
    if ($CheckOnly) { Add-Result $title "Spalte $name" 'WARNUNG' ($diff -join ', '); continue }
    $field.SchemaXml = $have.OuterXml
    $field.Update()
    Invoke-PnPQuery
    Add-Result $title "Spalte $name" 'KORRIGIERT' ($diff -join ', ')
  }
}

# ---------------------------------------------------------------------------------------------------------------
# Indizes (Schwellenwert für Listenansichten: 5.000 Einträge)
# ---------------------------------------------------------------------------------------------------------------
function Confirm-Indexes($Def) {
  $title = $Def.Title
  foreach ($n in $Def.Indexes) {
    $field = Get-OrNull { Get-PnPField -List $title -Identity $n }
    if (-not $field) { Add-Result $title "Index $n" 'FEHLT' 'Spalte fehlt'; continue }
    if ($field.Indexed) { Add-Result $title "Index $n" 'OK'; continue }
    if ($CheckOnly) { Add-Result $title "Index $n" 'FEHLT' 'nötig ab 5.000 Einträgen'; continue }
    Set-PnPField -List $title -Identity $n -Values @{ Indexed = $true } | Out-Null
    Add-Result $title "Index $n" 'ANGELEGT'
  }
}

# ---------------------------------------------------------------------------------------------------------------
# Ansichten
# ---------------------------------------------------------------------------------------------------------------
function Confirm-Views($Def) {
  $title = $Def.Title
  foreach ($v in $Def.Views) {
    $isDefault = [bool]$v.Default
    $view = Get-OrNull { Get-PnPView -List $title -Identity $v.Title -Includes ViewFields, ViewQuery, RowLimit, DefaultView, Paged }
    if (-not $view) {
      if ($CheckOnly) { Add-Result $title "Ansicht '$($v.Title)'" 'FEHLT'; continue }
      $params = @{ List = $title; Title = $v.Title; Fields = $v.Fields; Query = $v.Query; RowLimit = $v.RowLimit; Paged = $true }
      if ($isDefault) { $params.SetAsDefault = $true }
      Add-PnPView @params | Out-Null
      Add-Result $title "Ansicht '$($v.Title)'" 'ANGELEGT' ($(if ($isDefault) { 'Standardansicht' } else { '' }))
      continue
    }
    $diff = @()
    $haveFields = @($view.ViewFields)
    if (($haveFields -join '|') -ne ($v.Fields -join '|')) { $diff += 'Spalten' }
    if ((Format-Caml $view.ViewQuery) -ne (Format-Caml $v.Query)) { $diff += 'Filter/Sortierung' }
    if ([int]$view.RowLimit -ne [int]$v.RowLimit) { $diff += "Zeilen $($view.RowLimit) → $($v.RowLimit)" }
    if ($isDefault -and -not $view.DefaultView) { $diff += 'nicht Standardansicht' }
    if (-not $diff) { Add-Result $title "Ansicht '$($v.Title)'" 'OK'; continue }
    if ($CheckOnly) { Add-Result $title "Ansicht '$($v.Title)'" 'WARNUNG' ($diff -join ', '); continue }
    $values = @{ ViewQuery = $v.Query; RowLimit = [uint32]$v.RowLimit; Paged = $true }
    if ($isDefault) { $values.DefaultView = $true }
    Set-PnPView -List $title -Identity $view.Id -Fields $v.Fields -Values $values | Out-Null
    Add-Result $title "Ansicht '$($v.Title)'" 'KORRIGIERT' ($diff -join ', ')
  }
}

# ---------------------------------------------------------------------------------------------------------------
# Berechtigungen der Aktionsliste
# ---------------------------------------------------------------------------------------------------------------
function Confirm-Permissions($Def, $List) {
  $title = $Def.Title
  # Rollen über ihren Typ suchen (Namen sind je nach Sprache verschieden: „Mitwirken“/„Contribute“, „Vollzugriff“/„Full Control“)
  $roles = Get-PnPRoleDefinition
  $contribute = ($roles | Where-Object { $_.RoleTypeKind -eq 'Contributor' } | Select-Object -First 1).Name
  $full = ($roles | Where-Object { $_.RoleTypeKind -eq 'Administrator' } | Select-Object -First 1).Name

  $group = Get-OrNull { Get-PnPGroup -Identity $GroupName }
  if (-not $group) {
    if ($CheckOnly) { Add-Result $title "Gruppe '$GroupName'" 'FEHLT' }
    else {
      $group = New-PnPGroup -Title $GroupName -Description 'Darf im CoE Cockpit Aktionen ausführen (zusätzlich aktive Rolle „Power Platform Administrator“ nötig).'
      Add-Result $title "Gruppe '$GroupName'" 'ANGELEGT' 'Mitglieder hinzufügen!'
    }
  } else { Add-Result $title "Gruppe '$GroupName'" 'OK' }

  # Der Flow „Solution löschen“ liest die Gruppenmitglieder – das geht nur, wenn die Mitgliedschaft für alle sichtbar ist
  # (SharePoint-Standard bei neuen Gruppen: nur Mitglieder sehen die Mitglieder).
  if ($group) {
    $onlyMembers = $true
    try { $onlyMembers = [bool](Get-PnPProperty -ClientObject $group -Property OnlyAllowMembersViewMembership) }
    catch { $onlyMembers = [bool]$group.OnlyAllowMembersViewMembership }
    if (-not $onlyMembers) { Add-Result $title 'Mitgliedschaft der Gruppe für alle sichtbar' 'OK' }
    elseif ($CheckOnly) { Add-Result $title 'Mitgliedschaft der Gruppe für alle sichtbar' 'FEHLT' 'der Flow kann die Mitglieder sonst nicht prüfen' }
    else {
      Set-PnPGroup -Identity $group.Title -OnlyAllowMembersViewMembership $false | Out-Null
      Add-Result $title 'Mitgliedschaft der Gruppe für alle sichtbar' 'KORRIGIERT' 'damit der Flow die Mitglieder prüfen kann'
    }
  }

  if (-not $List.HasUniqueRoleAssignments) {
    if ($CheckOnly) { Add-Result $title 'Eigene Berechtigungen' 'FEHLT' 'Liste erbt die Berechtigungen der Website' }
    else {
      Set-PnPList -Identity $title -BreakRoleInheritance -CopyRoleAssignments:$false -ClearSubscopes | Out-Null
      Add-Result $title 'Eigene Berechtigungen' 'KORRIGIERT' 'Vererbung aufgehoben'
    }
  } else { Add-Result $title 'Eigene Berechtigungen' 'OK' }

  $owners = Get-PnPGroup -AssociatedOwnerGroup
  $want = @()
  if ($group) { $want += [pscustomobject]@{ Label = "Gruppe '$GroupName'"; Group = $group.Title; User = $null; Role = $contribute; Id = $group.Id } }
  if ($owners) { $want += [pscustomobject]@{ Label = "Websitebesitzer '$($owners.Title)'"; Group = $owners.Title; User = $null; Role = $full; Id = $owners.Id } }
  if ($FlowServiceAccount) {
    if ($CheckOnly) {
      # nur nachsehen – New-PnPUser (EnsureUser) würde das Konto in die Website eintragen
      $svc = Get-OrNull { Get-PnPUser | Where-Object { $_.Email -eq $FlowServiceAccount -or $_.LoginName -like "*|$FlowServiceAccount" } | Select-Object -First 1 }
    } else {
      $svc = Get-OrNull { New-PnPUser -LoginName $FlowServiceAccount }
    }
    if ($svc) { $want += [pscustomobject]@{ Label = "Flow-Dienstkonto $FlowServiceAccount"; Group = $null; User = $FlowServiceAccount; Role = $contribute; Id = $svc.Id } }
    elseif ($CheckOnly) { Add-Result $title "Flow-Dienstkonto $FlowServiceAccount" 'FEHLT' 'noch nicht auf der Website bzw. ohne Berechtigung' }
    else { Add-Result $title "Flow-Dienstkonto $FlowServiceAccount" 'FEHLER' 'Konto nicht gefunden' }
  }
  foreach ($w in $want) {
    $has = @()
    try { $has = @(Get-PnPListPermissions -Identity $title -PrincipalId $w.Id | ForEach-Object { $_.Name }) } catch { $has = @() }
    if ($has -contains $w.Role) { Add-Result $title $w.Label 'OK' $w.Role; continue }
    if ($CheckOnly) { Add-Result $title $w.Label 'FEHLT' $w.Role; continue }
    if ($w.Group) { Set-PnPListPermission -Identity $title -Group $w.Group -AddRole $w.Role | Out-Null }
    else { Set-PnPListPermission -Identity $title -User $w.User -AddRole $w.Role | Out-Null }
    Add-Result $title $w.Label 'KORRIGIERT' "$($w.Role) vergeben"
  }

  # weitere Berechtigte nur melden (nicht automatisch entfernen)
  $allowed = @($want | ForEach-Object { $_.Id })
  $ctx = Get-PnPContext
  # die ausführende Person bekommt beim Aufheben der Vererbung automatisch Vollzugriff – kein Fehler
  $me = $ctx.Web.CurrentUser
  $ctx.Load($me)
  $ctx.ExecuteQuery()
  $listObj = Get-PnPList -Identity $title
  $assignments = $listObj.RoleAssignments
  $ctx.Load($assignments)
  $ctx.ExecuteQuery()
  foreach ($a in $assignments) {
    $ctx.Load($a.Member)
    $ctx.ExecuteQuery()
    if ($a.PrincipalId -eq $me.Id) {
      Add-Result $title "Ausführende Person: $($a.Member.Title)" 'OK' 'Vollzugriff durch das Aufheben der Vererbung – kann bleiben oder von Hand entfernt werden'
    } elseif ($allowed -notcontains $a.PrincipalId) {
      Add-Result $title "Weitere Berechtigung: $($a.Member.Title)" 'WARNUNG' 'nicht vorgesehen – bitte prüfen (Flow-Dienstkonto? Dann -FlowServiceAccount angeben)'
    }
  }
}

# ---------------------------------------------------------------------------------------------------------------
# Ablauf
# ---------------------------------------------------------------------------------------------------------------
Write-Host ''
Write-Host "CoE Cockpit – SharePoint-Listen auf $SiteUrl $(if ($CheckOnly) { '(nur prüfen)' } else { '(prüfen und ergänzen)' })" -ForegroundColor White
Write-Host ''
foreach ($def in $Lists) {
  try {
    $list = Confirm-List $def
    if (-not $list) { continue }
    Confirm-Fields $def
    Confirm-Indexes $def
    Confirm-Views $def
    if ($def.UniquePermissions) { Confirm-Permissions $def $list }
  } catch {
    Add-Result $def.Title 'Prüfung' 'FEHLER' $_.Exception.Message
  }
}

Write-Host ''
$Report | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ('{0,-11} {1}' -f $_.Name, $_.Count) }
$problems = @($Report | Where-Object { $_.Status -in 'FEHLER', 'FEHLT', 'WARNUNG' })
if ($problems.Count) {
  Write-Host ''
  Write-Host 'Offene Punkte:' -ForegroundColor Yellow
  # als Text ausgeben – sonst geht die Tabelle durch das „exit“ unten verloren
  $problems | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
}
if (@($Report | Where-Object { $_.Status -eq 'FEHLER' }).Count) { exit 2 }
if ($CheckOnly -and $problems.Count) { exit 1 }
exit 0
