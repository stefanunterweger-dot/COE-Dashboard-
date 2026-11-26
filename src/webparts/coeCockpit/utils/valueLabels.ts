/**
 * Deutsche Anzeige der (englischen) Werte des CoE Starter Kits – nur im Webpart.
 * Die Daten in Dataverse bleiben unverändert; Abfragen ($filter) verwenden weiter die Originalwerte.
 * Die Übersetzung hängt an der Spalte, weil derselbe Wert je nach Spalte etwas anderes bedeutet
 * (z. B. „Standard“ = Managed bei der Schutzstufe, aber Standard-Connector bei der Stufe).
 */

type Labels = { [value: string]: string };

const YES_NO: Labels = { 'yes': 'Ja', 'no': 'Nein', 'true': 'Ja', 'false': 'Nein' };
const PUBLISH: Labels = { 'published': 'Veröffentlicht', 'unpublished': 'Nicht veröffentlicht', 'deleted': 'Gelöscht', 'deleted unpublished': 'Gelöscht (nicht veröffentlicht)', 'provisioned': 'Bereitgestellt', 'trained': 'Trainiert', 'draft': 'Entwurf' };

const FIELD_LABELS: { [field: string]: Labels } = {
  cr5d5_appisorphaned: YES_NO,
  cr5d5_flowisorphaned: YES_NO,
  admin_appisorphaned: YES_NO,
  admin_pvaisorphaned: YES_NO,
  admin_flowstate: { 'started': 'Aktiv', 'stopped': 'Angehalten', 'suspended': 'Gesperrt' },
  admin_governanceprotectionlevel: { 'basic': 'Nicht verwaltet', 'standard': 'Managed' },
  admin_environmentruntimestate: { 'enabled': 'Aktiviert', 'disabled': 'Deaktiviert' },
  admin_environmentstate: { 'ready': 'Bereit', 'disabled': 'Deaktiviert', 'deleted': 'Gelöscht', 'notspecified': 'Nicht angegeben' },
  admin_environmentsku: { 'production': 'Produktion', 'sandbox': 'Sandbox', 'developer': 'Developer', 'default': 'Default', 'trial': 'Testversion', 'subscriptionbasedtrial': 'Testversion (Abo)', 'teams': 'Teams', 'support': 'Support' },
  admin_rolename: { 'owner': 'Besitzer', 'canedit': 'Mitbesitzer', 'canview': 'Benutzer', 'canviewwithcomponentimport': 'Benutzer (mit Komponentenimport)' },
  admin_type: { 'user': 'Benutzer', 'group': 'Gruppe', 'tenant': 'Gesamte Organisation', 'serviceprincipal': 'Dienstprinzipal', 'system': 'System' },
  admin_pvastate: PUBLISH,
  admin_rpastate: PUBLISH,
  admin_aistatus: PUBLISH,
  admin_businessprocessflowstatus: { 'activated': 'Aktiviert', 'draft': 'Entwurf' },
  admin_portalwebsiterecordstatus: { 'active': 'Aktiv', 'inactive': 'Inaktiv' },
  admin_tier: { 'standard': 'Standard', 'premium': 'Premium', 'custom': 'Benutzerdefiniert' },
  admin_powerappstype: {
    'canvas': 'Canvas', 'model driven': 'Modellgesteuert', 'sharepointformapp': 'SharePoint-Formular',
    'custom page': 'Benutzerdefinierte Seite', 'component library': 'Komponentenbibliothek'
  },
  admin_usertype: { 'member': 'Mitglied', 'guest': 'Gast', 'app user': 'App-Benutzer', 'unknown': 'Unbekannt', 'microsoft account': 'Microsoft-Konto' },
  admin_operation: { 'launchpowerapp': 'App gestartet' },
  admin_resourcetype: { 'app': 'App', 'flow': 'Flow', 'environment': 'Environment' }
};

/** Deutscher Anzeigewert für einen CoE-Wert (unbekannte Werte bleiben unverändert). */
export function translateValue(field: string, value: string): string {
  const map = FIELD_LABELS[field];
  if (!map || !value) {
    return value;
  }
  const hit = map[value.trim().toLowerCase()];
  return hit !== undefined ? hit : value;
}
