# CoE Cockpit – SharePoint-Webpart

Zeigt die Dataverse-Tabellen des **Power Platform CoE Starter Kits** direkt in SharePoint an:
Apps, Flows, Environments und Maker sind vorkonfiguriert, weitere Tabellen lassen sich im Webpart über einen Editor hinzufügen.

- SPFx 1.21.1, React 17, Fluent UI 8
- Zugriff über `AadHttpClient` auf die Dataverse Web API (v9.2) mit den Rechten des angemeldeten Benutzers
- Passt sich automatisch an SharePoint-Theme und Abschnittsfarben an (auch dunkle Abschnitte)

## Funktionen

Nachbau der CoE-App **„Power Platform Admin View“** (CoE Starter Kit, Core Components) als SharePoint-Webpart – **nur lesend**,
alle Daten direkt aus den Dataverse-Tabellen des CoE (keine Power-Platform-Admin-APIs).

| Bereich | Beschreibung |
|---|---|
| Übersicht (Startseite) | Dashboard mit Gesamtzahl der Ressourcen, Countdown zur Sperre nicht verwalteter Developer-Environments (15.12.2026), Kennzahl-Kacheln mit 12-Monats-Verlauf, Verlauf „Neu erstellt pro Monat“, **Handlungsbedarf** (betroffene Developer-Environments, gesperrte Flows, verwaiste Apps/Flows, seit 90 Tagen ungenutzte Apps, Gast-Maker), Verteilungen (Apps nach Typ, Flows nach Status, Environments nach Typ) und Ranglisten (meistgenutzte Apps, größte Environments, Top-Connectors, aktivste Maker). Jedes Element führt per Klick in die gefilterte Liste oder öffnet den Datensatz. Nur Zählungen/Gruppierungen – auch bei sehr großen Tabellen schnell. |
| Navigation | Seitenleiste wie in der Admin View: **Inventar** (Apps, Flows, Environments, Connectors, Bots, Desktop Flows, Power Pages Sites, Business Process Flows, Solutions, AI-Modelle), **Governance** (DLP-Richtlinien), **Benutzer** (Maker, Power Platform User, Benutzerrollen, Sicherheitsrollen, Verbindungsidentitäten), **Nutzung** (Environment-Kapazität, Audit-Log), **Auswertungen** (Default Environment, Betroffen ab 15.12., Sicherheit & Freigaben), **ALM-Pipelines** (Bereitstellungen, Pipelines, Pipeline-Environments). Auf schmalen Seiten waagrecht. |
| Suche über alle Bereiche | Suchfeld im Kopf: sucht gleichzeitig in Apps, Flows, Environments, Makern, Connectors, Bots, Solutions, DLP-Richtlinien, Pipelines … (direkt in Dataverse, nicht nur in geladenen Daten). Treffer öffnen die Detailansicht, „Alle anzeigen“ springt in die Liste mit dem Suchbegriff. |
| Sicherheit & Freigaben | Risiken auf einen Blick: Apps für die gesamte Organisation, Apps/Flows an Gäste freigegeben, Custom Connectors über unverschlüsseltes HTTP, Environments ohne Sicherheitsgruppe bzw. ohne DLP-Richtlinie, deaktivierte Konten, Gast-Maker, Apps mit DLP-Verstößen, Environments mit den meisten System-Administratoren, DLP-Richtlinien. Jede Kachel führt in die gefilterte Liste. |
| DLP-Richtlinien | Alle Richtlinien aus dem CoE-Inventar mit Geltungsbereich; in der Detailansicht die zugeordneten **Environments**, **blockierten** und **Business-Connectors** |
| Maker 360° | Detailansicht eines Makers mit Apps, Flows, Bots, Desktop Flows, **Zugriff auf** (alle Freigaben/Rollen der Person), **Sicherheitsrollen** in allen Environments und eigenen Environments |
| ALM-Pipelines | Bereitstellungen (Solution, Version, Pipeline, Stufe, Status, ausgelöst von), Pipelines mit Stufen und Pipeline-Environments – live aus dem Dataverse des **Pipelines-Host-Environments** (Auswahl in den Eigenschaften) |
| Hinweis Adminrolle | Bereiche außerhalb des CoE-Environments (Solutions im Default Environment, ALM-Pipelines) werden beim Start geprüft. Fehlt der Zugriff, erscheint **vorab** ein Hinweis, welche Bereiche nicht sichtbar sind und dass die Rolle „Power Platform-Administrator“ (PIM) aktiviert werden muss; die Einträge sind in der Navigation mit einem Schloss markiert. |
| Tabellenansicht | Volltextsuche, Sortierung, Seitenumbruch, kompakte/normale Zeilenhöhe, Spaltenauswahl, CSV-Export für Excel |
| Spaltenfilter | Filter-Symbol in **jedem** Spaltenkopf: Werteliste mit Suche und Anzahl (wie in Excel), bei Datumsspalten Zeiträume. Aktive Filter als Chips, einzeln entfernbar. |
| Detailansicht | Seitenpanel mit Eigenschaften, Direktlinks (Power Apps, Power Automate, Admin Center, Dataverse), „Alle Felder“ und Reitern wie in den Admin-View-Formularen: **Freigegeben für**, **Verbindungen**, **Nutzung**, **Aktionen**, Ressourcen eines Environments (Apps, Flows, Bots, Desktop Flows, Power Pages, Solutions, Kapazität, Sicherheitsrollen), **Verwendet in Apps/Flows** … Zeilen sind anklickbar (Durchklicken), „Zurück“ führt zum vorherigen Datensatz. |
| Betroffen ab 15.12. | Alle **nicht verwalteten Developer-Environments** (Sperre 15.12.2026, Löschung ab 15.01.2027) mit Ersteller, Zustand, Anzahl Apps/Flows. Managed-Status aus „Governance Protection Level“ („Basic“ = nicht managed), per Auswahl in den Eigenschaften änderbar. |
| E-Mail an Maker | Im Reiter „Betroffen ab 15.12.“ öffnet das **Brief-Symbol** am Zeilenanfang (oder die Schaltfläche in der Detailansicht) eine **Vorschau der formatierten E-Mail** an alle Maker des Environments: Kopfbalken in der Theme-Farbe, Fristen-Kacheln (Rückmeldung, Sperre 15.12.2026, Löschung 15.01.2027), Steckbrief mit **Link zum Environment in make.powerapps.com**, Schaltfläche „Environment in Power Apps öffnen“, nummerierte **Rückmelde-Optionen** (1 = weiter benötigt → Managed + Premium-Lizenz, 2 = löschen, 3 = umziehen; optional Formular), Tipp zum Sichern (Solutions exportieren) und die **offizielle Microsoft-Ankündigung** (Abschnitt „Deprecation of unmanaged developer environments …“ auf Microsoft Learn). „In Outlook öffnen“ lädt die fertige Mail als Outlook-Entwurf (.eml) herunter – Datei öffnen, Outlook zeigt sie mit Empfängern, Betreff und formatiertem Text. Alternativ „Über Zwischenablage“ (Empfänger/Betreff per Link, Text mit Strg+V). Die Mail ist „Outlook-fest“ aufgebaut (Tabellen, Farben direkt am Text, keine runden Elemente), damit sie auch im klassischen Outlook sauber aussieht. Mit **„Direkt senden“** geht die Mail ohne Outlook über Microsoft Graph aus dem eigenen Postfach. **„Alle gefilterten verständigen“** schickt jedem gefilterten Environment eine **eigene E-Mail** – automatisch (mit Sicherheitsabfrage, Fortschritt, Anhalten, Fehlerliste, ca. 28 Mails pro Minute wegen des Exchange-Limits) oder **nacheinander in Outlook** (Schritt für Schritt). **Gemeinsames Versandprotokoll:** Jede Benachrichtigung (automatisch gesendet oder in Outlook gesendet und „Als versendet markiert“) wird in der SharePoint-Liste **„CoE Cockpit Benachrichtigungen“** auf der Website des Cockpits vermerkt – mit Environment, Zeitpunkt, Person, Art, Empfängern und Betreff. Alle Admins sehen in der Tabelle die Spalten „Maker verständigt / Verständigt am / Verständigt von“ und die Kachel „Noch nicht verständigt“; wer ein bereits verständigtes Environment erneut anschreiben will, bekommt eine deutliche Warnung. Beim Serienversand wird das Protokoll unmittelbar vor dem Senden neu gelesen – was ein anderer Admin inzwischen verständigt hat, wird übersprungen. „E-Mail-Adressen kopieren“ liefert die Adressen aller gefilterten Environments. Anwendungsbenutzer („# …“) und SYSTEM werden nicht angeschrieben. |
| Default Environment | Apps und Flows aus dem CoE-Inventar, Solutions live aus dem Dataverse des Default Environments |
| Tabellen-Editor | Tabellen **aus einer Liste wählen** (alle Admin-View-Tabellen als Vorlagen oder jede andere Tabelle der Umgebung), Spalten aus Dataverse auswählen, JSON-Import/-Export |
| Robustheit | Spalten, die es in der eigenen CoE-Version nicht gibt, werden übersprungen und als Hinweis angezeigt |

## Fertiges Paket & Vorschau

- **`sppkg/coe-cockpit.sppkg`** – fertig gebautes Paket für den App Catalog (siehe „Bereitstellung“).
- **`preview/coe-cockpit-vorschau.html`** – Vorschau im Browser mit Beispieldaten (über 5.000 Apps/Flows).
  Verwendet den echten Webpart-Code und **das vom SPFx-Build kompilierte CSS** (wie in SharePoint); nur die Dataverse-API ist simuliert.
  Neu bauen: im Projekt `npx gulp bundle`, danach `cd preview && npm install && npm run build`.

## Große Tabellen (mehr als 5.000 Einträge)

- Kennzahlen werden per OData-Aggregation **exakt bis 50.000** gezählt (Dataverse-Grenze für Aggregate);
  darüber wird „50.000+“ angezeigt. Unterstützt eine Umgebung keine Aggregation, wird auf `$count` (max. 5.000) zurückgefallen.
- Datensätze werden seitenweise (je 5.000) geladen und schon während des Ladens angezeigt.
- Ladelimit pro Tabelle: Standard 20.000, einstellbar bis 50.000. Ist eine Tabelle größer, zeigt das Webpart einen Hinweis;
  Suche und Spaltenfilter wirken dann nur auf die geladenen Einträge (nach Standardsortierung, z. B. zuletzt geändert).

## Build

### Azure DevOps Pipeline (empfohlen)

1. Projektordner in ein Azure-DevOps-Repo legen und pushen.
2. Neue Pipeline anlegen → „Existing Azure Pipelines YAML file“ → `azure-pipelines.yml`.
3. Pipeline ausführen. Die fertige `coe-cockpit.sppkg` liegt danach unter **Artifacts → sppkg**.

### Lokal

Voraussetzung: Node.js 22 (LTS).

```powershell
npm install
npx gulp bundle --ship
npx gulp package-solution --ship
```

Ergebnis: `sharepoint/solution/coe-cockpit.sppkg`

## Bereitstellung

1. `coe-cockpit.sppkg` in den **App Catalog** hochladen und „Bereitstellen“ wählen
   (Option „Für alle Websites verfügbar machen“ ist möglich, da `skipFeatureDeployment` aktiv ist).
2. **API-Zugriff genehmigen:** SharePoint Admin Center → *Erweitert* → *API-Zugriff* →
   Anforderung **Dynamics CRM – user_impersonation** genehmigen. Ohne diese Genehmigung kann das Webpart keine Daten laden.
   Für den automatischen E-Mail-Versand zusätzlich **Microsoft Graph – Mail.Send** genehmigen (optional; ohne diese
   Genehmigung funktioniert „In Outlook öffnen“ bzw. „Nacheinander in Outlook“ weiterhin). Achtung: API-Genehmigungen für
   SharePoint-Webparts gelten für alle SPFx-Lösungen im Tenant.
3. Webpart **CoE Cockpit** auf einer Seite einfügen und in den Eigenschaften die **URL der CoE-Umgebung** eintragen,
   z. B. `https://contoso-coe.crm4.dynamics.com`.

## Berechtigungen

Das Webpart liest mit den Rechten des angemeldeten Benutzers. Benutzer brauchen in der CoE-Umgebung eine Sicherheitsrolle
mit Lesezugriff auf die CoE-Tabellen (z. B. **Power Platform Admin SR** aus dem CoE Starter Kit oder System Administrator).
Benutzer ohne Rechte sehen eine verständliche Fehlermeldung statt Daten.

## Versandprotokoll (SharePoint-Liste)

Die Liste **„CoE Cockpit Benachrichtigungen“** legt das Cockpit beim ersten Vermerk selbst auf der Website an, auf der es eingebunden ist
(Spalten: Title = Environment, EnvironmentId, SentAt, SentBy, SentByEmail, Method, Recipients, Subject). Dafür braucht die Person beim
ersten Mal **Besitzer-Rechte** auf der Website; danach reichen **Bearbeiten-Rechte** auf der Liste. Es wird keine zusätzliche API-Berechtigung
benötigt (Zugriff mit den SharePoint-Rechten des angemeldeten Benutzers).

## Bereiche außerhalb des CoE (Adminrolle)

Fast alles liest das Cockpit aus dem CoE-Environment. Zwei Bereiche liegen in anderen Environments und werden dort mit den
Rechten des angemeldeten Benutzers gelesen:

| Bereich | Environment | Benötigt |
|---|---|---|
| Default Environment → Solutions | Default Environment | Leserechte auf `solution` (z. B. Environment Maker) |
| ALM-Pipelines | Pipelines-Host-Environment | Leserechte auf die Pipelines-Tabellen (z. B. System Administrator, Deployment Pipeline Administrator) |

Power-Platform-Admins aktivieren dafür die Rolle **Power Platform-Administrator** (Entra PIM) und übernehmen – falls nötig –
im Power Platform Admin Center beim Environment unter *Mitgliedschaft* → *Mich hinzufügen* die Rolle System Administrator.
Das Cockpit prüft den Zugriff beim Start und zeigt vorab, welche Bereiche gerade nicht sichtbar sind.

## Default Environment – Hinweise

- Apps und Flows werden über die Umgebungs-Verknüpfung im CoE-Inventar gefiltert (`_admin_appenvironment_value` bzw. `_admin_flowenvironment_value`).
- Solutions werden direkt aus dem Dataverse des Default Environments gelesen (Tabelle `solution`, nur sichtbare Solutions).
  Dafür braucht der Benutzer dort Leserechte (z. B. Environment Maker). Hat das Default Environment keine Dataverse-Datenbank,
  zeigt das Webpart einen entsprechenden Hinweis.
- Die API-Berechtigung „Dynamics CRM“ gilt für alle Dataverse-Umgebungen – es ist keine zusätzliche Genehmigung nötig.

## Konfiguration

| Eigenschaft | Beschreibung |
|---|---|
| URL der CoE-Umgebung | Dataverse-URL der Umgebung mit dem CoE Starter Kit |
| Übersicht (Dashboard) als Startseite anzeigen | Übersicht ein-/ausblenden |
| Bereich „Default Environment“ anzeigen | Reiter ein-/ausblenden |
| Reiter „Betroffen ab 15.12.“ anzeigen | Reiter ein-/ausblenden |
| Ansicht „Sicherheit & Freigaben“ anzeigen | Ansicht ein-/ausblenden |
| E-Mail an Maker: Automatischen Versand anbieten | Schaltflächen „Direkt senden“ und „Automatisch senden“ (Microsoft Graph, Mail.Send) ein-/ausblenden |
| E-Mail an Maker: Anrede | Sie oder Du |
| E-Mail an Maker: Frist für die Rückmeldung | 7, 10, 14 oder 21 Tage (spätestens 3 Tage vor dem Stichtag) |
| E-Mail an Maker: Link zur Microsoft-Information | Leer = offizielle Ankündigung: https://learn.microsoft.com/en-us/power-platform/important-changes-coming#deprecation-of-unmanaged-developer-environments-and-the-power-apps-developer-plan |
| E-Mail an Maker: Formular (optional) | z. B. Microsoft Forms; `{environment}` im Link wird durch den Namen ersetzt. Ohne Formular antworten die Maker per E-Mail. |
| E-Mail an Maker: CC / Grußzeile | z. B. Team-Postfach in CC, damit Rückmeldungen dort ankommen |
| Pipelines-Host-Environment | Auswahlliste mit allen Environments mit Dataverse aus dem CoE-Inventar; „Keine“ blendet die ALM-Pipelines aus |
| Spalte mit dem Managed-Status | Auswahlliste mit den Spalten der CoE-Environment-Tabelle. Standard: **automatisch erkennen** |
| Default Environment | Auswahlliste mit allen Umgebungen aus dem CoE-Inventar. Standard: **automatisch erkennen** (Umgebungstyp „Default“) |
| Tabellen | Über den Tabellen-Editor (Button im Eigenschaftenbereich oder „Tabellen verwalten“ im Bearbeitungsmodus) |
| Titel / Untertitel | Überschrift des Webparts |
| Kennzahlen anzeigen | Kennzahlen-Leiste ein-/ausblenden |
| Zeilenhöhe | Normal oder kompakt (Benutzer können zusätzlich selbst umschalten) |
| Einträge pro Seite | 10, 25, 50 oder 100 |
| Maximale Anzahl geladener Einträge | 1.000 – 50.000 pro Tabelle (Standard 20.000) |

### Tabellen-Konfiguration (JSON)

Im Tabellen-Editor gibt es einen Reiter **JSON Import/Export**. Beispiel für eine zusätzliche Tabelle:

```json
{
  "key": "solutions",
  "title": "Solutions",
  "logicalName": "admin_solution",
  "icon": "Package",
  "showInKpis": true,
  "columns": [
    { "field": "admin_displayname", "label": "Solution", "type": "title" },
    { "field": "_admin_environment_value", "label": "Umgebung", "type": "text" },
    { "field": "modifiedon", "label": "Geändert", "type": "date" }
  ],
  "links": []
}
```

Spaltentypen: `title`, `text`, `person`, `status`, `choice`, `boolean`, `date`, `datetime`, `number`, `url`.
Lookup-Spalten werden mit `_<feldname>_value` angegeben; das Webpart zeigt automatisch den Anzeigenamen des verknüpften Datensatzes.
Status-Farben: `success`, `warning`, `danger`, `info`, `accent`, `neutral`.

## Projektstruktur

```
src/webparts/coeCockpit/
├─ CoeCockpitWebPart.ts          Webpart, Eigenschaftenbereich, Theme
├─ components/
│  ├─ CoeCockpit.tsx             Hauptansicht (Kopf, Kennzahlen, Reiter, Toolbar, Tabelle)
│  ├─ OverviewView.tsx           Übersicht (Dashboard)
│  ├─ SecurityView.tsx           Sicherheit & Freigaben
│  ├─ GlobalSearch.tsx           Suche über alle Bereiche
│  ├─ Charts.tsx                 Diagramme (Verlauf, Balken, Sparklines)
│  ├─ CoeCockpit.module.scss     Design-System (CSS-Variablen aus dem Theme)
│  ├─ DataTable.tsx              Tabelle mit Sortierung und Ladezustand
│  ├─ Cells.tsx                  Zellen-Darstellung (Status, Person, Datum …)
│  ├─ KpiStrip.tsx               Kennzahlen-Kacheln
│  ├─ RecordPanel.tsx            Detailansicht
│  ├─ TableEditor.tsx            Tabellen-Editor
│  └─ themeVars.ts               Theme → CSS-Variablen
├─ models/                       Konfigurationsmodell und Standardtabellen
├─ services/DataverseService.ts  Dataverse Web API (Metadaten, Daten, Zählung)
└─ utils/format.ts               Formatierung, CSV, Hilfsfunktionen
```

## Nächste Ausbaustufe

Admin-Aktionen direkt aus der Detailansicht (Besitzer ändern, Flows aktivieren/deaktivieren, Freigaben verwalten,
Umgebungseinstellungen). Dafür werden zusätzlich die Power-Platform-Admin-APIs bzw. ein Power Automate Flow als Backend benötigt.
