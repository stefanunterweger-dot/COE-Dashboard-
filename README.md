# CoE Cockpit – SharePoint-Webpart

Zeigt die Dataverse-Tabellen des **Power Platform CoE Starter Kits** direkt in SharePoint an:
Apps, Flows, Environments und Maker sind vorkonfiguriert, weitere Tabellen lassen sich im Webpart über einen Editor hinzufügen.

- SPFx 1.21.1, React 17, Fluent UI 8
- Zugriff über `AadHttpClient` auf die Dataverse Web API (v9.2) mit den Rechten des angemeldeten Benutzers
- Passt sich automatisch an SharePoint-Theme und Abschnittsfarben an (auch dunkle Abschnitte)

## Funktionen

| Bereich | Beschreibung |
|---|---|
| Kennzahlen | Eine Kachel pro Tabelle mit Gesamtzahl und optionaler Hervorhebung (z. B. „12 verwaist“, „3 gesperrt“) |
| Tabellenansicht | Volltextsuche, Sortierung, Seitenumbruch, kompakte/normale Zeilenhöhe, Spaltenauswahl, CSV-Export für Excel |
| Spaltenfilter | **Jede Spalte** hat einen Filter im Spaltenkopf: Werteliste mit Suche und Anzahl (wie in Excel), bei Datumsspalten Zeiträume (Heute, Letzte 7/30/90 Tage, Älter als 1 Jahr …). Aktive Filter erscheinen als Chips in der Toolbar und lassen sich einzeln entfernen. Filter wirken zusammen, die Werte im Filter passen sich an die übrigen Filter an. |
| Default Environment | Eigener Reiter mit allen **Apps, Flows und Solutions** im Default Environment. Apps und Flows kommen aus dem CoE-Inventar, Solutions live aus dem Dataverse des Default Environments. Mit Infokopf (Name, Umgebungs-ID, Link ins Admin Center) und Anzahl je Bereich. |
| Detailansicht | Seitenpanel mit Eigenschaften, Status, Direktlinks (Power Apps, Power Automate, Admin Center, Dataverse) und „Alle Felder“ |
| Tabellen-Editor | Tabellen hinzufügen, Spalten aus Dataverse auswählen, Darstellung, Filter, Kennzahlen und Links festlegen, JSON-Import/-Export |
| Robustheit | Spalten, die es in der eigenen CoE-Version nicht gibt, werden automatisch übersprungen und als Hinweis angezeigt |

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
3. Webpart **CoE Cockpit** auf einer Seite einfügen und in den Eigenschaften die **URL der CoE-Umgebung** eintragen,
   z. B. `https://contoso-coe.crm4.dynamics.com`.

## Berechtigungen

Das Webpart liest mit den Rechten des angemeldeten Benutzers. Benutzer brauchen in der CoE-Umgebung eine Sicherheitsrolle
mit Lesezugriff auf die CoE-Tabellen (z. B. **Power Platform Admin SR** aus dem CoE Starter Kit oder System Administrator).
Benutzer ohne Rechte sehen eine verständliche Fehlermeldung statt Daten.

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
| Bereich „Default Environment“ anzeigen | Reiter ein-/ausblenden |
| Default Environment | Auswahlliste mit allen Umgebungen aus dem CoE-Inventar. Standard: **automatisch erkennen** (Umgebungstyp „Default“) |
| Dataverse-URL des Default Environments | Optional. Wird für die Solutions verwendet; leer = URL aus dem CoE-Inventar |
| Tabellen | Über den Tabellen-Editor (Button im Eigenschaftenbereich oder „Tabellen verwalten“ im Bearbeitungsmodus) |
| Titel / Untertitel | Überschrift des Webparts |
| Kennzahlen anzeigen | Kennzahlen-Leiste ein-/ausblenden |
| Zeilenhöhe | Normal oder kompakt (Benutzer können zusätzlich selbst umschalten) |
| Einträge pro Seite | 10, 25, 50 oder 100 |
| Maximale Anzahl geladener Einträge | 500 – 20.000 pro Tabelle |

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
