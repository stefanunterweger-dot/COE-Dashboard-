# CoE Cockpit – SharePoint-Webpart

Zeigt die Dataverse-Tabellen des **Power Platform CoE Starter Kits** direkt in SharePoint an:
Apps, Flows, Environments und Maker sind vorkonfiguriert, weitere Tabellen lassen sich im Webpart über einen Editor hinzufügen.

- SPFx 1.21.1, React 17, Fluent UI 8
- Zugriff über `AadHttpClient` auf die Dataverse Web API (v9.2) mit den Rechten des angemeldeten Benutzers
- Passt sich automatisch an SharePoint-Theme und Abschnittsfarben an (auch dunkle Abschnitte)

## Funktionen

Nachbau der CoE-App **„Power Platform Admin View“** (CoE Starter Kit, Core Components) als SharePoint-Webpart. Die Daten kommen
direkt aus den Dataverse-Tabellen des CoE; nur **Copilot-Credits / Abrechnungspläne** werden live über die Power Platform API gelesen.
Ändernde Aktionen führt das Cockpit direkt mit den Rechten der angemeldeten Person aus (aktive Rolle „Power Platform Administrator“) und
protokolliert jede in einer SharePoint-Liste; nur „Solution löschen“ läuft über einen Flow mit Dienstkonto (siehe „Aktionen“).

| Bereich | Beschreibung |
|---|---|
| Übersicht (Startseite) | Dashboard mit Gesamtzahl der Ressourcen, Countdown zur Sperre nicht verwalteter Developer-Environments (15.12.2026), Kennzahl-Kacheln mit 12-Monats-Verlauf, Verlauf „Neu erstellt pro Monat“ (Reihen per Klick umschaltbar: Apps, Flows, Bots, Desktop Flows, Power Pages, Environments, neue Maker – Auswahl bleibt im Browser gespeichert), **Handlungsbedarf** (betroffene Developer-Environments, gesperrte Flows, verwaiste Apps/Flows, seit 90 Tagen ungenutzte Apps, Gast-Maker), Verteilungen (Apps nach Typ, Flows nach Status, Environments nach Typ) und Ranglisten (meistgenutzte Apps, größte Environments, Top-Connectors, aktivste Maker). Jedes Element führt per Klick in die gefilterte Liste oder öffnet den Datensatz. Nur Zählungen/Gruppierungen – auch bei sehr großen Tabellen schnell. |
| Navigation | Seitenleiste wie in der Admin View: **Inventar** (Apps, Flows, Environments, Connectors, Bots, Desktop Flows, Power Pages Sites, Business Process Flows, Solutions, AI-Modelle), **Governance** (DLP-Richtlinien), **Benutzer** (Maker, Power Platform User, Benutzerrollen, Sicherheitsrollen, Verbindungsidentitäten), **Nutzung** (Environment-Kapazität, Audit-Log), **Auswertungen** (Default Environment, Betroffen ab 15.12., Sicherheit & Freigaben), **Lizenzen & Kapazität** (Copilot-Credits, Abrechnungspläne – live), **ALM-Pipelines** (Pipelines-Hosts; Bereitstellungen, Pipelines, Pipeline-Environments des eingestellten Hosts). Auf schmalen Seiten waagrecht. |
| Suche über alle Bereiche | Suchfeld im Kopf: sucht gleichzeitig in Apps, Flows, Environments, Makern, Connectors, Bots, Solutions, DLP-Richtlinien, Pipelines … (direkt in Dataverse, nicht nur in geladenen Daten). Treffer öffnen die Detailansicht, „Alle anzeigen“ springt in die Liste mit dem Suchbegriff. |
| Sicherheit & Freigaben | Risiken auf einen Blick: Apps für die gesamte Organisation, Apps/Flows an Gäste freigegeben, Custom Connectors über unverschlüsseltes HTTP, Environments ohne Sicherheitsgruppe bzw. ohne DLP-Richtlinie, deaktivierte Konten, Gast-Maker, Apps mit DLP-Verstößen, Environments mit den meisten System-Administratoren, DLP-Richtlinien. Jede Kachel führt in die gefilterte Liste. |
| DLP-Richtlinien | Alle Richtlinien aus dem CoE-Inventar mit Geltungsbereich; in der Detailansicht die zugeordneten **Environments**, **blockierten** und **Business-Connectors** |
| Maker 360° | Detailansicht eines Makers mit Apps, Flows, Bots, Desktop Flows, **Zugriff auf** (alle Freigaben/Rollen der Person), **Sicherheitsrollen** in allen Environments und eigenen Environments |
| Pipelines-Hosts (live, nur lesen) | ALM-Pipelines → „Pipelines-Hosts“: **alle Pipelines-Hosts** des Tenants – das Cockpit findet sie selbst (fragt jedes Dataverse-Environment aus dem CoE nach Pipeline-Environments; Ergebnis 24 h im Browser, „Hosts neu suchen“). Je Host die **Pipelines** mit Entwicklungs-Environment(s), Phasen → Ziel (in Reihenfolge), Besitzer und letzter Bereitstellung (Zeile anklicken = Details, Name öffnet die App „Deployment Pipeline Configuration“ im Host) und alle **Environments** mit Rolle (Entwicklung/Ziel), Pipelines, Prüfstatus und **Managed-Status aus dem CoE** – nicht verwaltete Ziele werden markiert (Microsoft blockiert Bereitstellungen dorthin seit Oktober 2026 nach Übergangsfrist). Filter „Nur mit Hinweis“, Suche, CSV. Der von Microsoft verwaltete „Platform host“ (persönliche Pipelines) ist nur im Admin Center sichtbar. Details unter [Pipelines-Hosts](#pipelines-hosts). |
| ALM-Pipelines (eingestellter Host) | Bereitstellungen (Solution, Version, Pipeline, Stufe, Status, ausgelöst von), Pipelines mit Stufen und Pipeline-Environments als Tabellen – live aus dem Dataverse des **Pipelines-Host-Environments** (Auswahl in den Eigenschaften) |
| Admin-Rolle (PIM) rechts oben | Beim Laden liest das Cockpit die **aktiven Entra-Rollen** der angemeldeten Person aus dem Anmelde-Token (Claim `wids`, keine zusätzliche Berechtigung): **grün** „Power Platform Admin aktiv“ (bzw. „Globaler Admin aktiv“), **gelb** „PIM-Rolle nicht aktiv“, grau, wenn nicht prüfbar. Ein Klick zeigt Details, den Link „In PIM aktivieren“ und „Erneut prüfen“ (holt ein frisches Token – nach dem Aktivieren dauert es oft 1–2 Minuten). Ohne aktive Rolle sind die **Aktions-Schaltflächen ausgeblendet** (siehe „Aktionen“). |
| Meine Rollen (PIM, live) | Am Rollen-Abzeichen „Rolle aktivieren“ bzw. „Meine Rollen“: nur die Entra-Rollen, für die die angemeldete Person in Privileged Identity Management **berechtigt** ist (also aktivieren kann) – aktiv (bis wann), nicht aktiv, Genehmigung ausstehend; keine anderen Rollen, keine dauerhaften Zuweisungen – mit den Regeln der Rolle (maximale Dauer, Begründung, Ticket, MFA, Genehmigung, Authentifizierungskontext). **Aktivieren** (Dauer bis zum Maximum, Begründung, ggf. Ticket) und **vorzeitig beenden** direkt im Cockpit über Microsoft Graph; danach prüft das Cockpit die Rolle automatisch nach (20 s, 1 Min., 2 Min.). Nur Selbst-Aktionen der angemeldeten Person; Rollen über Gruppen oder mit Authentifizierungskontext → Link ins PIM-Portal. Details unter [Meine Rollen (PIM)](#meine-rollen-pim). |
| Hinweis Adminrolle | Bereiche außerhalb des CoE-Environments (Solutions im Default Environment, ALM-Pipelines) werden beim Start geprüft. Fehlt der Zugriff, erscheint **vorab** ein Hinweis, welche Bereiche nicht sichtbar sind und dass die Rolle „Power Platform-Administrator“ (PIM) aktiviert werden muss; die Einträge sind in der Navigation mit einem Schloss markiert. |
| Tabellenansicht | Volltextsuche, Sortierung, Seitenumbruch, kompakte/normale Zeilenhöhe, Spaltenauswahl, CSV-Export für Excel. **Trefferzahl direkt neben dem Suchfeld**: „571 Einträge“, nach Suche/Filter hervorgehoben „9 von 571 Einträgen“. |
| Betroffen ab Februar 2027 | Ab **Februar 2027** öffnen Apps in **Managed Environments** nur noch für Personen mit Premium-Lizenz (Power Apps Premium, Power Automate Premium oder Dynamics 365 Enterprise mit Power-Apps-Rechten) – auch Standard-Apps ([Microsoft](https://learn.microsoft.com/en-us/power-platform/admin/managed-environment-licensing)). Die Auswertung zeigt **alle Berechtigten** der Apps in Managed Environments – Besitzer, Mitbesitzer, Benutzer und Mitglieder freigegebener Gruppen (auch verschachtelt, über Microsoft Graph `transitiveMembers`) –, ihre **Lizenzen live über Microsoft Graph** und wer die Apps in den letzten 30 Tagen geöffnet hat (CoE-Audit-Log). Reiter Personen / Environments / Apps, Filter „Ohne Premium“, CSV-Export, Countdown. Environments mit Pay-as-you-go für Power Apps werden erkannt, wenn die Power Platform API freigegeben ist. Benötigt die Graph-Berechtigungen `User.Read.All` und `GroupMember.Read.All` (Skript `-Licenses`). **Pipelines:** Kachel „Ziel einer Pipeline“ und Spalte „Pipelines“ bei den Environments – Managed Environments, die **Ziel einer Pipeline** sind, können nicht auf „nicht verwaltet“ zurück (Microsoft verlangt Managed für Pipeline-Ziele); Entwicklungs-Environments und Pipelines-Hosts sind gekennzeichnet (Daten aus „Pipelines-Hosts“). |
| Was ändert Microsoft? (live) | Start → „Was ändert Microsoft?“: Meldungen des **Microsoft 365 Message Centers** zur Power Platform (Power Apps, Power Automate, Copilot Studio, Dataverse, Power Pages, Power BI) live aus Microsoft Graph. Kacheln **Handlung nötig** (mit nächster Frist), **wichtige Änderungen**, **neu** (7 Tage); Filter nach Dienst, Sortierung nach Frist, Suche. Der Meldungstext wird in einem abgeschotteten Rahmen ohne Skripte angezeigt; Links ins Message Center und zur Roadmap. Berechtigung `ServiceMessage.Read.All` (Skript `-MessageCenter`); lesen dürfen u. a. Power Platform Administrator, Globaler Leser und „Leseberechtigter für Nachrichtencenter“. |
| Verwaiste Ressourcen | Auswertungen → „Verwaiste Ressourcen“: Apps und Flows, deren Besitzer laut CoE nicht mehr im Unternehmen ist, nach ehemaligem Besitzer gruppiert; Kacheln für aktive Flows und Einträge in Solutions. **Sammel-Zuweisung**: Einträge auswählen, neue Person wählen (CoE-Vorschlag oder E-Mail-Adresse), Begründung – das Cockpit führt den Besitzerwechsel direkt aus (Apps: neuer Besitzer; Flows: neuer Mitbesitzer), schreibt jede Zuweisung ins Aktionsprotokoll und setzt im CoE „verwaist“ zurück. Nur für Mitglieder der Admin-Gruppe mit aktiver Rolle; max. 100 pro Durchlauf, Anhalten möglich; CSV-Export. |
| Lizenz-Optimierung | Lizenzen & Kapazität → „Lizenz-Optimierung“: Wer hat eine Power-Platform-Premium-Lizenz (Power Apps Premium, Power Automate Premium, Plan 2 …; live aus Microsoft Graph), nutzt sie aber nicht? Power Apps gilt als genutzt bei einem Start einer **Premium-App** im Zeitraum (CoE-Audit-Log, Lizenzklasse laut CoE), Power Automate bei einem **eingeschalteten eigenen Flow** (CoE). Zeitraum 30/60/90/180 Tage, Status ungenutzt/teilweise/genutzt, **mögliche Einsparung** pro Monat und Jahr mit anpassbaren Listenpreisen (im Browser gespeichert), CSV-Export. Näherung – vor dem Entziehen nachfragen. |
| Dataverse-Speicher | Lizenzen & Kapazität → „Dataverse-Speicher“: **Kontingent des Tenants** für Datenbank, Datei und Log live aus der Power Platform API (`licensing/tenantCapacity`) mit Ampel (ab 80 % gelb, ab 100 % rot) und Statusmeldung von Microsoft; **Verbrauch je Environment** aus dem CoE (Environment-Kapazität) mit Anteil am Kontingent, Filter nach Typ, Suche, CSV, Klick öffnet das Environment. Ohne Power Platform API bleibt der Verbrauch laut CoE sichtbar. |
| DLP-Auswirkungsanalyse | Governance → „DLP-Auswirkungsanalyse“ – wie die CoE-App **„Data Policy Impact Analysis“**: **angewendete Richtlinien live** (Name, Gilt für, Standardgruppe, Business/Non-Business/Blockiert, erstellt/geändert von) mit Filter „Richtlinien für Environment“, **Entwürfe** und **Aufgabenliste** aus dem CoE. Befehle **Bearbeiten/Analysieren, Kopieren, Im Admin Center öffnen, Löschen, Neue Richtlinie**. Assistent: Name → Connectors (Gruppen wechseln, Standardgruppe, nicht blockierbare gesperrt) → Geltungsbereich → Environments → **Auswirkung** (betroffene Apps und Flows laut CoE, „neu betroffen“ vs. „bereits betroffen“, CSV, „Zur Aufgabenliste“) → Überprüfen (alle Änderungen) → **Als Entwurf speichern** (CoE) oder **Änderungen übernehmen / Als neue Richtlinie anlegen** (direkt im Tenant, mit Protokoll). Bearbeiten, Kopieren, Neue Richtlinie und Löschen nur für Gruppenmitglieder mit **aktiver** Admin-Rolle – ohne aktive Rolle nur die Liste (Stand aus dem CoE). Details unter [DLP-Auswirkungsanalyse](#dlp-auswirkungsanalyse). |
| Lizenzen je Benutzer (live) | Navigation → Lizenzen & Kapazität: alle Benutzer mit ihren zugewiesenen Lizenzen **live aus Microsoft Graph**. Filter nach einer oder mehreren Lizenzen („mind. eine“ / „alle“, auch „ohne Lizenz“), Suche nach Name, Benutzername, Abteilung; Kacheln als Schnellfilter (mit/ohne Lizenz, **deaktiviert mit Lizenz**, **Gäste mit Lizenz**); Lizenzübersicht mit zugewiesen, gekauft und frei (Klick = Filter); CSV-Export. Berechtigung `User.Read.All`, für gekauft/frei zusätzlich `Organization.Read.All` (Skript `-Licenses`) – ohne sie bleiben Namen und Zuweisungen sichtbar. |
| Copilot-Credits (live) | Copilot-Studio-Kapazität je Environment wie im Admin Center (Lizenzierung → Copilot Studio): gekauft, zugewiesen, noch zuweisbar, Prepaid- und Pay-as-you-go-Verbrauch, Status (im Rahmen / überschritten / überschritten – gedeckt), „Aus der Tenant-Kapazität ziehen“, verknüpfter Abrechnungsplan, Typ, Managed. Filter (Mit Zuweisung, Über Kapazität, Pay-as-you-go, Ohne Zuweisung), Suche, Sortierung; „Alle Environments prüfen“ fragt jedes Environment einzeln ab. **Ohne CoE/Dataverse** – direkt aus der Power Platform API. |
| Abrechnungspläne (live) | Alle Pay-as-you-go-Abrechnungspläne mit Status, Produkten, Azure-Abonnement, Ressourcengruppe (Link ins Azure-Portal), Region, Erstelldatum; aufklappbar mit den verknüpften Environments und deren Copilot-Verbrauch. |
| Inhalt der Solution | In der Detailansicht einer Solution: Reiter **Apps, Flows, Bots, Desktop Flows, Power Pages** mit Anzahl. Apps und Flows aus dem CoE-Inventar (Beziehungen „admin_PPSolution_Apps“ / „admin_PPSolution_Flow“, anklickbar), Bots, Desktop Flows und Power Pages live aus dem Environment der Solution (Tabelle `solutioncomponents`), außerdem Link „Solution in Power Apps öffnen“. |
| Spaltenfilter | Filter-Symbol in **jedem** Spaltenkopf: Werteliste mit Suche und Anzahl (wie in Excel), bei Datumsspalten Zeiträume. Aktive Filter als Chips, einzeln entfernbar. |
| Detailansicht | Seitenpanel mit Eigenschaften, Direktlinks (Power Apps, Power Automate, Admin Center, Dataverse), „Alle Felder“ und Reitern wie in den Admin-View-Formularen: **Freigegeben für**, **Verbindungen**, **Nutzung**, **Aktionen**, Ressourcen eines Environments (Apps, Flows, Bots, Desktop Flows, Power Pages, Solutions, Kapazität, Sicherheitsrollen), **Verwendet in Apps/Flows** … Zeilen sind anklickbar (Durchklicken), „Zurück“ führt zum vorherigen Datensatz. |
| Betroffen ab 15.12. | Alle **nicht verwalteten Developer-Environments** (Sperre 15.12.2026, Löschung ab 15.01.2027) mit Ersteller, Zustand, Anzahl Apps/Flows. Managed-Status aus „Governance Protection Level“ („Basic“ = nicht managed), per Auswahl in den Eigenschaften änderbar. |
| E-Mail an Maker | Im Reiter „Betroffen ab 15.12.“ öffnet das **Brief-Symbol** am Zeilenanfang (oder die Schaltfläche in der Detailansicht) eine **Vorschau der formatierten E-Mail** an alle Maker des Environments: Kopfbalken in der Theme-Farbe, Fristen-Kacheln (Rückmeldung, Sperre 15.12.2026, Löschung 15.01.2027), Steckbrief mit **Link zum Environment in make.powerapps.com**, Schaltfläche „Environment in Power Apps öffnen“, nummerierte **Rückmelde-Optionen** (1 = weiter benötigt → Managed + Premium-Lizenz, 2 = löschen, 3 = umziehen; optional Formular), Tipp zum Sichern (Solutions exportieren) und die **offizielle Microsoft-Ankündigung** (Abschnitt „Deprecation of unmanaged developer environments …“ auf Microsoft Learn). „In Outlook öffnen“ lädt die fertige Mail als Outlook-Entwurf (.eml) herunter – Datei öffnen, Outlook zeigt sie mit Empfängern, Betreff und formatiertem Text. Alternativ „Über Zwischenablage“ (Empfänger/Betreff per Link, Text mit Strg+V). Die Mail ist „Outlook-fest“ aufgebaut (Tabellen, Farben direkt am Text, keine runden Elemente), damit sie auch im klassischen Outlook sauber aussieht. Mit **„Direkt senden“** geht die Mail ohne Outlook über Microsoft Graph aus dem eigenen Postfach. **„Alle gefilterten verständigen“** schickt jedem gefilterten Environment eine **eigene E-Mail** – automatisch (mit Sicherheitsabfrage, Fortschritt, Anhalten, Fehlerliste, ca. 28 Mails pro Minute wegen des Exchange-Limits) oder **nacheinander in Outlook** (Schritt für Schritt). **Gemeinsames Versandprotokoll:** Jede Benachrichtigung (automatisch gesendet oder in Outlook gesendet und „Als versendet markiert“) wird in der SharePoint-Liste **„CoE Cockpit Benachrichtigungen“** auf der Website des Cockpits vermerkt – mit Environment, Zeitpunkt, Person, Art, Empfängern und Betreff. Alle Admins sehen in der Tabelle die Spalten „Maker verständigt / Verständigt am / Verständigt von“ und die Kachel „Noch nicht verständigt“; wer ein bereits verständigtes Environment erneut anschreiben will, bekommt eine deutliche Warnung. Beim Serienversand wird das Protokoll unmittelbar vor dem Senden neu gelesen – was ein anderer Admin inzwischen verständigt hat, wird übersprungen. „E-Mail-Adressen kopieren“ liefert die Adressen aller gefilterten Environments. Anwendungsbenutzer („# …“) und SYSTEM werden nicht angeschrieben. |
| Default Environment | Apps und Flows aus dem CoE-Inventar, Solutions live aus dem Dataverse des Default Environments |
| Nachfrage Verwendung | In der Detailansicht von **Apps, Flows, Environments, Connectors, Bots, Desktop Flows, Power Pages, Business Process Flows, Solutions und AI-Modellen**: E-Mail an den Besitzer laut CoE (Maker-Tabelle), ob die Komponente noch benötigt wird – mit Steckbrief (Typ, Name, Environment, Besitzer, letzte Verwendung), nummerierten Antwortmöglichkeiten, Frist und Link zum Öffnen. Versand wie bei „Betroffen ab 15.12.“: als Outlook-Entwurf, über die Zwischenablage oder direkt (Microsoft Graph). Jede Nachfrage wird im gemeinsamen Versandprotokoll vermerkt; wer erneut nachfragen will, sieht „Bereits nachgefragt – am … von …“. |
| E-Mail-Vorlagen | Webpart-Einstellungen → **„E-Mail & Vorlagen“**: Texte der Vorlagen „Nachfrage Verwendung“ und „Betroffen ab 15.12.“ (Betreff, Überschrift, Anrede, Einleitung, Bitte um Rückmeldung, Antwortmöglichkeiten, Hinweis, Grußformel) mit Platzhaltern wie `{Name}`, `{Typ}`, `{Environment}`, `{Besitzer}`, `{Frist}` und Live-Vorschau. Für „Nachfrage Verwendung“ optional eigene Fassungen je Komponententyp (z. B. nur für Flows) – nicht geänderte Felder kommen aus der gemeinsamen Vorlage. Gespeichert werden nur Abweichungen vom Standard; „Auf Standard zurücksetzen“ stellt den Originaltext wieder her. |
| Webpart-Einstellungen | Schaltfläche **„Webpart-Einstellungen“** oben neben „Tabellen verwalten“ (Bearbeitungsmodus): alle Einstellungen nach Bereichen (Verbindung, Ansichten, Default Environment, Betroffen ab 15.12., E-Mail an Maker, Aktionen, ALM-Pipelines, Darstellung, Tabellen) mit Status-Punkten, Einrichtung der Aktionen mit Checkliste. Der klassische Eigenschaftenbereich bleibt erreichbar (Abschnitte zugeklappt). |
| Tabellen-Editor | Tabellen **aus einer Liste wählen** (alle Admin-View-Tabellen als Vorlagen oder jede andere Tabelle der Umgebung), Spalten aus Dataverse auswählen, JSON-Import/-Export |
| SharePoint-App-Leiste | Die Leiste am linken Rand (z. B. „Discover, Publish, Build, OneDrive“) wird auf der Seite mit dem Cockpit ausgeblendet (Webpart-Einstellungen → Darstellung, Standard: an). Microsoft bietet dafür keinen offiziellen Schalter mehr; das Cockpit setzt einen Stil auf `#sp-appBar` – ändert Microsoft die Seite, bleibt die Leiste einfach sichtbar. |
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
   Für „Betroffen ab Februar 2027“ und „Lizenzen je Benutzer“ **Microsoft Graph – User.Read.All**, **GroupMember.Read.All** und
   **Organization.Read.All** genehmigen (optional; Lizenzprüfung, gekaufte/freie Lizenzen). Für „Was ändert Microsoft?“
   **Microsoft Graph – ServiceMessage.Read.All** genehmigen (optional; Skript `-MessageCenter`). Für **„Meine Rollen (PIM)“**
   **Microsoft Graph – RoleEligibilitySchedule.Read.Directory, RoleAssignmentSchedule.ReadWrite.Directory, RoleManagementPolicy.Read.Directory**
   (optional; Skript `-Pim`, Hinweise unter [Meine Rollen (PIM)](#meine-rollen-pim)).
   Für den automatischen E-Mail-Versand zusätzlich **Microsoft Graph – Mail.Send** genehmigen (optional; ohne diese
   Genehmigung funktioniert „In Outlook öffnen“ bzw. „Nacheinander in Outlook“ weiterhin). Für **Copilot-Credits und
   Abrechnungspläne** die drei Lese-Anforderungen der **Power Platform API** genehmigen (optional). Für die **Aktionen** (Quarantäne,
   Löschen, Besitzer, Freigaben, Flows ein/aus, Environments) **PowerApps Service – User**, **Microsoft Flow Service – User,
   Flows.Read.All, Flows.Manage.All** und **Power Platform API – EnvironmentManagement.Write, EnvironmentManagement.Settings.ReadWrite**
   genehmigen (optional; Skript `-Actions`). Für die **DLP-Auswirkungsanalyse** reicht **PowerApps Service – User** (Skript `-Dlp`,
   ist auch in `-Actions` enthalten). Achtung: API-Genehmigungen für
   SharePoint-Webparts gelten für alle SPFx-Lösungen im Tenant. **Erscheinen unter „API-Zugriff“ keine Anforderungen**, setzt das
   Skript [`scripts/Grant-CoeCockpitPermissions.ps1`](scripts/Grant-CoeCockpitPermissions.ps1) dieselben Berechtigungen
   (Microsoft Graph PowerShell, als Globaler Administrator bzw. Administrator für privilegierte Rollen).
3. Webpart **CoE Cockpit** auf einer Seite einfügen und in den Eigenschaften die **URL der CoE-Umgebung** eintragen,
   z. B. `https://contoso-coe.crm4.dynamics.com`.

## Berechtigungen

Das Webpart liest mit den Rechten des angemeldeten Benutzers. Benutzer brauchen in der CoE-Umgebung eine Sicherheitsrolle
mit Lesezugriff auf die CoE-Tabellen (z. B. **Power Platform Admin SR** aus dem CoE Starter Kit oder System Administrator).
Benutzer ohne Rechte sehen eine verständliche Fehlermeldung statt Daten.

## Aktionen für Apps, Flows und Environments

Im Bereich **„Aktionen“** der Detailansicht (unten):

| Ressource | Aktionen | Ausführung |
|---|---|---|
| App | In Quarantäne stellen / Aus Quarantäne holen, Besitzer ändern, Freigabe hinzufügen (Benutzer/Mitbesitzer), Freigabe entfernen, Löschen | direkt im Webpart |
| Flow | Ausschalten / Einschalten, Besitzer ändern, Mitbesitzer hinzufügen, Freigabe entfernen, Löschen | direkt im Webpart |
| Environment (nur „Betroffen ab 15.12.“) | Auf Managed umstellen, Environment deaktivieren / aktivieren (dazu „Maker per E-Mail verständigen“) | direkt im Webpart |
| DLP-Richtlinie (nur „DLP-Auswirkungsanalyse“) | Änderungen übernehmen, Als neue Richtlinie anlegen (z. B. nach „Kopieren“), Löschen | direkt im Webpart |
| Solution | Solution löschen – nur **nicht verwaltete** Solutions (nur der Container; Apps, Flows und andere Komponenten bleiben im Environment). Verwaltete und System-/Microsoft-Solutions sind gesperrt. | Flow (Dienstkonto) |

**Grundsatz:** Was die Rolle **Power Platform Administrator** kann, führt das Cockpit **direkt** aus – im Namen der angemeldeten Person,
ohne Flow und ohne Dienstkonto. Microsoft prüft dabei bei jedem Aufruf selbst, ob die Rolle **aktiv** ist (PIM wird also erzwungen), und im
Audit-Log steht die echte Person. Nur was zusätzlich Rechte **im jeweiligen Environment** braucht (System Administrator bzw. System Customizer),
läuft als Auftrag über den Flow – derzeit nur „Solution löschen“.

**Protokoll:** Jede Aktion steht in der SharePoint-Liste **„CoE Cockpit Aktionen“** (Aktion, Ressource, IDs, Zielperson, Pflicht-Begründung,
Person, Ergebnis, Zeitpunkt) und im **Aktionsprotokoll** (Governance):
- Direkte Aktionen: Das Cockpit legt den Eintrag **vor** dem Aufruf mit Status „In Arbeit“ an (scheitert das, wird die Aktion nicht ausgeführt)
  und schreibt danach „Erledigt“, „Fehler“ oder „Abgelehnt“ mit Ergebnistext zurück.
- Flow-Aufträge: Eintrag mit Status „Angefordert“; der Flow übernimmt nur diese (Auslösebedingung) und schreibt das Ergebnis zurück.

**Sichtbarkeit:** Die Aktions-Schaltflächen (Quarantäne, Löschen, Besitzer, Freigaben, Flows ein/aus, Environments, Solution löschen) sehen nur
Personen, die **Mitglied der SharePoint-Gruppe** (Standard „CoE Cockpit Admins“) sind **und** die Rolle **„Power Platform Administrator“**
(oder „Globaler Administrator“) **gerade aktiv** haben – geprüft beim Laden über den Claim `wids` im Anmelde-Token, dasselbe Ergebnis wie das
Rollen-Abzeichen rechts oben. Ist die Rolle nicht aktiv, sind die Schaltflächen ausgeblendet; Gruppenmitglieder sehen stattdessen einen Hinweis
(„In PIM aktivieren, danach ‚Erneut prüfen‘“). Die E-Mail-Funktionen („Nachfrage Verwendung“, „Maker per E-Mail verständigen“) und das
Aktionsprotokoll brauchen nur die Gruppenmitgliedschaft. Für alle anderen ist der Abschnitt „Aktionen“ ausgeblendet.

**Regeln:**
- Löschen nie, wenn die App bzw. der Flow **Teil einer Solution** ist (CoE-Felder „Is Solution App“ / „Is Solution Flow“, vor dem Löschen frisch
  gelesen; unbekannt = gesperrt). Bots werden nicht gelöscht.
- Optional (Standard: an): Löschen erst, wenn die App **in Quarantäne** bzw. der Flow **ausgeschaltet** ist.
- Pro Ressource immer nur ein offener Auftrag; eine erledigte Environment-Aktion wird bis zum nächsten CoE-Sync (max. 60 Minuten) nicht noch einmal angeboten.

**Direkte Aufrufe**

| Aktion | Schnittstelle (im Namen der angemeldeten Person) | CoE-Inventar danach |
|---|---|---|
| In Quarantäne / Aus Quarantäne | `POST api.powerapps.com/…/scopes/admin/environments/{env}/apps/{app}/setPowerAppQuarantineState` | „Quarantine App“ + Datum, „App Is Quarantined“ |
| Löschen (App) | `DELETE …/apps/{app}` | „App Deleted“ + Datum |
| Besitzer ändern (App) | `POST …/apps/{app}/modifyAppOwner` (bisheriger Besitzer wird Mitbesitzer) | „App Owner“ |
| Freigabe hinzufügen / entfernen (App) | `POST …/apps/{app}/modifyPermissions` | – (nächster CoE-Sync) |
| Flow aus / ein | `POST api.flow.microsoft.com/…/scopes/admin/environments/{env}/flows/{flow}/stop` bzw. `/start` | „Flow State“ |
| Löschen (Flow) | `DELETE …/flows/{flow}` | „Flow Deleted“ + Datum |
| Besitzer / Mitbesitzer / Freigabe entfernen (Flow) | `POST …/flows/{flow}/modifyOwners` (als Mitbesitzer – den Ersteller kann Power Automate nicht ändern) | „Derived Owner“ (Besitzer ändern) |
| Auf Managed umstellen | Power Platform API `POST /environmentmanagement/environments/{id}/governanceSetting/enableManaged` | – (nächster CoE-Sync) |
| Environment deaktivieren / aktivieren | Power Platform API `POST /environmentmanagement/environments/{id}/Disable` (Grund „UserRequested“) bzw. `/Enable` | – (nächster CoE-Sync) |
| DLP-Richtlinie ändern / anlegen / löschen | `PATCH` / `POST` / `DELETE api.bap.microsoft.com/providers/PowerPlatform.Governance/v2/policies[/{id}]` (immer die vollständige Richtlinie) | – (nächster CoE-Sync) |

Die Power-Apps- und Power-Automate-Aufrufe sind dieselben Admin-Schnittstellen, die die Connectors „Power Apps for Admins“ / „Power Automate
for Admins“ und Microsofts PowerShell-Modul `Microsoft.PowerApps.Administration.PowerShell` verwenden; sie erlauben Aufrufe aus dem Browser
(CORS). Die Environment-Aufrufe sind offiziell dokumentierte Endpunkte der Power Platform API (derzeit „Preview“). Zielpersonen, die nicht im CoE
stehen, werden über ihre E-Mail-Adresse in Microsoft Graph gesucht (`User.Read.All`). Das CoE-Inventar zieht das Cockpit mit den Rechten der
Person nach; klappt das nicht, steht ein Hinweis im Ergebnis – der nächste CoE-Sync korrigiert es ohnehin.

**Einrichtung (Webpart-Einstellungen → „Aktionen“):**
1. **SharePoint-Gruppe** wählen oder mit „Gruppe anlegen“ erstellen (Standard: **„CoE Cockpit Admins“**).
2. **„Aktionsliste einrichten“**: legt die Liste an und gibt ihr eigene Berechtigungen – nur die Gruppe (Bearbeiten) und die Websitebesitzer (Vollzugriff).
3. Berechtigungen für die direkten Aufrufe freigeben: `.\scripts\Grant-CoeCockpitPermissions.ps1 -Actions` (oder unter „API-Zugriff“ genehmigen).
4. Für „Solution löschen“: Flow „CoE Cockpit | Aktion ausführen“ im CoE-Environment importieren (siehe nächster Abschnitt); dort denselben Gruppennamen eintragen.

### Flow „CoE Cockpit | Aktion ausführen“ (nur „Solution löschen“)

Fertige, nicht verwaltete Lösung: [`flow/CoECockpitAktionen_1_2_0_0.zip`](flow/CoECockpitAktionen_1_2_0_0.zip)
(erzeugt mit [`flow/build-flow-solution.py`](flow/build-flow-solution.py); lesbare Definition: `flow/CoECockpit-Aktionausfuehren.definition.json`).

Version 1.2.0.0 ist bereinigt: Sie enthält nur noch, was ein Dienstkonto mit Rechten im Environment braucht – **2 Verbindungsreferenzen**
(SharePoint, Dataverse) statt 7. Power Apps for Admins, Power Automate for Admins, Power Automate Management, Power Platform for Admins V2 und
Office 365 Users werden nicht mehr gebraucht.

**Ablauf:** Neuer Eintrag mit Status „Angefordert“ in „CoE Cockpit Aktionen“ (Auslösebedingung – direkt ausgeführte Aktionen stehen schon auf
„In Arbeit“ und werden ignoriert) → Status „In Arbeit“ → prüfen, ob „Erstellt von“ Mitglied der SharePoint-Gruppe ist (sonst „Abgelehnt“) →
Solution im Environment lesen (nur nicht verwaltet, keine System-Solution) → löschen (`DeleteRecordWithOrganization`) → CoE-Inventar
„Solution Deleted“ + Datum → Status „Erledigt“, „Fehler“ oder „Abgelehnt“ zurückschreiben. Andere Aktionen lehnt der Flow mit dem Hinweis ab,
das Webpart zu aktualisieren.

**Import (im CoE-Environment):**
1. make.powerapps.com → CoE-Environment → *Lösungen* → *Importieren* → ZIP auswählen.
2. **Verbindungen** mit einem **Dienstkonto** anlegen bzw. auswählen: **SharePoint** und **Microsoft Dataverse**. Das Dienstkonto braucht
   **System Administrator** (oder System Customizer) in den Environments, in denen Solutions gelöscht werden sollen, und Schreibrechte im CoE.
3. **Umgebungsvariablen:** SharePoint-Website des Cockpits (z. B. `https://contoso.sharepoint.com/sites/powerplatform`),
   Name der Aktionsliste (Standard „CoE Cockpit Aktionen“ – alternativ die Listen-ID), SharePoint-Gruppe (Standard „CoE Cockpit Admins“).
4. Dem Dienstkonto auf der Liste **„CoE Cockpit Aktionen“** die Berechtigung **Bearbeiten** geben (die Liste hat eigene Berechtigungen).
5. Flow einschalten (falls er nach dem Import aus ist) und mit einer Test-Solution prüfen.

**Update von Version 1.1.0.0:** Die neue ZIP einfach erneut importieren. Bei nicht verwalteten Lösungen entfernt der Import keine Bestandteile –
die nicht mehr benötigten Verbindungsreferenzen (Power Apps for Admins, Power Automate for Admins, Power Automate Management, Power Platform for
Admins V2, Office 365 Users) danach in der Lösung bzw. unter *Verbindungsreferenzen* löschen.

**Hinweise:**
- Mitgliedschaft wird direkt in der SharePoint-Gruppe geprüft; Personen, die nur über eine darin enthaltene Entra-Gruppe Mitglied sind, werden abgelehnt.
- Die Flow-Definition konnte hier nicht gegen einen echten Tenant getestet werden – bitte zuerst mit einer Test-Solution prüfen.

## DLP-Auswirkungsanalyse

Nachbau der CoE-App **„Data Policy Impact Analysis“** (DLP Editor V2) im Cockpit – mit denselben Datenquellen:

| Was | Woher |
|---|---|
| Angewendete Richtlinien | **live** über die Admin-Schnittstelle `api.bap.microsoft.com/providers/PowerPlatform.Governance/v2/policies` – dieselbe wie der Connector „Power Platform for Admins V2“ (den auch die CoE-App nutzt) und die PowerShell-Befehle `Get-/New-/Set-/Remove-DlpPolicy`. Token für **PowerApps Service – User**. |
| Ohne aktive Rolle | Richtlinien aus der CoE-Tabelle **„DLP Policy“** (Stand letzter Sync; Environments und blockierte/Business-Connectors aus den N:N-Beziehungen, Standardgruppe dort nicht erfasst) – nur die Liste; Bearbeiten, Kopieren, Neue Richtlinie und Löschen sind ausgeblendet. |
| Connectors | CoE-Tabelle **„Connector“** (ohne Custom Connectors – die regelt Microsoft über URL-Muster im Admin Center) |
| Auswirkung | CoE-Felder **„App Connections“ / „Flow Connections“** der Apps und Flows in den Environments, für die die Richtlinie gilt (bis 50.000 je Typ) |
| Entwürfe | CoE-Tabelle **„Data Policy Drafts“** – dieselben Entwürfe wie in der CoE-App |
| Aufgabenliste | CoE-Tabelle **„DLP Impact Analysis“** – Einträge wie der CoE-Flow „DLP Editor \| Parse impacted resources“; Entscheidung, Status und Benachrichtigung der Maker weiter in der CoE-App |

**Auswirkung:** Betroffen ist eine App bzw. ein Flow, der einen **blockierten** Connector nutzt oder **Business- und Non-Business-Connectors mischt**
(Connectors ohne Zuordnung zählen zur Standardgruppe). Beim Bearbeiten einer bestehenden Richtlinie zeigt das Cockpit zusätzlich, was **neu** betroffen wäre
und was schon mit der bisherigen Richtlinie gegen sie verstößt. Grundlage ist der letzte CoE-Sync.

**Sichtbarkeit:** Bearbeiten/Analysieren, Kopieren, Neue Richtlinie, Löschen und alle Speichern-Schaltflächen sehen nur Mitglieder der SharePoint-Gruppe
mit **aktiver** Rolle „Power Platform Administrator“ (wie bei allen Aktionen). Alle anderen sehen nur die Liste der Richtlinien, die Entwürfe und die
Aufgabenliste sowie „Im Admin Center öffnen“.

**Speichern im Tenant** (mit Pflicht-Begründung und Eintrag im Aktionsprotokoll):
- **Änderungen übernehmen**: liest die Richtlinie unmittelbar vorher neu; hat sie inzwischen jemand anderes geändert, wird **nicht** gespeichert
  (Status „Abgelehnt“). Zurückgeschrieben wird die **vollständige** Richtlinie – Aktions- und Endpunkt-Regeln bleiben unverändert.
- **Als neue Richtlinie anlegen**: nach „Kopieren“ oder „Neue Richtlinie“ bzw. aus einem Entwurf.
- **Löschen**: nur nach Eingabe des Richtliniennamens. Entwürfe löschen ändert nichts im Tenant.
- Connectors, die Microsoft nicht blockieren lässt (z. B. SharePoint, Office 365 Outlook, Teams, Dataverse), lassen sich nicht nach „Blockiert“ verschieben.

Die DLP-Schnittstelle ist – wie die Power-Apps-/Power-Automate-Admin-Aufrufe der übrigen Aktionen – nicht als eigene REST-API dokumentiert,
sondern nur über den Connector und PowerShell. Geprüft: Sie erlaubt Aufrufe aus dem Browser (CORS, alle Methoden) und verlangt ein Token für
`https://service.powerapps.com/` – genau das, was „PowerApps Service – User“ freigibt. Ohne Admin-Rolle liefert sie „Zugriff verweigert“; dann zeigt das Cockpit den CoE-Stand.
Die Entwürfe und die Aufgabenliste brauchen Schreibrechte auf diese CoE-Tabellen (z. B. „Power Platform Admin SR“).

## Pipelines-Hosts

Übersicht über **alle** Power Platform Pipelines – nur lesen, mit den Rechten der angemeldeten Person (Rolle „Power Platform Administrator“
bzw. Lesezugriff auf die Pipelines-Tabellen im Host).

- **Hosts finden:** Das Cockpit fragt jedes Environment mit Dataverse aus dem CoE nach der Tabelle „Deployment Environment“
  (`deploymentenvironments`, eine Zeile). Wo Einträge stehen, ist ein Pipelines-Host. Das Ergebnis bleibt 24 Stunden im Browser
  (nur vollständige Suchen ohne „Zugriff verweigert“); „Hosts neu suchen“ startet sie neu. Der Host aus den Webpart-Einstellungen ist immer dabei.
- **Daten je Host** (Microsoft Learn „Pipeline table reference“): Pipelines (`deploymentpipeline`), Phasen mit Ziel und Vorgänger
  (`deploymentstage`), Environments mit Rolle, Environment-ID, Prüfstatus und Fehlermeldung (`deploymentenvironment`), die letzten
  1.000 Bereitstellungen (`deploymentstagerun`). Die Entwicklungs-Environments je Pipeline kommen aus der N:N-Beziehung Pipeline ↔ Environment
  (Name aus den Metadaten, per `$expand`); ist sie nicht lesbar, leitet das Cockpit sie aus den Bereitstellungen ab.
- **Managed-Status** der Environments aus dem CoE („Governance Protection Level“). Ziel-Environments müssen Managed sein.
- **Platform host:** Persönliche Pipelines der Maker liegen in einem von Microsoft verwalteten Host, den nur das Admin Center
  (Deployment → Platform host) zeigt; dafür gibt es keine dokumentierte Schnittstelle.
- **Betroffen ab Februar 2027** nutzt dieselben Daten: Pipeline-Ziele bleiben Managed.

## Meine Rollen (PIM)

Jede Person sieht im Cockpit **nur die Entra-Rollen, für die sie in Privileged Identity Management berechtigt ist** (keine anderen Rollen des Tenants, keine dauerhaften Zuweisungen), und kann sie aktivieren und vorzeitig beenden –
live über Microsoft Graph (`/roleManagement/directory`):

| Was | Schnittstelle |
|---|---|
| Berechtigte Rollen | `GET roleEligibilityScheduleInstances/filterByCurrentUser(on='principal')` |
| Aktive Rollen | `GET roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')` |
| Ausstehende Anfragen | `GET roleAssignmentScheduleRequests/filterByCurrentUser(on='principal')` (Status „PendingApproval“) |
| Regeln der Rolle | `GET /policies/roleManagementPolicyAssignments` (Dauer, Begründung, Ticket, MFA, Genehmigung, Authentifizierungskontext) |
| Aktivieren / Beenden | `POST roleAssignmentScheduleRequests` mit `selfActivate` bzw. `selfDeactivate` – immer für die angemeldete Person |

- **Nur Selbst-Aktionen:** Das Cockpit schickt ausschließlich `selfActivate`/`selfDeactivate` mit der eigenen Objekt-ID. Rollen für andere
  zuweisen oder Berechtigungen ändern kann man darüber nicht. Microsoft prüft jede Anfrage gegen die Regeln der Rolle und protokolliert sie
  im **Entra-Audit-Log** (keine zusätzliche SharePoint-Liste).
- **MFA:** Verlangt die Rolle MFA, klappt die Aktivierung, wenn die SharePoint-Sitzung mit MFA angemeldet wurde; sonst meldet das Cockpit das
  und verweist aufs PIM-Portal. Rollen mit **Authentifizierungskontext** (Conditional Access) und Berechtigungen **über Gruppen** (PIM für
  Gruppen) gehen nur im PIM-Portal.
- Nach dem Aktivieren dauert es oft 1–2 Minuten, bis die Rolle im Anmelde-Token steht; das Abzeichen prüft automatisch nach.
- Eine gerade aktivierte Rolle lässt sich laut Microsoft erst nach 5 Minuten wieder beenden.
- **Berechtigungen** (Skript `-Pim`): Microsoft Graph `RoleEligibilitySchedule.Read.Directory`, `RoleAssignmentSchedule.ReadWrite.Directory`,
  `RoleManagementPolicy.Read.Directory`. **Achtung:** Wie alle SPFx-Berechtigungen gilt `RoleAssignmentSchedule.ReadWrite.Directory` für
  **alle** Webparts im Tenant. Eine normale Person kann damit nur eigene berechtigte Rollen aktivieren; ein Webpart, das ein Administrator für
  privilegierte Rollen öffnet, könnte damit aber in dessen Namen Rollen zuweisen. Deshalb nur vertrauenswürdige Lösungen im App-Katalog zulassen.
  Ausschalten: Webpart-Einstellungen → Ansichten → „Meine Rollen (PIM)“.

## Copilot-Credits und Abrechnungspläne (live)

Diese beiden Ansichten (Navigation → **Lizenzen & Kapazität**) brauchen weder CoE noch Dataverse. Sie lesen mit den Rechten des
angemeldeten Benutzers direkt aus der **Power Platform API** (`https://api.powerplatform.com`, Bereich Licensing):

| Abfrage | Zweck |
|---|---|
| `GET /licensing/entitlements/MCSMessages` | Tenant-Summen: gekauft, zugewiesen, verbraucht, Pay-as-you-go |
| `GET /licensing/tenantCapacity/currencyReports` | Rückfall für gekauft/zugewiesen |
| `GET /licensing/allocationsByEnvironment` | Zuweisungen je Environment inkl. „Aus der Tenant-Kapazität ziehen“ |
| `GET /licensing/entitlements/MCSMessages/resources` | Verbrauch im laufenden Monat (findet Environments ohne Zuweisung) |
| `GET /licensing/environments/{id}/entitlements` | Details je Environment: Zuweisung, Verbrauch, Status, Pay-as-you-go, Typ, Managed |
| `GET /licensing/billingPolicies` + `/{id}/environments` | Abrechnungspläne und verknüpfte Environments |
| `GET /environmentmanagement/environments` | Namen aller Environments („Alle Environments prüfen“) |

**Voraussetzungen**
1. Freigabe der Power Platform API für SharePoint-Webparts: Anforderungen unter „API-Zugriff“ genehmigen oder
   `.\scripts\Grant-CoeCockpitPermissions.ps1 -PowerPlatformApi` ausführen (legt bei Bedarf auch den Dienstprinzipal „Power Platform API“ an).
2. Wer die Ansicht öffnet, braucht die **aktive** Rolle **Power Platform Administrator** (oder Globaler Administrator) – bei PIM vorher aktivieren,
   danach die Seite neu laden (Tokens werden bis zu einer Stunde zwischengespeichert; notfalls ab- und wieder anmelden).

Fehlt eines davon, zeigt das Cockpit statt der Daten einen klaren Hinweis (mit Link zur PIM-Aktivierung). Ein- und ausblenden lassen sich
die Ansichten in den Webpart-Eigenschaften unter „Lizenzen & Kapazität (live)“. Zuweisungen ändern geht weiterhin im Admin Center
(Lizenzierung → Copilot Studio → Kapazität verwalten).

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
| Solution → Bots, Desktop Flows, Power Pages | Environment der Solution | Leserechte auf `solutioncomponents`, `workflows`, `bots`, `powerpagesites` (z. B. System Administrator) |
| ALM-Pipelines, Pipelines-Hosts | Pipelines-Host-Environment(s) | Leserechte auf die Pipelines-Tabellen (z. B. System Administrator, Deployment Pipeline Administrator; als Power Platform Administrator vorhanden) |

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
| Ansichten „Copilot-Credits“ und „Abrechnungspläne“ anzeigen | Live-Ansichten über die Power Platform API ein-/ausblenden |
| E-Mail-Vorlagen | Webpart-Einstellungen → „E-Mail & Vorlagen“ (Texte, Platzhalter, Vorschau; je Komponententyp optional) |
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
│  ├─ CopilotCreditsView.tsx     Copilot-Credits und Abrechnungspläne (live, Power Platform API)
│  ├─ SolutionContents.tsx       Inhalt einer Solution (Detailansicht)
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
├─ services/PowerPlatformApiService.ts  Power Platform API (Licensing) mit Versions-Fallback und Seitenabruf
└─ utils/format.ts               Formatierung, CSV, Hilfsfunktionen
```

## Nächste Ausbaustufe

- Copilot-Credits zuweisen direkt aus dem Cockpit (Power Platform API `licensing/allocationsV2`, Rolle Power Platform Administrator).
