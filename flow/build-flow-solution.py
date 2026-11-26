#!/usr/bin/env python3
"""
Erzeugt die importierbare Power-Platform-Lösung „CoE Cockpit Aktionen“ (nicht verwaltet) mit dem Flow
„CoE Cockpit | Aktion ausführen“. Aufruf:  python3 flow/build-flow-solution.py   → flow/CoECockpitAktionen_<Version>.zip

Ab Webpart-Version 1.14 führt das CoE Cockpit alle Aktionen, für die die Rolle „Power Platform Administrator“
reicht, selbst aus (im Namen der angemeldeten Person). Der Flow mit Dienstkonto bleibt nur für Aktionen, die
zusätzlich Rechte im jeweiligen Environment brauchen – derzeit „Solution löschen“ (System Administrator bzw.
System Customizer im Ziel-Environment).

Der Flow arbeitet nur Aufträge mit Status „Angefordert“ ab (Auslösebedingung). Einträge, die das Webpart
für direkt ausgeführte Aktionen ins Protokoll schreibt („In Arbeit“ → „Erledigt“/„Fehler“), ignoriert er.
  1. Status „In Arbeit“
  2. Prüfen, ob die anfragende Person (Erstellt von) Mitglied der SharePoint-Gruppe ist – sonst „Abgelehnt“
  3. Aktion ausführen (Dataverse-Verbindung des Dienstkontos)
  4. CoE-Inventar nachziehen (Solution gelöscht) – Fehler dabei sind nur ein Hinweis
  5. Status „Erledigt“, „Fehler“ oder „Abgelehnt“ mit Ergebnistext und Zeitpunkt zurückschreiben
"""
import io
import json
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
VERSION = '1.2.0.0'
SOLUTION = 'CoECockpitAktionen'
FLOW_ID = '7c0c0c4a-3b1e-4f6a-9d2e-5a1b2c3d4e01'
FLOW_NAME = 'CoE Cockpit | Aktion ausführen'
FLOW_FILE = f'CoECockpit-Aktionausfuehren-{FLOW_ID.upper()}.json'

# ---------- Umgebungsvariablen ----------
ENV_VARS = [
    ('coeck_CockpitSiteUrl', 'CoE Cockpit – SharePoint-Website',
     'URL der SharePoint-Website, auf der das CoE Cockpit eingebunden ist und die Liste „CoE Cockpit Aktionen“ liegt, z. B. https://contoso.sharepoint.com/sites/powerplatform', None),
    ('coeck_ActionListName', 'CoE Cockpit – Name der Aktionsliste',
     'Titel oder ID der SharePoint-Liste mit den Aufträgen (Standard: CoE Cockpit Aktionen).', 'CoE Cockpit Aktionen'),
    ('coeck_AdminGroupName', 'CoE Cockpit – SharePoint-Gruppe',
     'Nur Aufträge von Mitgliedern dieser SharePoint-Gruppe werden ausgeführt (dieselbe Gruppe wie in den Webpart-Eigenschaften, Standard: CoE Cockpit Admins).', 'CoE Cockpit Admins'),
]
P_SITE = "@parameters('CoE Cockpit – SharePoint-Website (coeck_CockpitSiteUrl)')"
P_LIST = "parameters('CoE Cockpit – Name der Aktionsliste (coeck_ActionListName)')"
P_GROUP = "parameters('CoE Cockpit – SharePoint-Gruppe (coeck_AdminGroupName)')"

# ---------- Verbindungsreferenzen ----------
CONN = {
    'sp': ('shared_sharepointonline', 'coeck_SharePoint', 'CoE Cockpit – SharePoint', 'Liest Aufträge aus der Liste „CoE Cockpit Aktionen“ und schreibt das Ergebnis zurück.'),
    'dv': ('shared_commondataserviceforapps', 'coeck_Dataverse', 'CoE Cockpit – Dataverse',
           'Löscht Solutions im Ziel-Environment (Dienstkonto mit System Administrator dort) und aktualisiert das CoE-Inventar.'),
}


def api(key):
    return f'/providers/Microsoft.PowerApps/apis/{CONN[key][0]}'


def op(key, operation, params, run_after=None, retry=False):
    a = {
        'type': 'OpenApiConnection',
        'inputs': {
            'host': {'connectionName': CONN[key][0], 'operationId': operation, 'apiId': api(key)},
            'parameters': params,
            'authentication': "@parameters('$authentication')"
        }
    }
    if retry:
        a['inputs']['retryPolicy'] = {'type': 'exponential', 'count': 4, 'interval': 'PT15S'}
    if run_after is not None:
        a['runAfter'] = run_after
    return a


def sp_http(method, uri, body=None, run_after=None, merge=False):
    headers = {'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata'}
    if merge:
        headers.update({'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE'})
    params = {'dataset': P_SITE, 'parameters/method': method, 'parameters/uri': uri, 'parameters/headers': headers}
    if body is not None:
        params['parameters/body'] = body
    return op('sp', 'HttpRequest', params, run_after)


def set_var(name, value, run_after=None):
    a = {'type': 'SetVariable', 'inputs': {'name': name, 'value': value}}
    if run_after is not None:
        a['runAfter'] = run_after
    return a


def append_var(name, value, run_after=None):
    a = {'type': 'AppendToStringVariable', 'inputs': {'name': name, 'value': value}}
    if run_after is not None:
        a['runAfter'] = run_after
    return a


def err_of(action):
    """Fehlermeldung einer fehlgeschlagenen Aktion."""
    return (f"@{{coalesce(body('{action}')?['error']?['message'], body('{action}')?['message'], "
            f"actions('{action}')?['error']?['message'], 'Aktion {action} ist fehlgeschlagen.')}}")


def guarded(steps, ok_text, coe_steps=None):
    """
    Folge von Admin-Aktionen; jede Aktion bekommt einen Fehler-Merker. Nach Erfolg: Ergebnistext setzen und
    optional das CoE-Inventar nachziehen (Fehler dort nur als Hinweis).
    steps: [(name, action)] – die Aktionen laufen nacheinander.
    """
    acts = {}
    prev = None
    for name, action in steps:
        action = dict(action)
        if prev:
            action['runAfter'] = {prev: ['Succeeded']}
        else:
            action.pop('runAfter', None)
        acts[name] = action
        acts[f'{name}_Fehler'] = set_var('Fehler', err_of(name), {name: ['Failed', 'TimedOut']})
        prev = name
    acts['Ergebnis_setzen'] = set_var('Ergebnis', ok_text, {prev: ['Succeeded']})
    last = 'Ergebnis_setzen'
    for name, action in (coe_steps or []):
        action = dict(action)
        action['runAfter'] = {last: ['Succeeded']}
        acts[name] = action
        acts[f'{name}_Hinweis'] = append_var('Hinweis', ' (CoE-Inventar nicht aktualisiert: ' + err_of(name) + ')', {name: ['Failed', 'TimedOut']})
        last = name
    return acts


# Werte aus dem Auftrag
ACTION = "@triggerOutputs()?['body/Action']"
RTYPE = "triggerOutputs()?['body/ResourceType']"
RID = "@triggerOutputs()?['body/ResourceId']"
COEID = "@triggerOutputs()?['body/CoeRecordId']"


def is_type(t):
    return {'equals': [f'@{RTYPE}', t]}


def _containers(node):
    """Alle Aktions-Container (dicts mit Aktionen) unterhalb eines Aktions-Dicts."""
    yield node
    for a in node.values():
        if not isinstance(a, dict):
            continue
        if 'actions' in a and isinstance(a['actions'], dict):
            yield from _containers(a['actions'])
        if isinstance(a.get('else'), dict) and isinstance(a['else'].get('actions'), dict):
            yield from _containers(a['else']['actions'])
        for c in (a.get('cases') or {}).values():
            yield from _containers(c.get('actions', {}))
        if isinstance(a.get('default'), dict):
            yield from _containers(a['default'].get('actions', {}))


def unique_names(actions, suffix):
    """Logic Apps verlangt eindeutige Aktionsnamen im ganzen Flow → Namen je Fall mit Suffix versehen."""
    names = set()
    for c in _containers(actions):
        names.update(c.keys())
    mapping = {n: f'{n}_{suffix}' for n in names}

    def fix(x):
        if isinstance(x, str):
            for n in sorted(mapping, key=len, reverse=True):
                x = x.replace(f"('{n}')", f"('{mapping[n]}')").replace(f"Aktion {n} ", f"Aktion {mapping[n]} ")
            return x
        if isinstance(x, list):
            return [fix(v) for v in x]
        if isinstance(x, dict):
            out = {}
            for k, v in x.items():
                if k == 'runAfter' and isinstance(v, dict):
                    out[k] = {mapping.get(n, n): st for n, st in v.items()}
                else:
                    out[k] = fix(v)
            return out
        return x

    def rename(container):
        out = {}
        for k, a in container.items():
            a = fix(a)
            if isinstance(a, dict):
                if isinstance(a.get('actions'), dict):
                    a['actions'] = rename(a['actions'])
                if isinstance(a.get('else'), dict) and isinstance(a['else'].get('actions'), dict):
                    a['else']['actions'] = rename(a['else']['actions'])
                for c in (a.get('cases') or {}).values():
                    c['actions'] = rename(c.get('actions', {}))
                if isinstance(a.get('default'), dict):
                    a['default']['actions'] = rename(a['default'].get('actions', {}))
            out[mapping.get(k, k)] = a
        return out

    return rename(actions)


# ---------- Fälle ----------
def case_soldelete():
    """Nicht verwaltete Solution löschen (nur der Container – Komponenten bleiben). EnvironmentId = Dataverse-URL des Environments."""
    org = "@triggerOutputs()?['body/EnvironmentId']"
    acts = {
        'Solution_lesen': op('dv', 'GetItemWithOrganization', {
            'organization': org, 'entityName': 'solutions', 'recordId': RID, '$select': 'uniquename,ismanaged,isvisible'}),
        'Solution_lesen_Fehler': set_var('Fehler', err_of('Solution_lesen'), {'Solution_lesen': ['Failed', 'TimedOut']}),
        'Loeschen_erlaubt': {
            'type': 'If',
            'runAfter': {'Solution_lesen': ['Succeeded']},
            'expression': {'and': [
                {'equals': ["@body('Solution_lesen')?['ismanaged']", False]},
                {'not': {'contains': ["@createArray('default','active','basic','system','crmservicesolution')", "@toLower(coalesce(body('Solution_lesen')?['uniquename'], ''))"]}}
            ]},
            'actions': guarded(
                [('Solution_loeschen', op('dv', 'DeleteRecordWithOrganization', {'organization': org, 'entityName': 'solutions', 'recordId': RID}, retry=True))],
                'Die Solution wurde gelöscht. Ihre Komponenten (Apps, Flows, Tabellen …) bleiben im Environment erhalten.',
                [('CoE_Solution_geloescht', op('dv', 'UpdateRecord', {
                    'entityName': 'admin_solutions', 'recordId': COEID, 'item/admin_solutiondeleted': True, 'item/admin_solutiondeletedon': '@utcNow()'}))]),
            'else': {'actions': {
                'Abgelehnt_verwaltet': set_var('Abgelehnt', 'Verwaltete Solutions und System-Solutions werden nicht gelöscht (Deinstallation nur über das Admin Center bzw. ALM).')
            }}
        }
    }
    return {'Nur_Solution': {
        'type': 'If', 'expression': is_type('solution'), 'actions': acts,
        'else': {'actions': {'Abgelehnt_keine_Solution': set_var('Abgelehnt', 'Diese Aktion gibt es nur für Solutions.')}}
    }}


CASES = {
    'soldelete': case_soldelete(),
}


def strip_none(x):
    if isinstance(x, dict):
        return {k: strip_none(v) for k, v in x.items() if v is not None}
    if isinstance(x, list):
        return [strip_none(v) for v in x]
    return x


def definition():
    # Liste über Titel oder ID (GUID) ansprechen
    list_seg = (f"@{{if(equals(length(trim({P_LIST})), 36), concat('lists(guid''', trim({P_LIST}), ''')'), "
                f"concat('lists/getbytitle(''', replace(trim({P_LIST}), '''', ''''''), ''')'))}}")
    item_uri = f"_api/web/{list_seg}/items(@{{triggerOutputs()?['body/ID']}})"

    switch = {
        'type': 'Switch',
        'expression': ACTION,
        'cases': {f'Fall_{k}': {'case': k, 'actions': unique_names(v, k)} for k, v in CASES.items()},
        'default': {'actions': {'Nicht_im_Flow': set_var('Abgelehnt', "@{concat('Die Aktion „', triggerOutputs()?['body/Action'], '“ führt das CoE Cockpit (ab Version 1.14) direkt aus, nicht dieser Flow. Bitte das Webpart aktualisieren.')}")}},
        'runAfter': {}
    }

    member_check = {
        'type': 'If',
        'runAfter': {'Mitglied_pruefen': ['Succeeded']},
        'expression': {'greater': ["@length(body('Mitglied_pruefen'))", 0]},
        'actions': {
            'Aktion': switch
        },
        'else': {'actions': {
            'Nicht_Mitglied': set_var('Abgelehnt', f"@{{concat(triggerOutputs()?['body/Author/DisplayName'], ' ist nicht Mitglied der SharePoint-Gruppe „', {P_GROUP}, '“. Der Auftrag wurde nicht ausgeführt.')}}")
        }}
    }

    scope = {
        'type': 'Scope',
        'runAfter': {'Status_In_Arbeit': ['Succeeded', 'Failed']},
        'actions': {
            'Gruppenmitglieder_lesen': sp_http('GET', f"_api/web/sitegroups/getbyname('@{{{P_GROUP}}}')/users?$select=LoginName,Email&$top=5000"),
            'Mitglied_pruefen': {
                'type': 'Query',
                'runAfter': {'Gruppenmitglieder_lesen': ['Succeeded']},
                'inputs': {
                    'from': "@coalesce(body('Gruppenmitglieder_lesen')?['value'], body('Gruppenmitglieder_lesen')?['d']?['results'], createArray())",
                    'where': "@or(equals(toLower(item()?['LoginName']), toLower(triggerOutputs()?['body/Author/Claims'])), and(not(empty(item()?['Email'])), equals(toLower(item()?['Email']), toLower(triggerOutputs()?['body/Author/Email']))))"
                }
            },
            'Mitgliedschaft': member_check
        }
    }

    end_status = ("if(not(empty(variables('Abgelehnt'))), 'Abgelehnt', if(not(empty(variables('Fehler'))), 'Fehler', "
                  "if(empty(variables('Ergebnis')), 'Fehler', 'Erledigt')))")
    end_text = ("if(not(empty(variables('Abgelehnt'))), variables('Abgelehnt'), if(not(empty(variables('Fehler'))), variables('Fehler'), "
                "if(empty(variables('Ergebnis')), 'Unerwarteter Fehler im Flow – Details im Ausführungsverlauf von „CoE Cockpit | Aktion ausführen“.', "
                "concat(variables('Ergebnis'), variables('Hinweis')))))")
    end_body = f"@string(addProperty(addProperty(addProperty(json('{{}}'), 'Status', {end_status}), 'Result', {end_text}), 'CompletedAt', utcNow()))"

    actions = {
        'Ergebnis_init': {'type': 'InitializeVariable', 'inputs': {'variables': [{'name': 'Ergebnis', 'type': 'string', 'value': ''}]}, 'runAfter': {}},
        'Fehler_init': {'type': 'InitializeVariable', 'inputs': {'variables': [{'name': 'Fehler', 'type': 'string', 'value': ''}]}, 'runAfter': {'Ergebnis_init': ['Succeeded']}},
        'Abgelehnt_init': {'type': 'InitializeVariable', 'inputs': {'variables': [{'name': 'Abgelehnt', 'type': 'string', 'value': ''}]}, 'runAfter': {'Fehler_init': ['Succeeded']}},
        'Hinweis_init': {'type': 'InitializeVariable', 'inputs': {'variables': [{'name': 'Hinweis', 'type': 'string', 'value': ''}]}, 'runAfter': {'Abgelehnt_init': ['Succeeded']}},
        'Status_In_Arbeit': sp_http('POST', item_uri, '{"Status": "In Arbeit"}', {'Hinweis_init': ['Succeeded']}, merge=True),
        'Ausfuehrung': scope,
        'Status_zurueckschreiben': sp_http('POST', item_uri, end_body, {'Ausfuehrung': ['Succeeded', 'Failed', 'Skipped', 'TimedOut']}, merge=True)
    }
    # Fehler im Scope (z. B. Gruppe nicht gefunden) als Fehlertext übernehmen
    scope['actions']['Gruppe_nicht_lesbar'] = set_var(
        'Fehler', f"@{{concat('Die SharePoint-Gruppe „', {P_GROUP}, '“ konnte nicht gelesen werden: ', coalesce(body('Gruppenmitglieder_lesen')?['error']?['message'], body('Gruppenmitglieder_lesen')?['odata.error']?['message']?['value'], 'unbekannter Fehler'))}}",
        {'Gruppenmitglieder_lesen': ['Failed', 'TimedOut']})

    trigger = {
        'Wenn_ein_Auftrag_angelegt_wird': {
            'recurrence': {'interval': 1, 'frequency': 'Minute'},
            'splitOn': "@triggerOutputs()?['body/value']",
            'type': 'OpenApiConnection',
            'inputs': {
                'host': {'connectionName': CONN['sp'][0], 'operationId': 'GetOnNewItems', 'apiId': api('sp')},
                'parameters': {'dataset': P_SITE, 'table': f'@{P_LIST}'},
                'authentication': "@parameters('$authentication')"
            },
            'description': 'Neuer Auftrag aus dem CoE Cockpit (Liste „CoE Cockpit Aktionen“). Nur Status „Angefordert“ – direkt im Webpart ausgeführte Aktionen stehen schon auf „In Arbeit“.',
            'conditions': [{'expression': "@or(empty(triggerOutputs()?['body/Status']), equals(triggerOutputs()?['body/Status'], 'Angefordert'))"}]
        }
    }

    params = {
        '$connections': {'defaultValue': {}, 'type': 'Object'},
        '$authentication': {'defaultValue': {}, 'type': 'SecureObject'},
    }
    for schema, display, desc, default in ENV_VARS:
        params[f'{display} ({schema})'] = {'defaultValue': default or '', 'type': 'String', 'metadata': {'schemaName': schema, 'description': desc}}

    return strip_none({
        'properties': {
            'connectionReferences': {
                CONN[k][0]: {'runtimeSource': 'embedded', 'connection': {'connectionReferenceLogicalName': CONN[k][1]}, 'api': {'name': CONN[k][0]}}
                for k in CONN
            },
            'definition': {
                '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
                'contentVersion': '1.0.0.0',
                'parameters': params,
                'triggers': trigger,
                'actions': actions,
                'outputs': {}
            },
            'templateName': ''
        },
        'schemaVersion': '1.0.0.0'
    })


# ---------- Lösungsdateien ----------
CONTENT_TYPES = ('<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="xml" ContentType="application/octet-stream" /><Default Extension="json" ContentType="application/octet-stream" /></Types>')

ADDRESS = ''.join(f'<{t} xsi:nil="true"></{t}>' for t in [
    'City', 'County', 'Country', 'Fax', 'FreightTermsCode', 'ImportSequenceNumber', 'Latitude', 'Line1', 'Line2', 'Line3', 'Longitude',
    'Name', 'PostalCode', 'PostOfficeBox', 'PrimaryContactName'])


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def solution_xml():
    addr = (f'<Address><AddressNumber>{{n}}</AddressNumber><AddressTypeCode>1</AddressTypeCode>{ADDRESS}'
            '<ShippingMethodCode>1</ShippingMethodCode><StateOrProvince xsi:nil="true"></StateOrProvince><Telephone1 xsi:nil="true"></Telephone1>'
            '<Telephone2 xsi:nil="true"></Telephone2><Telephone3 xsi:nil="true"></Telephone3><TimeZoneRuleVersionNumber xsi:nil="true"></TimeZoneRuleVersionNumber>'
            '<UPSZone xsi:nil="true"></UPSZone><UTCOffset xsi:nil="true"></UTCOffset><UTCConversionTimeZoneCode xsi:nil="true"></UTCConversionTimeZoneCode></Address>')
    return f'''<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.25123.166" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrganizationVersion="9.2.25123.166" OrganizationSchemaType="Standard" CRMServerServiceabilityVersion="9.2.25123.00166">
  <SolutionManifest>
    <UniqueName>{SOLUTION}</UniqueName>
    <LocalizedNames>
      <LocalizedName description="CoE Cockpit Aktionen" languagecode="1033" />
    </LocalizedNames>
    <Descriptions>
      <Description description="Führt die Aufträge des SharePoint-Webparts „CoE Cockpit“ aus, die Rechte im Environment brauchen (derzeit „Solution löschen“). Alle anderen Aktionen führt das Webpart selbst aus." languagecode="1033" />
    </Descriptions>
    <Version>{VERSION}</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>coecockpit</UniqueName>
      <LocalizedNames>
        <LocalizedName description="CoE Cockpit" languagecode="1033" />
      </LocalizedNames>
      <Descriptions>
        <Description description="CoE Cockpit" languagecode="1033" />
      </Descriptions>
      <EMailAddress xsi:nil="true"></EMailAddress>
      <SupportingWebsiteUrl xsi:nil="true"></SupportingWebsiteUrl>
      <CustomizationPrefix>coeck</CustomizationPrefix>
      <CustomizationOptionValuePrefix>47121</CustomizationOptionValuePrefix>
      <Addresses>
        {addr.replace('{n}', '1')}
        {addr.replace('{n}', '2')}
      </Addresses>
    </Publisher>
    <RootComponents>
      <RootComponent type="29" id="{{{FLOW_ID}}}" behavior="0" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
'''


def customizations_xml():
    refs = '\n'.join(f'''    <connectionreference connectionreferencelogicalname="{CONN[k][1]}">
      <connectionreferencedisplayname>{esc(CONN[k][2])}</connectionreferencedisplayname>
      <connectorid>{api(k)}</connectorid>
      <description>{esc(CONN[k][3])}</description>
      <iscustomizable>1</iscustomizable>
      <promptingbehavior>0</promptingbehavior>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>''' for k in CONN)
    desc = 'Führt Aufträge mit Status „Angefordert“ aus der SharePoint-Liste „CoE Cockpit Aktionen“ aus, die Rechte im Environment brauchen (derzeit „Solution löschen“).'
    return f'''<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrganizationVersion="9.2.25123.166" OrganizationSchemaType="Standard" CRMServerServiceabilityVersion="9.2.25123.00166">
  <Entities />
  <Roles />
  <Workflows>
    <Workflow WorkflowId="{{{FLOW_ID}}}" Name="{esc(FLOW_NAME)}" Description="{esc(desc)}">
      <JsonFileName>/Workflows/{FLOW_FILE}</JsonFileName>
      <Type>1</Type>
      <Subprocess>0</Subprocess>
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>0</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <StateCode>1</StateCode>
      <StatusCode>2</StatusCode>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>1.0.0.0</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode="1033" description="{esc(FLOW_NAME)}" />
      </LocalizedNames>
      <Descriptions>
        <Description languagecode="1033" description="{esc(desc)}" />
      </Descriptions>
    </Workflow>
  </Workflows>
  <FieldSecurityProfiles />
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls />
  <EntityDataProviders />
  <connectionreferences>
{refs}
  </connectionreferences>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>
'''


def env_var_xml(schema, display, desc, default):
    d = f'\n  <defaultvalue>{esc(default)}</defaultvalue>' if default else ''
    return f'''<?xml version="1.0" encoding="utf-8"?>
<environmentvariabledefinition schemaname="{schema}">{d}
  <description default="{esc(desc)}">
    <label description="{esc(desc)}" languagecode="1033" />
  </description>
  <displayname default="{esc(display)}">
    <label description="{esc(display)}" languagecode="1033" />
  </displayname>
  <introducedversion>1.0.0.0</introducedversion>
  <iscustomizable>1</iscustomizable>
  <isrequired>1</isrequired>
  <secretstore>0</secretstore>
  <type>100000000</type>
</environmentvariabledefinition>
'''


def main():
    flow = definition()
    out = os.path.join(HERE, f'{SOLUTION}_{VERSION.replace(".", "_")}.zip')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('solution.xml', solution_xml())
        z.writestr('customizations.xml', customizations_xml())
        z.writestr(f'Workflows/{FLOW_FILE}', json.dumps(flow, ensure_ascii=False, indent=2))
        for schema, display, desc, default in ENV_VARS:
            z.writestr(f'environmentvariabledefinitions/{schema}/environmentvariabledefinition.xml', env_var_xml(schema, display, desc, default))
    with open(out, 'wb') as f:
        f.write(buf.getvalue())
    # Flow-Definition zusätzlich lesbar ablegen (für Prüfung/Vergleich)
    with open(os.path.join(HERE, 'CoECockpit-Aktionausfuehren.definition.json'), 'w', encoding='utf-8') as f:
        json.dump(flow, f, ensure_ascii=False, indent=2)
    print('OK', out, len(buf.getvalue()), 'Bytes')


if __name__ == '__main__':
    main()
