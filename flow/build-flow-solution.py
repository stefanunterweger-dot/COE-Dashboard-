#!/usr/bin/env python3
"""
Erzeugt die importierbare Power-Platform-Lösung „CoE Cockpit Aktionen“ (nicht verwaltet) mit dem Flow
„CoE Cockpit | Aktion ausführen“. Aufruf:  python3 flow/build-flow-solution.py   → flow/CoECockpitAktionen_1_0_0_0.zip

Der Flow arbeitet die Aufträge ab, die das Webpart in der SharePoint-Liste „CoE Cockpit Aktionen“ anlegt:
  1. Status „In Arbeit“
  2. Prüfen, ob die anfragende Person (Erstellt von) Mitglied der SharePoint-Gruppe ist – sonst „Abgelehnt“
  3. Aktion ausführen (Admin-Connectors, Verbindungen eines Admin-Dienstkontos)
  4. CoE-Inventar nachziehen (Quarantäne, gelöscht, Besitzer, Flow-Status) – Fehler dabei sind nur ein Hinweis
  5. Status „Erledigt“, „Fehler“ oder „Abgelehnt“ mit Ergebnistext und Zeitpunkt zurückschreiben
"""
import io
import json
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
VERSION = '1.0.0.0'
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
    'apps': ('shared_powerappsforadmins', 'coeck_PowerAppsAdmin', 'CoE Cockpit – Power Apps for Admins', 'Quarantäne, Besitzer, Freigaben und Löschen von Canvas-Apps (Admin-Dienstkonto).'),
    'flows': ('shared_microsoftflowforadmins', 'coeck_PowerAutomateAdmin', 'CoE Cockpit – Power Automate for Admins', 'Mitbesitzer, Freigaben und Löschen von Flows (Admin-Dienstkonto).'),
    'flowmgmt': ('shared_flowmanagement', 'coeck_PowerAutomateManagement', 'CoE Cockpit – Power Automate Management', 'Flows ein- und ausschalten (Admin-Dienstkonto).'),
    'ppv2': ('shared_powerplatformadminv2', 'coeck_PowerPlatformAdminV2', 'CoE Cockpit – Power Platform for Admins V2', 'Environments auf Managed umstellen, deaktivieren und aktivieren (Admin-Dienstkonto).'),
    'dv': ('shared_commondataserviceforapps', 'coeck_Dataverse', 'CoE Cockpit – Dataverse (CoE)', 'Liest und aktualisiert das CoE-Inventar im aktuellen Environment.'),
    'users': ('shared_office365users', 'coeck_Office365Users', 'CoE Cockpit – Office 365 Users', 'Ermittelt die Entra-Objekt-ID zu einer E-Mail-Adresse.'),
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
T = "triggerOutputs()?['body/{}']"
ACTION = "@triggerOutputs()?['body/Action']"
RTYPE = "triggerOutputs()?['body/ResourceType']"
RID = "@triggerOutputs()?['body/ResourceId']"
ENV = "@triggerOutputs()?['body/EnvironmentId']"
COEID = "@triggerOutputs()?['body/CoeRecordId']"
ZIEL = "@variables('Ziel')"
PTYPE = "@{if(equals(triggerOutputs()?['body/TargetIsGroup'], true), 'group', 'user')}"


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
def case_quarantine(on):
    status = 'Quarantined' if on else 'Unquarantined'
    return guarded(
        [('Quarantaene_setzen', op('apps', 'Set-AppQuarantineState', {
            'environment': ENV, 'app': RID, 'api-version': '2021-02-01', 'Content-Type': 'application/json', 'body/quarantineStatus': status}, retry=True))],
        'Die App ist jetzt in Quarantäne.' if on else 'Die App ist nicht mehr in Quarantäne.',
        [('CoE_Quarantaene', op('dv', 'UpdateRecord', {
            'entityName': 'admin_apps', 'recordId': COEID, 'item/admin_quarantineapp': on,
            'item/admin_quarantineappdate': "@utcNow()" if on else None}))])


def solution_check(entity, field, label):
    """Löschen nur, wenn die Ressource laut CoE NICHT Teil einer Solution ist (unbekannt = ablehnen)."""
    return {
        'CoE_Solution_lesen': op('dv', 'GetItem', {'entityName': entity, 'recordId': COEID, '$select': field}),
        'CoE_Solution_lesen_Fehler': set_var('Fehler', err_of('CoE_Solution_lesen'), {'CoE_Solution_lesen': ['Failed', 'TimedOut']}),
        'Nicht_in_Solution': {
            'type': 'If',
            'runAfter': {'CoE_Solution_lesen': ['Succeeded']},
            'expression': {'equals': [f"@body('CoE_Solution_lesen')?['{field}']", False]},
            'actions': {},
            'else': {'actions': {
                'Abgelehnt_Solution': set_var('Abgelehnt', f'{label} ist Teil einer Solution (oder das ist im CoE nicht erfasst) – Löschen nur über die Solution (ALM).')
            }}
        }
    }


def case_delete():
    app = solution_check('admin_apps', 'admin_issolutionapp', 'Die App')
    app['Nicht_in_Solution']['actions'] = guarded(
        [('App_loeschen', op('apps', 'Remove-AdminApp', {'environment': ENV, 'app': RID, 'api-version': '2016-11-01'}, retry=True))],
        'Die App wurde im Tenant gelöscht.',
        [('CoE_App_geloescht', op('dv', 'UpdateRecord', {'entityName': 'admin_apps', 'recordId': COEID, 'item/admin_appdeleted': True, 'item/admin_appdeletedon': '@utcNow()'}))])
    flow = solution_check('admin_flows', 'admin_issolutionflow', 'Der Flow')
    flow['Nicht_in_Solution']['actions'] = guarded(
        [('Flow_loeschen', op('flows', 'Remove-AdminFlow', {'environment': ENV, 'flow': RID, 'api-version': '2016-11-01'}, retry=True))],
        'Der Flow wurde im Tenant gelöscht.',
        [('CoE_Flow_geloescht', op('dv', 'UpdateRecord', {'entityName': 'admin_flows', 'recordId': COEID, 'item/admin_flowdeleted': True, 'item/admin_flowdeletedon': '@utcNow()'}))])
    app = unique_names(app, 'App')
    flow = unique_names(flow, 'Flow')
    return {'Loeschen_App_oder_Flow': {
        'type': 'If', 'expression': is_type('app'),
        'actions': app,
        'else': {'actions': {'Loeschen_Flow_pruefen': {
            'type': 'If', 'expression': is_type('flow'), 'actions': flow,
            'else': {'actions': {'Abgelehnt_Typ': set_var('Abgelehnt', 'Löschen ist nur für Apps und Flows vorgesehen.')}}
        }}}
    }}


def app_or_flow(app_acts, flow_acts):
    app_acts = unique_names(app_acts, 'App')
    flow_acts = unique_names(flow_acts, 'Flow')
    return {'App_oder_Flow': {
        'type': 'If', 'expression': is_type('app'), 'actions': app_acts,
        'else': {'actions': {'Nur_Flow': {
            'type': 'If', 'expression': is_type('flow'), 'actions': flow_acts,
            'else': {'actions': {'Abgelehnt_Ressource': set_var('Abgelehnt', 'Diese Aktion gibt es nur für Apps und Flows.')}}
        }}}
    }}


def case_setowner():
    app = guarded(
        [('App_Besitzer', op('apps', 'Set-AdminAppOwner', {
            'environment': ENV, 'app': RID, 'api-version': '2016-11-01', 'Content-Type': 'application/json',
            'body/roleForOldAppOwner': 'CanEdit', 'body/newAppOwner': ZIEL}, retry=True))],
        "@{concat(triggerOutputs()?['body/TargetName'], ' ist jetzt Besitzer der App; der bisherige Besitzer ist Mitbesitzer.')}",
        [('CoE_App_Besitzer', op('dv', 'UpdateRecord', {'entityName': 'admin_apps', 'recordId': COEID, 'item/admin_AppOwner@odata.bind': "admin_makers(@{variables('Ziel')})"}))])
    flow = guarded(
        [('Flow_Mitbesitzer', op('flows', 'Edit-AdminFlowOwnerRole', {
            'environment': ENV, 'flow': RID, 'api-version': '2016-11-01', 'Content-Type': 'application/json',
            'body/put': [{'properties/principal/id': ZIEL, 'properties/principal/type': PTYPE, 'properties/roleName': 'CanEdit'}]}, retry=True))],
        "@{concat(triggerOutputs()?['body/TargetName'], ' ist jetzt Mitbesitzer des Flows. Den Ersteller eines Flows außerhalb einer Solution kann Power Automate nicht ändern.')}",
        [('CoE_Flow_Besitzer', op('dv', 'UpdateRecord', {'entityName': 'admin_flows', 'recordId': COEID, 'item/admin_DerivedOwner@odata.bind': "admin_makers(@{variables('Ziel')})"}))])
    return app_or_flow(app, flow)


def case_share(role):
    app = guarded(
        [('App_Freigabe', op('apps', 'Edit-AdminAppRoleAssignment', {
            'environment': ENV, 'app': RID, 'api-version': '2016-11-01',
            'body/put': [{'properties/roleName': role, 'properties/principal/id': ZIEL, 'properties/principal/type': PTYPE, 'properties/principal/tenantId': '@null'}]}, retry=True))],
        "@{concat('Freigabe für ', triggerOutputs()?['body/TargetName'], ' hinzugefügt (" + ('Mitbesitzer' if role == 'CanEdit' else 'Benutzer') + ").')}")
    flow = guarded(
        [('Flow_Freigabe', op('flows', 'Edit-AdminFlowOwnerRole', {
            'environment': ENV, 'flow': RID, 'api-version': '2016-11-01', 'Content-Type': 'application/json',
            'body/put': [{'properties/principal/id': ZIEL, 'properties/principal/type': PTYPE, 'properties/roleName': 'CanEdit'}]}, retry=True))],
        "@{concat(triggerOutputs()?['body/TargetName'], ' ist jetzt Mitbesitzer des Flows.')}")
    return app_or_flow(app, flow)


def case_removeuser():
    app = guarded(
        [('App_Freigabe_entfernen', op('apps', 'Edit-AdminAppRoleAssignment', {
            'environment': ENV, 'app': RID, 'api-version': '2016-11-01', 'body/delete': [{'id': ZIEL}]}, retry=True))],
        "@{concat('Freigabe für ', triggerOutputs()?['body/TargetName'], ' entfernt.')}")
    flow = guarded(
        [('Flow_Freigabe_entfernen', op('flows', 'Edit-AdminFlowOwnerRole', {
            'environment': ENV, 'flow': RID, 'api-version': '2016-11-01', 'Content-Type': 'application/json', 'body/delete': [{'id': ZIEL}]}, retry=True))],
        "@{concat('Freigabe für ', triggerOutputs()?['body/TargetName'], ' entfernt.')}")
    return app_or_flow(app, flow)


def case_flow_state(on):
    return guarded(
        [('Flow_' + ('einschalten' if on else 'ausschalten'), op('flowmgmt', 'StartFlow' if on else 'StopFlow', {'environmentName': ENV, 'flowName': RID}, retry=True))],
        'Der Flow ist eingeschaltet.' if on else 'Der Flow ist ausgeschaltet.',
        [('CoE_Flow_Status', op('dv', 'UpdateRecord', {'entityName': 'admin_flows', 'recordId': COEID, 'item/admin_flowstate': 'Started' if on else 'Stopped'}))])


def case_env(kind):
    if kind == 'managed':
        action = op('ppv2', 'EnableManagedEnvironment', {'environmentId': ENV, 'api-version': '2024-10-01'}, retry=True)
        text = 'Das Environment wird zum Managed Environment (Vorgang läuft bei Microsoft weiter, Dauer wenige Minuten).'
    elif kind == 'disable':
        action = op('ppv2', 'DisableEnvironment', {'environmentId': ENV, 'api-version': '2024-10-01', 'body/reason': 'UserRequested'}, retry=True)
        text = 'Das Environment wird deaktiviert (Vorgang läuft bei Microsoft weiter, Dauer wenige Minuten).'
    else:
        action = op('ppv2', 'EnableEnvironment', {'environmentId': ENV, 'api-version': '2024-10-01'}, retry=True)
        text = 'Das Environment wird wieder aktiviert (Vorgang läuft bei Microsoft weiter, Dauer wenige Minuten).'
    return {'Nur_Environment': {
        'type': 'If', 'expression': is_type('environment'),
        'actions': guarded([('Environment_' + kind, action)], text),
        'else': {'actions': {'Abgelehnt_kein_Environment': set_var('Abgelehnt', 'Diese Aktion gibt es nur für Environments.')}}
    }}


CASES = {
    'quarantine': case_quarantine(True),
    'unquarantine': case_quarantine(False),
    'delete': case_delete(),
    'setowner': case_setowner(),
    'addviewer': case_share('CanView'),
    'addeditor': case_share('CanEdit'),
    'removeuser': case_removeuser(),
    'flowoff': case_flow_state(False),
    'flowon': case_flow_state(True),
    'envmanaged': case_env('managed'),
    'envdisable': case_env('disable'),
    'envenable': case_env('enable'),
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
    needs_target = "@contains(createArray('setowner','addviewer','addeditor','removeuser'), triggerOutputs()?['body/Action'])"

    switch = {
        'type': 'Switch',
        'expression': ACTION,
        'cases': {f'Fall_{k}': {'case': k, 'actions': unique_names(v, k)} for k, v in CASES.items()},
        'default': {'actions': {'Unbekannte_Aktion': set_var('Abgelehnt', "@{concat('Unbekannte Aktion: ', triggerOutputs()?['body/Action'])}")}},
        'runAfter': {'Ziel_ermitteln': ['Succeeded']}
    }

    ziel = {
        'type': 'If',
        'expression': {'and': [
            {'equals': [needs_target, True]},
            {'not': {'equals': ["@length(variables('Ziel'))", 36]}}
        ]},
        'actions': {
            'Person_suchen': op('users', 'UserProfile_V2', {'id': "@variables('Ziel')"}),
            'Ziel_ID_setzen': set_var('Ziel', "@body('Person_suchen')?['id']", {'Person_suchen': ['Succeeded']}),
            'Person_nicht_gefunden': set_var('Abgelehnt', "@{concat('Die Person „', variables('Ziel'), '“ wurde nicht gefunden. Gruppen bitte mit ihrer Objekt-ID angeben.')}",
                                             {'Person_suchen': ['Failed', 'TimedOut']})
        },
        'else': {'actions': {}}
    }

    member_check = {
        'type': 'If',
        'runAfter': {'Mitglied_pruefen': ['Succeeded']},
        'expression': {'greater': ["@length(body('Mitglied_pruefen'))", 0]},
        'actions': {
            'Ziel_ermitteln': ziel,
            'Aktion_ausfuehren': {
                'type': 'If',
                'runAfter': {'Ziel_ermitteln': ['Succeeded', 'Failed']},
                'expression': {'equals': ["@empty(variables('Abgelehnt'))", True]},
                'actions': {'Aktion': dict(switch, runAfter={})},
                'else': {'actions': {}}
            }
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
        'Ziel_init': {'type': 'InitializeVariable', 'inputs': {'variables': [{'name': 'Ziel', 'type': 'string', 'value': "@{trim(coalesce(triggerOutputs()?['body/TargetUser'], ''))}"}]}, 'runAfter': {'Hinweis_init': ['Succeeded']}},
        'Status_In_Arbeit': sp_http('POST', item_uri, '{"Status": "In Arbeit"}', {'Ziel_init': ['Succeeded']}, merge=True),
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
            'description': 'Neuer Auftrag aus dem CoE Cockpit (Liste „CoE Cockpit Aktionen“).'
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
      <Description description="Führt die Aufträge des SharePoint-Webparts „CoE Cockpit“ aus (Liste „CoE Cockpit Aktionen“)." languagecode="1033" />
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
    desc = 'Führt Aufträge aus der SharePoint-Liste „CoE Cockpit Aktionen“ aus (Quarantäne, Löschen, Besitzer, Freigaben, Flow ein/aus, Environment Managed/deaktivieren/aktivieren).'
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
