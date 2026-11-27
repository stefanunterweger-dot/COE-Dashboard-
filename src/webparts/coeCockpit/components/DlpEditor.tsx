import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { ChoiceGroup } from '@fluentui/react/lib/ChoiceGroup';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable, Tile } from './CopilotCreditsView';
import {
  addTask, computeImpact, connectorId, diffConnectors, DLP_GROUPS, DlpApi, DlpGroup, DlpScope, GROUP_LABEL, groupOf, IDlpBase, IDlpConnector,
  IDlpImpact, IDlpImpactResult, IDlpWorking, isDirty, loadInventory, loadTasks, ppacPolicyUrl, saveDraft, SCOPE_LABEL, toDefinition
} from '../models/dlp';
import { IActionRequest } from '../models/actions';
import { ActionService } from '../services/ActionService';
import { DataverseService } from '../services/DataverseService';
import { DirectActionError } from '../services/DirectActionService';
import { PIM_URL } from '../services/PowerPlatformApiService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * Editor der DLP-Auswirkungsanalyse (wie „Edit Policy“ in der CoE-App): Name → Connectors → Geltungsbereich → Environments →
 * Auswirkung → Überprüfen. Speichern als Entwurf (CoE) oder – für Admins mit aktiver Rolle – direkt im Tenant (mit Protokoll).
 */
export interface IDlpEditorProps {
  working: IDlpWorking;
  base: IDlpBase;
  unblockable: string[];
  api: DlpApi;
  service: DataverseService;
  maxRecords: number;
  canAct: boolean;
  actHint: string;
  actions: ActionService | undefined;
  requestedBy: string;
  onOpen: (logicalName: string, id: string) => void;
  /** changed = im Tenant oder als Entwurf gespeichert */
  onClose: (changed: boolean) => void;
  onActionsChanged: () => void;
}

const STEPS = ['Name', 'Connectors', 'Geltungsbereich', 'Environments', 'Auswirkung', 'Überprüfen'];
const MAX_TASKS = 500;
const PAGE = 200;

const CHIP: { [g in DlpGroup]: string } = { Confidential: styles.dlpChipBusiness, General: styles.dlpChipGeneral, Blocked: styles.dlpChipBlocked };
const GROUP_ICON: { [g in DlpGroup]: string } = { Confidential: 'Lock', General: 'Unlock', Blocked: 'Blocked2' };
const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));
const csvEsc = (v: string): string => (/[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);

type SaveKind = 'update' | 'create';

export const DlpEditor: React.FC<IDlpEditorProps> = props => {
  const { base } = props;
  const [w, setW] = React.useState<IDlpWorking>(props.working);
  const [step, setStep] = React.useState<number>(0);
  const [visited, setVisited] = React.useState<{ [s: number]: true }>({ 0: true });
  const [group, setGroup] = React.useState<DlpGroup>('Confidential');
  const [conSearch, setConSearch] = React.useState<string>('');
  const [conSel, setConSel] = React.useState<{ [id: string]: true }>({});
  const [conLimit, setConLimit] = React.useState<number>(PAGE);
  const [envSearch, setEnvSearch] = React.useState<[string, string]>(['', '']);
  const [envSel, setEnvSel] = React.useState<{ [name: string]: true }>({});
  const [impact, setImpact] = React.useState<IDlpImpactResult | undefined>(undefined);
  const [impactKey, setImpactKey] = React.useState<string>('');
  const [impactLoading, setImpactLoading] = React.useState<boolean>(false);
  const [impactError, setImpactError] = React.useState<string>('');
  const [impactKind, setImpactKind] = React.useState<'app' | 'flow'>('app');
  const [onlyNew, setOnlyNew] = React.useState<boolean>(false);
  const [impactSearch, setImpactSearch] = React.useState<string>('');
  const [impactSel, setImpactSel] = React.useState<{ [key: string]: true }>({});
  const [impactLimit, setImpactLimit] = React.useState<number>(PAGE);
  const [notice, setNotice] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);
  const [busy, setBusy] = React.useState<string>('');
  const [saveDialog, setSaveDialog] = React.useState<SaveKind | undefined>(undefined);
  const [leaveDialog, setLeaveDialog] = React.useState<boolean>(false);
  const [taskProgress, setTaskProgress] = React.useState<{ done: number; total: number; failed: number } | undefined>(undefined);
  const [changed, setChanged] = React.useState<boolean>(false);
  const keyOf = (x: IDlpWorking): string => JSON.stringify([x.displayName.trim(), x.draftName.trim(), x.defaultGroup, x.groups, x.scope, x.envs]);
  /** Stand beim Öffnen bzw. beim letzten Speichern (Entwurf/neue Richtlinie) */
  const [savedKey, setSavedKey] = React.useState<string>(() => keyOf(props.working));

  const unblockable = React.useMemo(() => {
    const m: { [id: string]: true } = {};
    props.unblockable.forEach(n => { m[connectorId(n).toLowerCase()] = true; });
    return m;
  }, [props.unblockable]);

  /** alle Connectors: CoE-Inventar plus die, die nur in der Richtlinie stehen */
  const connectors: IDlpConnector[] = React.useMemo(() => {
    const known: { [id: string]: true } = {};
    const list = base.connectors.slice();
    list.forEach(c => { known[c.id.toLowerCase()] = true; });
    Object.keys(w.groups).forEach(id => {
      if (!known[id]) {
        const ref = w.refs[id];
        const name = ref ? ref.name : id.split('/').pop() || id;
        list.push({ id: ref ? ref.id : id, name, displayName: name, tier: '', publisher: '', apps: 0, flows: 0 });
        known[id] = true;
      }
    });
    return list.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }, [base.connectors, w.groups, w.refs]);
  const conLabel = React.useMemo(() => {
    const m: { [name: string]: string } = {};
    connectors.forEach(c => { m[c.name.toLowerCase()] = c.displayName; });
    return (n: string): string => m[n.toLowerCase()] || n;
  }, [connectors]);
  const envLabel = React.useMemo(() => {
    const m: { [name: string]: string } = {};
    base.envs.forEach(e => { m[e.name.toLowerCase().replace(/^default-/, '')] = e.displayName; });
    return (n: string): string => m[n.toLowerCase().replace(/^default-/, '')] || n;
  }, [base.envs]);

  const update = (patch: Partial<IDlpWorking>): void => setW(prev => ({ ...prev, ...patch }));
  const single = w.scope === 'SingleEnvironment';
  const needEnvs = w.scope !== 'AllEnvironments';
  const dirty = isDirty(w, connectors);
  /** Änderungen, die weder im Tenant noch als Entwurf gespeichert sind */
  const unsaved = w.mode === 'edit' && w.policy ? dirty : keyOf(w) !== savedKey;
  const diff = diffConnectors(w, connectors);
  const nameOk = w.displayName.trim().length > 0;
  const envOk = !needEnvs || w.scope === 'ExceptEnvironments' || w.envs.length > 0;
  const stateKey = JSON.stringify([w.defaultGroup, w.groups, w.scope, w.envs]);

  const go = (s: number): void => {
    let target = s;
    if (target === 3 && !needEnvs) {
      target = s > step ? 4 : 2;
    }
    setStep(target);
    setVisited(prev => ({ ...prev, [target]: true }));
  };

  // Auswirkung berechnen, sobald der Schritt geöffnet wird (und neu, wenn sich die Regeln geändert haben)
  const runImpact = (fresh: boolean): void => {
    setImpactLoading(true);
    setImpactError('');
    const key = stateKey;
    loadInventory(props.service, props.maxRecords, fresh).then(inv => {
      setImpact(computeImpact(w, inv, base.envs));
      setImpactKey(key);
      setImpactSel({});
      setImpactLoading(false);
    }, e => { setImpactError(`Inventar aus dem CoE nicht lesbar: ${errText(e)}`); setImpactLoading(false); });
  };
  React.useEffect(() => {
    if ((step === 4 || step === 5) && impactKey !== stateKey && !impactLoading) {
      runImpact(false);
    }
  }, [step, stateKey]);

  // ---------- Connectors ----------
  const q = conSearch.trim().toLowerCase();
  const groupCount = (g: DlpGroup): number => connectors.filter(c => groupOf(w, c.id) === g).length;
  const visibleCons = connectors.filter(c => groupOf(w, c.id) === group && (!q || c.displayName.toLowerCase().indexOf(q) >= 0 || c.name.toLowerCase().indexOf(q) >= 0));
  const selectedCons = visibleCons.filter(c => conSel[c.id.toLowerCase()]);
  const move = (to: DlpGroup): void => {
    const groups = { ...w.groups };
    const refs = { ...w.refs };
    let skipped = 0;
    selectedCons.forEach(c => {
      const id = c.id.toLowerCase();
      if (to === 'Blocked' && unblockable[id]) {
        skipped++;
        return;
      }
      groups[id] = to;
      if (!refs[id]) {
        refs[id] = { id: c.id, name: c.name, type: 'Microsoft.PowerApps/apis' };
      }
    });
    update({ groups, refs });
    setConSel({});
    setNotice(skipped ? { type: MessageBarType.warning, text: `${skipped} Connector(s) lassen sich laut Microsoft nicht blockieren und bleiben in ihrer Gruppe.` } : undefined);
  };
  const conCols: IColumn<IDlpConnector>[] = [
    {
      key: 'sel', label: '', sort: c => (conSel[c.id.toLowerCase()] ? 0 : 1),
      render: c => (
        <input type="checkbox" aria-label={`${c.displayName} auswählen`} checked={!!conSel[c.id.toLowerCase()]}
          onChange={() => setConSel(prev => { const n = { ...prev }; const id = c.id.toLowerCase(); if (n[id]) { delete n[id]; } else { n[id] = true; } return n; })} />
      )
    },
    { key: 'name', label: 'Connector', sort: c => c.displayName.toLowerCase(), render: c => <span className={styles.ulUser}><b>{c.displayName}</b><span className={styles.muted}>{c.name}</span></span> },
    { key: 'tier', label: 'Stufe', sort: c => c.tier, render: c => c.tier || <span className={styles.muted}>–</span> },
    { key: 'pub', label: 'Herausgeber', sort: c => c.publisher.toLowerCase(), render: c => c.publisher || <span className={styles.muted}>–</span> },
    { key: 'apps', label: 'Apps', numeric: true, sort: c => c.apps, render: c => formatNumber(c.apps) },
    { key: 'flows', label: 'Flows', numeric: true, sort: c => c.flows, render: c => formatNumber(c.flows) },
    {
      key: 'hint', label: 'Hinweis', sort: c => (w.base && groupOf(w.base, c.id) !== groupOf(w, c.id) ? 0 : 1),
      render: c => {
        const id = c.id.toLowerCase();
        const was = w.base ? groupOf(w.base, c.id) : undefined;
        return (
          <span className={styles.dlpChips}>
            {was && was !== groupOf(w, c.id) && <Pill tone="warning" title="In dieser Sitzung verschoben">vorher {GROUP_LABEL[was]}</Pill>}
            {!w.groups[id] && <Pill tone="neutral" title="Nicht ausdrücklich zugeordnet – gilt über die Standardgruppe">Standard</Pill>}
            {unblockable[id] && <Pill tone="info" title="Lässt sich laut Microsoft nicht blockieren">nicht blockierbar</Pill>}
          </span>
        );
      }
    }
  ];

  // ---------- Environments ----------
  const inPolicy: { [n: string]: true } = {};
  w.envs.forEach(n => { inPolicy[n.toLowerCase().replace(/^default-/, '')] = true; });
  const envKey = (n: string): string => n.toLowerCase().replace(/^default-/, '');
  const policyEnvs = w.envs.map(n => ({ name: n, displayName: envLabel(n), type: (base.envs.filter(e => envKey(e.name) === envKey(n))[0] || { type: '' }).type }));
  const otherEnvs = base.envs.filter(e => !inPolicy[envKey(e.name)]).map(e => ({ name: e.name, displayName: e.displayName, type: e.type }));
  const filterEnvs = (list: { name: string; displayName: string; type: string }[], s: string): { name: string; displayName: string; type: string }[] => {
    const t = s.trim().toLowerCase();
    return list.filter(e => !t || e.displayName.toLowerCase().indexOf(t) >= 0 || e.name.toLowerCase().indexOf(t) >= 0 || e.type.toLowerCase().indexOf(t) >= 0)
      .sort((a, b) => a.displayName.localeCompare(b.displayName));
  };
  const addEnvs = (): void => {
    const add = otherEnvs.filter(e => envSel[e.name]).map(e => e.name);
    update({ envs: w.envs.concat(add) });
    setEnvSel({});
  };
  const removeEnvs = (): void => {
    update({ envs: w.envs.filter(n => !envSel[n]) });
    setEnvSel({});
  };
  const envList = (title: string, side: 0 | 1, list: { name: string; displayName: string; type: string }[]): JSX.Element => {
    const shown = filterEnvs(list, envSearch[side]);
    return (
      <div className={styles.dlpList} data-dlp-envs={side === 0 ? 'policy' : 'other'}>
        <div className={styles.dlpListHead}>{title} <span className={styles.ccSegmentCount}>{formatNumber(list.length)}</span></div>
        <div style={{ padding: '8px 12px' }}>
          <SearchBox placeholder="Environment suchen" value={envSearch[side]} onChange={(_e, v) => setEnvSearch(side === 0 ? [v || '', envSearch[1]] : [envSearch[0], v || ''])} />
        </div>
        <div className={styles.dlpListBody}>
          {shown.slice(0, 400).map(e => (
            <label key={e.name} className={styles.dlpListRow}>
              <input type="checkbox" checked={!!envSel[e.name]} disabled={single}
                onChange={() => setEnvSel(prev => { const n = { ...prev }; if (n[e.name]) { delete n[e.name]; } else { n[e.name] = true; } return n; })} />
              <span><b>{e.displayName}</b></span>
              <span className={styles.grow} />
              <span className={styles.muted}>{e.type}</span>
            </label>
          ))}
          {!shown.length && <div className={styles.dlpListRow}><span className={styles.muted}>Keine Environments.</span></div>}
          {shown.length > 400 && <div className={styles.dlpListRow}><span className={styles.muted}>{formatNumber(shown.length - 400)} weitere – bitte suchen.</span></div>}
        </div>
      </div>
    );
  };

  // ---------- Auswirkung ----------
  const impactList = impact ? (impactKind === 'app' ? impact.apps : impact.flows) : [];
  const iq = impactSearch.trim().toLowerCase();
  const visibleImpact = impactList.filter(i => (!onlyNew || !i.already)
    && (!iq || i.resource.name.toLowerCase().indexOf(iq) >= 0 || i.resource.envName.toLowerCase().indexOf(iq) >= 0 || i.resource.owner.toLowerCase().indexOf(iq) >= 0));
  const ikey = (i: IDlpImpact): string => `${i.resource.kind}:${i.resource.coeId}`;
  const newCount = (list: IDlpImpact[]): number => list.filter(i => !i.already).length;
  const impactCols: IColumn<IDlpImpact>[] = ([] as IColumn<IDlpImpact>[]).concat(props.canAct ? [{
    key: 'sel', label: '', sort: (i: IDlpImpact) => (impactSel[ikey(i)] ? 0 : 1),
    render: (i: IDlpImpact) => (
      <input type="checkbox" aria-label={`${i.resource.name} auswählen`} checked={!!impactSel[ikey(i)]}
        onChange={() => setImpactSel(prev => { const n = { ...prev }; const k = ikey(i); if (n[k]) { delete n[k]; } else { n[k] = true; } return n; })} />
    )
  }] : []).concat([
    {
      key: 'name', label: impactKind === 'app' ? 'App' : 'Flow', sort: i => i.resource.name.toLowerCase(),
      render: i => <button type="button" className={styles.linkButton} onClick={() => props.onOpen(i.resource.kind === 'app' ? 'admin_app' : 'admin_flow', i.resource.coeId)}>{i.resource.name || '(ohne Namen)'}</button>
    },
    { key: 'env', label: 'Environment', sort: i => i.resource.envName.toLowerCase(), render: i => i.resource.envName || envLabel(i.resource.env) },
    { key: 'owner', label: impactKind === 'app' ? 'Besitzer' : 'Ersteller', sort: i => i.resource.owner.toLowerCase(), render: i => i.resource.owner || <span className={styles.muted}>–</span> },
    {
      key: 'conflict', label: 'Konflikt', title: 'Blau = Business, grau = Non-Business, rot = blockiert', sort: i => i.blocked.length * 100 + i.business.length + i.nonBusiness.length,
      render: i => (
        <span className={styles.dlpChips}>
          {i.blocked.map(n => <span key={`b${n}`} className={`${styles.dlpChip} ${CHIP.Blocked}`} title="blockiert">{conLabel(n)}</span>)}
          {i.business.map(n => <span key={`c${n}`} className={`${styles.dlpChip} ${CHIP.Confidential}`} title="Business">{conLabel(n)}</span>)}
          {i.nonBusiness.map(n => <span key={`g${n}`} className={`${styles.dlpChip} ${CHIP.General}`} title="Non-Business">{conLabel(n)}</span>)}
        </span>
      )
    },
    {
      key: 'state', label: 'Status', sort: i => (i.already ? 1 : 0),
      render: i => (i.already
        ? <Pill tone="neutral" title="Verstößt schon gegen die bisherige Richtlinie">bereits betroffen</Pill>
        : <span className={`${styles.dlpChip} ${styles.dlpChipNew}`}>{w.base ? 'neu betroffen' : 'betroffen'}</span>)
    },
    { key: 'mod', label: 'Geändert', sort: i => i.resource.modified, render: i => (i.resource.modified ? formatDate(new Date(i.resource.modified)) : <span className={styles.muted}>–</span>) }
  ]);
  const exportImpact = (): void => {
    if (!impact) {
      return;
    }
    const rows = impact.apps.concat(impact.flows);
    const lines = ['Typ;Name;Environment;Besitzer/Ersteller;Business;Non-Business;Blockiert;Status;Geändert'].concat(rows.map(i => [
      i.resource.kind === 'app' ? 'App' : 'Flow', i.resource.name, i.resource.envName || envLabel(i.resource.env), i.resource.owner,
      i.business.map(conLabel).join(', '), i.nonBusiness.map(conLabel).join(', '), i.blocked.map(conLabel).join(', '),
      i.already ? 'bereits betroffen' : 'neu betroffen', i.resource.modified ? formatDate(new Date(i.resource.modified)) : ''
    ].map(csvEsc).join(';')));
    downloadText(`DLP-Auswirkung_${w.displayName.replace(/[\\/:*?"<>|]/g, '_')}_${new Date().toISOString().slice(0, 10)}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };
  const toTaskList = async (): Promise<void> => {
    if (!impact) {
      return;
    }
    const all = impact.apps.concat(impact.flows);
    const chosen = all.filter(i => impactSel[ikey(i)]);
    const list = (chosen.length ? chosen : all.filter(i => !onlyNew || !i.already)).slice(0, MAX_TASKS);
    setNotice(undefined);
    let existing: { [k: string]: true } = {};
    try {
      (await loadTasks(props.service)).filter(t => t.policyName === w.displayName.trim()).forEach(t => { existing[t.resourceCoeId] = true; });
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Die CoE-Tabelle „DLP Impact Analysis“ ist nicht lesbar: ${errText(e)}` });
      return;
    }
    const todo = list.filter(i => !existing[i.resource.coeId.toLowerCase()]);
    existing = {};
    setTaskProgress({ done: 0, total: todo.length, failed: 0 });
    let failed = 0;
    let lastError = '';
    for (let n = 0; n < todo.length; n++) {
      try {
        await addTask(props.service, todo[n], w.displayName.trim(), w.policyName || '', conLabel);
      } catch (e) {
        failed++;
        lastError = errText(e);
      }
      setTaskProgress({ done: n + 1, total: todo.length, failed });
    }
    const skipped = list.length - todo.length;
    setNotice({
      type: failed ? MessageBarType.warning : MessageBarType.success,
      text: `${formatNumber(todo.length - failed)} Einträge in die Aufgabenliste übernommen${skipped ? `, ${formatNumber(skipped)} standen schon drin` : ''}.`
        + (failed ? ` ${formatNumber(failed)} Fehler (zuletzt: ${lastError}).` : '') + ' Benachrichtigen kannst du die Maker über die CoE-App „DLP Impact Analysis“.'
    });
    setImpactSel({});
  };

  // ---------- Speichern ----------
  const draft = async (): Promise<void> => {
    setBusy('draft');
    setNotice(undefined);
    try {
      const id = await saveDraft(props.service, w, connectors);
      update({ draftId: id || w.draftId });
      setSavedKey(keyOf(w));
      setChanged(true);
      setNotice({ type: MessageBarType.success, text: `Als Entwurf „${(w.draftName || w.displayName).trim()}“ gespeichert (CoE-Tabelle „Data Policy Drafts“ – auch in der CoE-App sichtbar).` });
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Entwurf nicht gespeichert: ${errText(e)}` });
    }
    setBusy('');
  };

  const summaryText = (): string => {
    const parts: string[] = [];
    if (w.policy && w.displayName.trim() !== w.policy.displayName) {
      parts.push(`Name: „${w.policy.displayName}“ → „${w.displayName.trim()}“`);
    }
    if (w.base && w.defaultGroup !== w.base.defaultGroup) {
      parts.push(`Standardgruppe: ${GROUP_LABEL[w.base.defaultGroup]} → ${GROUP_LABEL[w.defaultGroup]}`);
    }
    if (diff.length) {
      parts.push(`Connectors: ${diff.slice(0, 30).map(d => `${d.connector.displayName} ${GROUP_LABEL[d.from]} → ${GROUP_LABEL[d.to]}`).join('; ')}${diff.length > 30 ? ` (+${diff.length - 30})` : ''}`);
    }
    if (!w.base) {
      parts.push(`Gruppen: ${DLP_GROUPS.map(g => `${GROUP_LABEL[g]} ${groupCount(g)}`).join(', ')}; Standard ${GROUP_LABEL[w.defaultGroup]}`);
    }
    if (!w.base || w.scope !== w.base.scope || envChanges().added.length || envChanges().removed.length) {
      parts.push(`Geltungsbereich: ${SCOPE_LABEL[w.scope]}${needEnvs ? ` (${w.envs.length} Environments)` : ''}`);
    }
    if (impact && impactKey === stateKey) {
      parts.push(`Auswirkung laut CoE: ${impact.apps.length} Apps, ${impact.flows.length} Flows${w.base ? ` (neu: ${newCount(impact.apps)} Apps, ${newCount(impact.flows)} Flows)` : ''}`);
    }
    return parts.join(' | ');
  };

  const envChanges = (): { added: string[]; removed: string[] } => {
    if (!w.base) {
      return { added: [], removed: [] };
    }
    const before = w.base.envs.map(envKey);
    const now = w.envs.map(envKey);
    return { added: w.envs.filter(n => before.indexOf(envKey(n)) < 0), removed: w.base.envs.filter(n => now.indexOf(envKey(n)) < 0) };
  };

  const saveLive = async (kind: SaveKind, reason: string): Promise<void> => {
    if (!props.actions) {
      return;
    }
    setBusy(kind);
    setNotice(undefined);
    const req: IActionRequest = {
      action: kind === 'update' ? 'dlpupdate' : 'dlpcreate', resourceType: 'dlppolicy', resourceId: w.policyName || '', coeRecordId: '',
      resourceName: w.displayName.trim(), environmentId: '', reason: reason.trim()
    };
    let logId: number | undefined;
    try {
      logId = await props.actions.add(req, props.requestedBy, 'In Arbeit', 'Wird direkt im CoE Cockpit ausgeführt (Rechte der angemeldeten Person) …');
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Protokolleintrag nicht möglich – nicht ausgeführt: ${errText(e)}` });
      setBusy('');
      return;
    }
    try {
      let text: string;
      if (kind === 'update') {
        if (!w.policyName || !w.policy) {
          throw new DirectActionError('rejected', 'Die Richtlinie ist nicht live geladen.');
        }
        // Schutz vor gleichzeitigen Änderungen: nur speichern, wenn die Richtlinie seit dem Öffnen unverändert ist
        const current = await props.api.get(w.policyName);
        if (w.policy.lastModifiedTime && current.lastModifiedTime && current.lastModifiedTime !== w.policy.lastModifiedTime) {
          const by = current.lastModifiedBy && current.lastModifiedBy.displayName ? ` von ${current.lastModifiedBy.displayName}` : '';
          throw new DirectActionError('rejected', `Die Richtlinie wurde inzwischen${by} geändert (${formatDate(new Date(current.lastModifiedTime), true)}). Bitte schließen und neu öffnen – deine Änderungen wurden nicht gespeichert.`);
        }
        const saved = await props.api.update(toDefinition(w, connectors, false));
        text = `Richtlinie geändert. ${summaryText()}`;
        update({
          policy: saved && saved.name ? saved : w.policy,
          base: { defaultGroup: w.defaultGroup, groups: { ...w.groups }, scope: w.scope, envs: w.envs.slice() }
        });
      } else {
        const created = await props.api.create(toDefinition(w, connectors, true));
        text = `Richtlinie angelegt${created && created.name ? ` (ID ${created.name})` : ''}. ${summaryText()}`;
        if (created && created.name) {
          update({ mode: 'edit', policyName: created.name, policy: created, base: { defaultGroup: w.defaultGroup, groups: { ...w.groups }, scope: w.scope, envs: w.envs.slice() } });
        }
        setSavedKey(keyOf(w));
      }
      await props.actions.complete(logId, 'Erledigt', text.slice(0, 4000)).catch(() => undefined);
      setChanged(true);
      setNotice({ type: MessageBarType.success, text: kind === 'update' ? 'Die Richtlinie wurde im Tenant geändert und im Aktionsprotokoll vermerkt.' : 'Die Richtlinie wurde im Tenant angelegt und im Aktionsprotokoll vermerkt.' });
    } catch (e) {
      const msg = errText(e);
      await props.actions.complete(logId, e instanceof DirectActionError && e.kind === 'rejected' ? 'Abgelehnt' : 'Fehler', msg.slice(0, 4000)).catch(() => undefined);
      setNotice({ type: MessageBarType.error, text: msg });
    }
    props.onActionsChanged();
    setBusy('');
  };

  const close = (): void => {
    if (unsaved) {
      setLeaveDialog(true);
    } else {
      props.onClose(changed);
    }
  };

  // ---------- Darstellung ----------
  const stepDisabled = (i: number): boolean => (i === 3 && !needEnvs) || (i > 0 && !nameOk);
  const title = w.mode === 'edit' ? 'Richtlinie bearbeiten' : w.mode === 'copy' ? 'Richtlinie kopieren' : w.mode === 'draft' ? 'Entwurf bearbeiten' : 'Neue Richtlinie';
  const impactSummary = impact && impactKey === stateKey
    ? `${formatNumber(impact.apps.length)} Apps und ${formatNumber(impact.flows.length)} Flows betroffen${w.base ? ` – davon neu: ${formatNumber(newCount(impact.apps))} Apps, ${formatNumber(newCount(impact.flows))} Flows` : ''}`
    : impactLoading ? 'wird berechnet …' : 'noch nicht berechnet';
  const env = envChanges();

  const body = (): JSX.Element => {
    switch (step) {
      case 0:
        return (
          <>
            <h3 className={styles.dlpTitle}>Name der Richtlinie</h3>
            <TextField label="Anzeigename" required={true} value={w.displayName} onChange={(_e, v) => update({ displayName: v || '' })} styles={{ root: { maxWidth: 480 } }} />
            {(w.mode !== 'edit') && (
              <TextField label="Name des Entwurfs (nur für „Als Entwurf speichern“)" value={w.draftName} onChange={(_e, v) => update({ draftName: v || '' })} styles={{ root: { maxWidth: 480 } }} />
            )}
            {w.mode === 'copy' && <MessageBar messageBarType={MessageBarType.info}>Kopie: Connectors, Standardgruppe und Environments stammen aus der Vorlage. Gespeichert wird sie als Entwurf oder als neue Richtlinie – die Vorlage bleibt unverändert.</MessageBar>}
            {w.mode === 'edit' && !w.policy && <MessageBar messageBarType={MessageBarType.warning}>Diese Richtlinie stammt aus dem CoE (Stand letzter Sync) – analysieren geht, ändern im Tenant nur, wenn sie live geladen werden kann (Rolle „Power Platform Administrator“ aktiv).</MessageBar>}
          </>
        );
      case 1:
        return (
          <>
            <h3 className={styles.dlpTitle}>Connectors</h3>
            <div className={styles.muted}>Ein Connector gehört genau zu einer Gruppe. Business- und Non-Business-Connectors dürfen nicht in derselben App bzw. demselben Flow vorkommen; blockierte gar nicht.</div>
            <div className={styles.toolbar} style={{ padding: 0 }}>
              <div className={styles.ccSegments} role="group" aria-label="Gruppe">
                {DLP_GROUPS.map(g => (
                  <button key={g} type="button" className={`${styles.ccSegment} ${group === g ? styles.ccSegmentActive : ''}`} aria-pressed={group === g} data-dlp-group={g}
                    onClick={() => { setGroup(g); setConSel({}); setConLimit(PAGE); }}>
                    <Icon iconName={GROUP_ICON[g]} /> {GROUP_LABEL[g]} <span className={styles.ccSegmentCount}>{formatNumber(groupCount(g))}</span>{w.defaultGroup === g ? ' · Standard' : ''}
                  </button>
                ))}
              </div>
              <span className={styles.grow} />
              <Dropdown
                label="Standardgruppe (neue Connectors)"
                selectedKey={w.defaultGroup}
                options={DLP_GROUPS.map(g => ({ key: g, text: GROUP_LABEL[g] }))}
                onChange={(_e, o) => o && update({ defaultGroup: o.key as DlpGroup })}
                styles={{ root: { width: 220 } }}
              />
            </div>
            <div className={styles.toolbar} style={{ padding: 0 }}>
              <SearchBox placeholder="Connector suchen" value={conSearch} onChange={(_e, v) => { setConSearch(v || ''); setConLimit(PAGE); }} styles={{ root: { width: 240 } }} />
              <button type="button" className={styles.textButton} onClick={() => {
                const all = selectedCons.length === visibleCons.length && visibleCons.length > 0;
                const n: { [id: string]: true } = {};
                if (!all) { visibleCons.forEach(c => { n[c.id.toLowerCase()] = true; }); }
                setConSel(n);
              }} disabled={!visibleCons.length}>{selectedCons.length === visibleCons.length && visibleCons.length > 0 ? 'Auswahl aufheben' : `Alle ${formatNumber(visibleCons.length)} auswählen`}</button>
              <span className={styles.grow} />
              {DLP_GROUPS.filter(g => g !== group).map(g => (
                <DefaultButton key={g} text={`Nach ${GROUP_LABEL[g]} (${formatNumber(selectedCons.length)})`} iconProps={{ iconName: GROUP_ICON[g] }} disabled={!selectedCons.length} onClick={() => move(g)} />
              ))}
            </div>
            <SortTable<IDlpConnector>
              columns={conCols}
              rows={visibleCons}
              rowKey={c => c.id}
              initial={{ key: 'name', desc: false }}
              empty={connectors.length ? 'Keine Connectors in dieser Gruppe bzw. für diese Suche.' : 'Keine Connectors im CoE gefunden.'}
              limit={conLimit}
            />
            {visibleCons.length > conLimit && (
              <div className={styles.ccFoot}><span>{formatNumber(conLimit)} von {formatNumber(visibleCons.length)} angezeigt.</span>
                <button type="button" className={styles.textButton} onClick={() => setConLimit(conLimit + PAGE)}>Weitere anzeigen</button></div>
            )}
            <div className={styles.ccFoot}>Connectors aus dem CoE-Inventar (ohne Custom Connectors – die regelt Microsoft über URL-Muster im Admin Center). Aktionen- und Endpunkt-Regeln der Richtlinie bleiben unverändert.</div>
          </>
        );
      case 2:
        return (
          <>
            <h3 className={styles.dlpTitle}>Geltungsbereich</h3>
            {single
              ? <MessageBar messageBarType={MessageBarType.info}>Environment-Richtlinie: gilt nur für „{envLabel(w.envs[0] || '')}“. Der Geltungsbereich lässt sich nicht ändern (für eine andere Auswahl die Richtlinie kopieren).</MessageBar>
              : (
                <ChoiceGroup
                  selectedKey={w.scope}
                  options={(['AllEnvironments', 'ExceptEnvironments', 'OnlyEnvironments'] as DlpScope[]).map(s => ({ key: s, text: SCOPE_LABEL[s] }))}
                  onChange={(_e, o) => o && update({ scope: o.key as DlpScope })}
                />
              )}
          </>
        );
      case 3:
        return (
          <>
            <h3 className={styles.dlpTitle}>Environments</h3>
            <div className={styles.muted}>{w.scope === 'ExceptEnvironments' ? 'Die Richtlinie gilt für alle Environments außer den hier eingetragenen.' : 'Die Richtlinie gilt nur für die hier eingetragenen Environments.'}</div>
            <div className={styles.dlpCols}>
              {envList(w.scope === 'ExceptEnvironments' ? 'Ausgenommen' : 'In der Richtlinie', 0, policyEnvs)}
              {envList('Weitere Environments (CoE)', 1, otherEnvs)}
            </div>
            {!single && (
              <div className={styles.toolbar} style={{ padding: 0 }}>
                <DefaultButton text="Ausgewählte hinzufügen" iconProps={{ iconName: 'Add' }} disabled={!otherEnvs.some(e => envSel[e.name])} onClick={addEnvs} />
                <DefaultButton text="Ausgewählte entfernen" iconProps={{ iconName: 'Remove' }} disabled={!w.envs.some(n => envSel[n])} onClick={removeEnvs} />
              </div>
            )}
            {base.envProblem && <div className={styles.muted}><Icon iconName="Info" /> {base.envProblem}</div>}
          </>
        );
      case 4:
        return (
          <>
            <h3 className={styles.dlpTitle}>Auswirkung</h3>
            <div className={styles.muted}>Apps und Flows laut CoE-Inventar, die mit diesen Regeln einen blockierten Connector nutzen oder Business- und Non-Business-Connectors mischen.</div>
            {impactLoading && <ProgressIndicator label="Inventar wird ausgewertet …" />}
            {impactError && <MessageBar messageBarType={MessageBarType.error}>{impactError}</MessageBar>}
            {impact && (
              <>
                <div className={styles.secGrid}>
                  <Tile label="Betroffene Apps" hint={w.base ? `davon neu: ${formatNumber(newCount(impact.apps))}` : `von ${formatNumber(impact.analyzedApps)} geprüften`} icon="AppIconDefault"
                    tone={impact.apps.length ? styles.ovToneCritical : styles.ovToneGood} value={formatNumber(impact.apps.length)} active={impactKind === 'app'} onClick={() => { setImpactKind('app'); setImpactLimit(PAGE); }} />
                  <Tile label="Betroffene Flows" hint={w.base ? `davon neu: ${formatNumber(newCount(impact.flows))}` : `von ${formatNumber(impact.analyzedFlows)} geprüften`} icon="Flow"
                    tone={impact.flows.length ? styles.ovToneCritical : styles.ovToneGood} value={formatNumber(impact.flows.length)} active={impactKind === 'flow'} onClick={() => { setImpactKind('flow'); setImpactLimit(PAGE); }} />
                  <Tile label="Geprüft" hint={`${formatNumber(impact.analyzedApps)} Apps, ${formatNumber(impact.analyzedFlows)} Flows`} icon="Globe" tone={styles.ovToneInfo}
                    value={`${formatNumber(impact.envCount)} Env.`} />
                </div>
                {impact.truncated && <div className={styles.muted}><Icon iconName="Info" /> Sehr großes Inventar – ausgewertet wurden die ersten {formatNumber(props.maxRecords)} Apps bzw. Flows mit Verbindungen.</div>}
                <div className={styles.toolbar} style={{ padding: 0 }}>
                  <SearchBox placeholder="Name, Environment oder Person" value={impactSearch} onChange={(_e, v) => { setImpactSearch(v || ''); setImpactLimit(PAGE); }} styles={{ root: { width: 260 } }} />
                  {w.base && (
                    <button type="button" className={`${styles.ccSegment} ${onlyNew ? styles.ccSegmentActive : ''}`} aria-pressed={onlyNew} onClick={() => setOnlyNew(!onlyNew)}>Nur neu betroffene</button>
                  )}
                  <span className={styles.resultBadge} role="status"><strong>{formatNumber(visibleImpact.length)}</strong> {impactKind === 'app' ? 'Apps' : 'Flows'}</span>
                  <span className={styles.grow} />
                  <DefaultButton text="Neu berechnen" iconProps={{ iconName: 'Refresh' }} onClick={() => runImpact(true)} disabled={impactLoading} />
                  {props.canAct && (
                    <DefaultButton
                      text={Object.keys(impactSel).length ? `Ausgewählte zur Aufgabenliste (${formatNumber(Object.keys(impactSel).length)})` : 'Alle zur Aufgabenliste'}
                      iconProps={{ iconName: 'TaskManager' }}
                      disabled={!!taskProgress && taskProgress.done < taskProgress.total || !(impact.apps.length + impact.flows.length)}
                      onClick={() => { toTaskList().catch(e => setNotice({ type: MessageBarType.error, text: errText(e) })); }}
                    />
                  )}
                  <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportImpact} disabled={!(impact.apps.length + impact.flows.length)} />
                </div>
                {taskProgress && taskProgress.done < taskProgress.total && (
                  <ProgressIndicator label={`Aufgabenliste: ${formatNumber(taskProgress.done)} von ${formatNumber(taskProgress.total)}`} percentComplete={taskProgress.total ? taskProgress.done / taskProgress.total : 1} />
                )}
                <SortTable<IDlpImpact>
                  columns={impactCols}
                  rows={visibleImpact}
                  rowKey={ikey}
                  initial={{ key: 'state', desc: false }}
                  empty={impactList.length ? 'Keine Einträge für diesen Filter.' : `Keine betroffenen ${impactKind === 'app' ? 'Apps' : 'Flows'} – sehr gut.`}
                  limit={impactLimit}
                />
                {visibleImpact.length > impactLimit && (
                  <div className={styles.ccFoot}><span>{formatNumber(impactLimit)} von {formatNumber(visibleImpact.length)} angezeigt.</span>
                    <button type="button" className={styles.textButton} onClick={() => setImpactLimit(impactLimit + PAGE)}>Weitere anzeigen</button></div>
                )}
                <div className={styles.ccFoot}>Konflikt: <span className={`${styles.dlpChip} ${styles.dlpChipBlocked}`}>blockiert</span> <span className={`${styles.dlpChip} ${styles.dlpChipBusiness}`}>Business</span> <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`}>Non-Business</span> · Grundlage: CoE-Felder „App Connections“ / „Flow Connections“ (letzter CoE-Sync). Custom Connectors zählen zur Standardgruppe. „Zur Aufgabenliste“ trägt die Einträge in die CoE-Tabelle „DLP Impact Analysis“ ein (höchstens {MAX_TASKS} je Durchlauf).</div>
              </>
            )}
          </>
        );
      default:
        return (
          <>
            <h3 className={styles.dlpTitle}>Überprüfen</h3>
            <div className={styles.dlpSummary}>
              <div className={styles.dlpSummaryItem}><span>Richtlinie</span><b>{w.displayName.trim() || '–'}</b></div>
              <div className={styles.dlpSummaryItem}><span>Standardgruppe</span><b>{GROUP_LABEL[w.defaultGroup]}</b></div>
              <div className={styles.dlpSummaryItem}><span>Connectors</span><b>{DLP_GROUPS.map(g => `${GROUP_LABEL[g]} ${formatNumber(groupCount(g))}`).join(' · ')}</b></div>
              <div className={styles.dlpSummaryItem}><span>Geltungsbereich</span><b>{SCOPE_LABEL[w.scope]}{needEnvs ? ` (${formatNumber(w.envs.length)})` : ''}</b></div>
              <div className={styles.dlpSummaryItem}><span>Auswirkung (CoE)</span><b>{impactSummary}</b></div>
            </div>
            {w.base && (
              <section className={styles.dlpList}>
                <div className={styles.dlpListHead}>Änderungen gegenüber der {w.mode === 'edit' ? 'Live-Richtlinie' : 'Vorlage'}</div>
                <div style={{ padding: '8px 12px' }}>
                  {!dirty ? <span className={styles.muted}>Keine Änderungen.</span> : (
                    <ul className={styles.dlpChanges}>
                      {w.policy && w.displayName.trim() !== w.policy.displayName && <li>Name: „{w.policy.displayName}“ → „{w.displayName.trim()}“</li>}
                      {w.defaultGroup !== w.base.defaultGroup && <li>Standardgruppe: {GROUP_LABEL[w.base.defaultGroup]} → {GROUP_LABEL[w.defaultGroup]}</li>}
                      {diff.slice(0, 50).map(d => (
                        <li key={d.connector.id}>{d.connector.displayName}: <span className={`${styles.dlpChip} ${CHIP[d.from]}`}>{GROUP_LABEL[d.from]}</span> → <span className={`${styles.dlpChip} ${CHIP[d.to]}`}>{GROUP_LABEL[d.to]}</span></li>
                      ))}
                      {diff.length > 50 && <li>… und {formatNumber(diff.length - 50)} weitere Connectors</li>}
                      {w.scope !== w.base.scope && <li>Geltungsbereich: {SCOPE_LABEL[w.base.scope]} → {SCOPE_LABEL[w.scope]}</li>}
                      {needEnvs && env.added.length > 0 && <li>Environments hinzugefügt: {env.added.map(envLabel).join(', ')}</li>}
                      {needEnvs && env.removed.length > 0 && <li>Environments entfernt: {env.removed.map(envLabel).join(', ')}</li>}
                    </ul>
                  )}
                </div>
              </section>
            )}
            {!props.canAct && props.actHint && <div className={styles.ccWarnings}><div className={styles.muted}><Icon iconName="Info" /> {props.actHint}</div></div>}
            <div className={styles.toolbar} style={{ padding: 0 }}>
              {props.canAct && w.mode === 'edit' && (
                <PrimaryButton text="Änderungen übernehmen" iconProps={{ iconName: 'LightningBolt' }} disabled={!!busy || !w.policy || !dirty || !nameOk || !envOk} onClick={() => setSaveDialog('update')} />
              )}
              {props.canAct && w.mode !== 'edit' && (
                <PrimaryButton text="Als neue Richtlinie anlegen" iconProps={{ iconName: 'LightningBolt' }} disabled={!!busy || !nameOk || !envOk} onClick={() => setSaveDialog('create')} />
              )}
              {props.canAct && (
                <DefaultButton text={w.draftId ? 'Entwurf aktualisieren' : 'Als Entwurf speichern'} iconProps={{ iconName: 'Save' }} disabled={!!busy || !nameOk} onClick={() => { draft().catch(() => setBusy('')); }} />
              )}
              <DefaultButton text="Im Admin Center öffnen" iconProps={{ iconName: 'OpenInNewTab' }} href={ppacPolicyUrl(w.mode === 'edit' ? w.policyName || '' : '')} target="_blank" />
            </div>
            {busy && <ProgressIndicator label={busy === 'draft' ? 'Entwurf wird gespeichert …' : 'Wird im Tenant gespeichert …'} />}
            <div className={styles.ccFoot}>
              „Änderungen übernehmen“ und „Als neue Richtlinie anlegen“ wirken sofort im Tenant – mit deinen Rechten (Rolle „Power Platform Administrator“) und mit Eintrag im Aktionsprotokoll.
              Entwürfe liegen in der CoE-Tabelle „Data Policy Drafts“ und ändern nichts im Tenant.
            </div>
          </>
        );
    }
  };

  return (
    <div className={styles.ov}>
      <section className={styles.ovCard}>
        <div className={styles.toolbar}>
          <div className={styles.dlpCrumb}>
            <button type="button" className={styles.linkButton} onClick={close}>DLP-Auswirkungsanalyse</button>
            <Icon iconName="ChevronRight" />
            <b>{title}</b>
            <span className={styles.muted}>· {w.displayName}</span>
          </div>
          <span className={styles.grow} />
          {unsaved && <Pill tone="warning">ungespeicherte Änderungen</Pill>}
          <DefaultButton text="Schließen" iconProps={{ iconName: 'Cancel' }} onClick={close} />
        </div>
        {notice && (
          <div style={{ padding: '0 24px 8px' }}>
            <MessageBar messageBarType={notice.type} isMultiline={true} onDismiss={() => setNotice(undefined)}>
              {notice.text}
              {notice.type === MessageBarType.error && /Power Platform Administrator|PIM|Zugriff verweigert/i.test(notice.text) && <> <a className={styles.link} href={PIM_URL} target="_blank" rel="noopener noreferrer">Rolle in PIM aktivieren</a></>}
            </MessageBar>
          </div>
        )}
        <div className={styles.dlpWizard}>
          <ol className={styles.dlpSteps}>
            {STEPS.map((s, i) => (
              <li key={s}>
                <button type="button" className={`${styles.dlpStep} ${step === i ? styles.dlpStepActive : ''} ${visited[i] && step !== i ? styles.dlpStepDone : ''}`}
                  disabled={stepDisabled(i)} aria-current={step === i ? 'step' : undefined} onClick={() => go(i)} data-dlp-step={i}>
                  <span className={styles.dlpStepDot}>{visited[i] && step !== i ? <Icon iconName="CheckMark" /> : i + 1}</span>
                  {s}
                </button>
              </li>
            ))}
          </ol>
          <div className={styles.dlpBody}>
            {body()}
            <div className={styles.dlpFooter}>
              <DefaultButton text="Zurück" disabled={step === 0} onClick={() => go(step - 1)} />
              {step < STEPS.length - 1 && <PrimaryButton text="Weiter" disabled={!nameOk || (step === 3 && !envOk)} onClick={() => go(step + 1)} />}
              {step === 3 && !envOk && <span className={styles.muted}>Bitte mindestens ein Environment eintragen.</span>}
            </div>
          </div>
        </div>
      </section>
      {saveDialog && (
        <SaveDialog
          kind={saveDialog}
          name={w.displayName.trim()}
          impact={impactSummary}
          onCancel={() => setSaveDialog(undefined)}
          onConfirm={reason => { const k = saveDialog; setSaveDialog(undefined); saveLive(k, reason).catch(() => setBusy('')); }}
        />
      )}
      {leaveDialog && (
        <Dialog hidden={false} onDismiss={() => setLeaveDialog(false)} dialogContentProps={{ type: DialogType.normal, title: 'Änderungen verwerfen?', subText: 'Die Änderungen sind weder im Tenant noch als Entwurf gespeichert.' }}>
          <DialogFooter>
            <PrimaryButton text="Verwerfen" onClick={() => props.onClose(changed)} />
            <DefaultButton text="Weiter bearbeiten" onClick={() => setLeaveDialog(false)} />
          </DialogFooter>
        </Dialog>
      )}
    </div>
  );
};

/** Bestätigung mit Begründung (steht im Aktionsprotokoll). */
const SaveDialog: React.FC<{ kind: SaveKind; name: string; impact: string; onCancel: () => void; onConfirm: (reason: string) => void }> = props => {
  const [reason, setReason] = React.useState<string>('');
  return (
    <Dialog
      hidden={false}
      onDismiss={props.onCancel}
      minWidth={520}
      dialogContentProps={{ type: DialogType.normal, title: props.kind === 'update' ? 'Richtlinie im Tenant ändern' : 'Neue Richtlinie im Tenant anlegen', subText: `„${props.name}“` }}
    >
      <div className={styles.notify}>
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
          {props.kind === 'update'
            ? 'Die Änderung wirkt sofort. Apps und Flows, die danach gegen die Richtlinie verstoßen, werden blockiert bzw. ausgesetzt.'
            : 'Die Richtlinie gilt sofort für die gewählten Environments. Apps und Flows, die gegen sie verstoßen, werden blockiert bzw. ausgesetzt.'}
          {' '}Auswirkung laut CoE: {props.impact}.
        </MessageBar>
        <TextField label="Begründung (für das Aktionsprotokoll)" required={true} multiline={true} rows={2} value={reason} onChange={(_e, v) => setReason(v || '')} />
      </div>
      <DialogFooter>
        <PrimaryButton text={props.kind === 'update' ? 'Ändern' : 'Anlegen'} iconProps={{ iconName: 'LightningBolt' }} disabled={reason.trim().length < 5} onClick={() => props.onConfirm(reason)} />
        <DefaultButton text="Abbrechen" onClick={props.onCancel} />
      </DialogFooter>
    </Dialog>
  );
};

