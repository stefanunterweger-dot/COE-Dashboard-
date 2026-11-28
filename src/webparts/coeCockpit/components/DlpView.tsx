import * as React from 'react';
import { AadHttpClientFactory } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable } from './CopilotCreditsView';
import { DlpEditor } from './DlpEditor';
import {
  appliesTo, DlpApi, DlpGroup, emptyWorking, groupOf, GROUP_LABEL, IDlpBase, IDlpListItem, IDlpTask, IDlpWorking, itemFromLive, loadCoePolicies, loadDlpBase,
  loadDrafts, loadTasks, PPAC_DLP_URL, ppacPolicyUrl, resetDlpCache, scopeText, unblockableMap, workingFrom
} from '../models/dlp';
import { IActionRequest } from '../models/actions';
import { ActionService } from '../services/ActionService';
import { DataverseService } from '../services/DataverseService';
import { DirectActionError } from '../services/DirectActionService';
import { PIM_URL } from '../services/PowerPlatformApiService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * „DLP-Auswirkungsanalyse“ (Governance) – entspricht der CoE-App „Data Policy Impact Analysis“:
 * angewendete Richtlinien (live), Entwürfe und Aufgabenliste (CoE); Bearbeiten, Kopieren, Löschen, im Admin Center öffnen.
 */
export interface IDlpViewProps {
  service: DataverseService;
  factory: AadHttpClientFactory;
  /** Obergrenze je Typ für die Auswertung des Inventars */
  maxRecords: number;
  reloadKey: string;
  canAct: boolean;
  actHint: string;
  actions: ActionService | undefined;
  requestedBy: string;
  onOpen: (logicalName: string, id: string) => void;
  onActionsChanged: () => void;
}

type Tab = 'live' | 'draft' | 'tasks';
const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));
const LOGWARN = ' Hinweis: Der Eintrag im Aktionsprotokoll konnte nicht abgeschlossen werden (steht noch auf „In Arbeit“) – bitte in der Liste „CoE Cockpit Aktionen“ nachtragen.';
/** CSV-Feld: Trennzeichen/Anführungszeichen maskieren; Formeln (=, +, -, @ am Anfang – Namen kommen von Makern) mit ' entschärfen. */
const csvEsc = (raw: string): string => {
  const v = /^[=+\-@\t\r]/.test(raw) ? `'${raw}` : raw;
  return /[;"\n\r]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
};
const dateCell = (v: string): React.ReactNode => (v ? formatDate(new Date(v)) : <span className={styles.muted}>–</span>);

export const DlpView: React.FC<IDlpViewProps> = props => {
  const api = React.useMemo(() => new DlpApi(props.factory), [props.factory]);
  const [tab, setTab] = React.useState<Tab>('live');
  const [base, setBase] = React.useState<IDlpBase | undefined>(undefined);
  const [items, setItems] = React.useState<IDlpListItem[] | undefined>(undefined);
  const [source, setSource] = React.useState<'live' | 'coe'>('live');
  const [liveError, setLiveError] = React.useState<{ text: string; kind: string } | undefined>(undefined);
  const [drafts, setDrafts] = React.useState<IDlpListItem[] | undefined>(undefined);
  const [draftError, setDraftError] = React.useState<string>('');
  const [tasks, setTasks] = React.useState<IDlpTask[] | undefined>(undefined);
  const [taskError, setTaskError] = React.useState<string>('');
  const [unblockable, setUnblockable] = React.useState<string[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [loadedAt, setLoadedAt] = React.useState<Date | undefined>(undefined);
  const [selected, setSelected] = React.useState<string>('');
  const [search, setSearch] = React.useState<string>('');
  const [envFilter, setEnvFilter] = React.useState<string>('');
  const [taskPolicy, setTaskPolicy] = React.useState<string>('');
  const [editing, setEditing] = React.useState<IDlpWorking | undefined>(undefined);
  const [opening, setOpening] = React.useState<boolean>(false);
  const [notice, setNotice] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);
  const [deleting, setDeleting] = React.useState<IDlpListItem | undefined>(undefined);
  const gen = React.useRef<number>(0);

  const load = (fresh: boolean): void => {
    const my = ++gen.current;
    const alive = (): boolean => gen.current === my;
    setLoading(true);
    if (fresh) {
      resetDlpCache();
    }
    api.unblockable().then(r => { if (alive()) { setUnblockable(r.names); } }, () => undefined);
    loadDrafts(props.service).then(d => { if (alive()) { setDrafts(d); setDraftError(''); } },
      e => { if (alive()) { setDrafts([]); setDraftError(`Entwürfe nicht lesbar (CoE-Tabelle „Data Policy Drafts“): ${errText(e)}`); } });
    loadTasks(props.service).then(t => { if (alive()) { setTasks(t); setTaskError(''); } },
      e => { if (alive()) { setTasks([]); setTaskError(`Aufgabenliste nicht lesbar (CoE-Tabelle „DLP Impact Analysis“): ${errText(e)}`); } });
    loadDlpBase(props.service, fresh).then(async b => {
      if (!alive()) {
        return;
      }
      setBase(b);
      try {
        const list = await api.list();
        if (!alive()) {
          return;
        }
        const mapped = list.map(itemFromLive);
        setItems(mapped);
        setSource('live');
        setLiveError(undefined);
        setLoading(false);
        setLoadedAt(new Date());
        // Connector-Gruppen je Richtlinie nachladen (die Liste enthält sie nicht immer)
        const todo = mapped.filter(i => !i.groups);
        const next = async (): Promise<void> => {
          const it = todo.shift();
          if (!it || !alive()) {
            return;
          }
          try {
            const full = itemFromLive(await api.get(it.id));
            if (alive()) {
              setItems(prev => (prev || []).map(p => (p.key === it.key ? full : p)));
            }
          } catch {
            // Details bleiben offen – beim Bearbeiten wird ohnehin neu geladen
          }
          await next();
        };
        await Promise.all([next(), next(), next(), next()]);
      } catch (e) {
        if (!alive()) {
          return;
        }
        const kind = e instanceof DirectActionError ? e.kind : 'failed';
        setLiveError({ text: errText(e), kind });
        setSource('coe');
        try {
          const coe = await loadCoePolicies(props.service, b.envs);
          if (alive()) {
            setItems(coe);
            setSource('coe');
          }
        } catch (e2) {
          if (alive()) {
            setItems([]);
            setLiveError({ text: `${errText(e)} – auch die CoE-Tabelle „DLP Policy“ ist nicht lesbar: ${errText(e2)}`, kind });
          }
        }
        if (alive()) {
          setLoading(false);
          setLoadedAt(new Date());
        }
      }
    }, e => {
      if (alive()) {
        setBase({ connectors: [], envs: [], envProblem: errText(e) });
        setItems([]);
        setLoading(false);
      }
    });
  };
  React.useEffect(() => { load(false); }, [props.service, props.reloadKey]);

  const envLabel = React.useMemo(() => {
    const m: { [n: string]: string } = {};
    (base ? base.envs : []).forEach(e => { m[e.name.toLowerCase().replace(/^default-/, '')] = e.displayName; });
    return (n: string): string => m[n.toLowerCase().replace(/^default-/, '')] || n;
  }, [base]);

  const unblockableIds = React.useMemo(() => unblockableMap(unblockable), [unblockable]);
  const counts = (it: IDlpListItem): { [g in DlpGroup]: number } | undefined => {
    if (!it.groups || !base) {
      return undefined;
    }
    const c = { Confidential: 0, General: 0, Blocked: 0 };
    const seen: { [id: string]: true } = {};
    const rules = { defaultGroup: it.defaultGroup || 'General' as DlpGroup, groups: it.groups };
    base.connectors.forEach(con => {
      const id = con.id.toLowerCase();
      seen[id] = true;
      // nicht blockierbare Connectors zählen bei Standardgruppe „Blockiert“ als Non-Business (wie bei Microsoft)
      c[groupOf(rules, id, unblockableIds)]++;
    });
    Object.keys(it.groups).forEach(id => { if (!seen[id]) { c[it.groups![id]]++; } });
    return c;
  };

  const list = tab === 'draft' ? (drafts || []) : (items || []);
  const q = search.trim().toLowerCase();
  const visible = list.filter(i => (!q || i.displayName.toLowerCase().indexOf(q) >= 0 || (i.draftName || '').toLowerCase().indexOf(q) >= 0
    || i.createdBy.toLowerCase().indexOf(q) >= 0 || i.modifiedBy.toLowerCase().indexOf(q) >= 0)
    // nicht lesbare Richtlinien (Environments unbekannt) beim Environment-Filter immer zeigen – mit Hinweis
    && (!envFilter || !!i.loadError || appliesTo(i.scope, i.envs, envFilter)));
  const current = list.filter(i => i.key === selected)[0];

  const open = async (it: IDlpListItem, mode: IDlpWorking['mode']): Promise<void> => {
    setNotice(undefined);
    if (it.loadError) {
      setNotice({ type: MessageBarType.error, text: `„${it.displayName}“ lässt sich nicht analysieren: ${it.loadError} Bitte „Aktualisieren“ oder die Rolle „Power Platform Administrator“ aktivieren (Live-Ansicht).` });
      return;
    }
    if (it.source !== 'live') {
      setEditing(workingFrom(it, mode));
      return;
    }
    setOpening(true);
    try {
      // immer frisch laden: gespeichert wird die vollständige, aktuelle Richtlinie
      const fresh = itemFromLive(await api.get(it.id));
      setEditing(workingFrom(fresh, mode));
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Richtlinie nicht ladbar: ${errText(e)}` });
    }
    setOpening(false);
  };

  // ---------- Tabellen ----------
  const groupCell = (it: IDlpListItem): React.ReactNode => {
    if (it.loadError) {
      return <span className={styles.muted} title={it.loadError}><Icon iconName="Warning" /> nicht lesbar</span>;
    }
    const c = counts(it);
    if (!c) {
      return <span className={styles.muted}>…</span>;
    }
    return (
      <span className={styles.dlpChips}>
        <span className={`${styles.dlpChip} ${styles.dlpChipBusiness}`} title="Business">{formatNumber(c.Confidential)}</span>
        <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`} title="Non-Business">{formatNumber(c.General)}</span>
        <span className={`${styles.dlpChip} ${styles.dlpChipBlocked}`} title="Blockiert">{formatNumber(c.Blocked)}</span>
      </span>
    );
  };
  const policyCols: IColumn<IDlpListItem>[] = [
    {
      key: 'name', label: 'Name', sort: i => i.displayName.toLowerCase(),
      render: i => (
        <span className={styles.ulUser}>
          {props.canAct
            ? <button type="button" className={styles.linkButton} onClick={e => { e.stopPropagation(); setSelected(i.key); open(i, i.source === 'draft' ? 'draft' : 'edit').catch(() => setOpening(false)); }}>{i.displayName || '(ohne Namen)'}</button>
            : <b>{i.displayName || '(ohne Namen)'}</b>}
          {i.source === 'draft' && i.draftName && i.draftName !== i.displayName && <span className={styles.muted}>Entwurf: {i.draftName}</span>}
        </span>
      )
    },
    { key: 'scope', label: 'Gilt für', sort: i => (i.loadError ? '' : scopeText(i, envLabel).toLowerCase()),
      render: i => (i.loadError ? <span className={styles.muted} title={i.loadError}>unbekannt</span> : <span title={i.envs.map(envLabel).join(', ')}>{scopeText(i, envLabel)}</span>) },
    { key: 'default', label: 'Standardgruppe', sort: i => i.defaultGroup || '', render: i => (i.defaultGroup && i.source !== 'coe' ? GROUP_LABEL[i.defaultGroup] : <span className={styles.muted}>–</span>) },
    { key: 'groups', label: 'Business · Non-Business · Blockiert', sort: i => { const c = counts(i); return c ? c.Blocked : -1; }, render: groupCell },
    { key: 'createdBy', label: 'Erstellt von', sort: i => i.createdBy.toLowerCase(), render: i => i.createdBy || <span className={styles.muted}>–</span> },
    { key: 'created', label: 'Erstellt', sort: i => i.created, render: i => dateCell(i.created) },
    { key: 'modifiedBy', label: 'Geändert von', sort: i => i.modifiedBy.toLowerCase(), render: i => i.modifiedBy || <span className={styles.muted}>–</span> },
    { key: 'modified', label: 'Geändert', sort: i => i.modified, render: i => dateCell(i.modified) }
  ];

  const taskList = tasks || [];
  const tq = search.trim().toLowerCase();
  const visibleTasks = taskList.filter(t => (!taskPolicy || t.policyName === taskPolicy)
    && (!tq || t.name.toLowerCase().indexOf(tq) >= 0 || t.env.toLowerCase().indexOf(tq) >= 0 || t.owner.toLowerCase().indexOf(tq) >= 0 || t.policyName.toLowerCase().indexOf(tq) >= 0));
  const taskPolicies = taskList.map(t => t.policyName).filter((p, i, a) => !!p && a.indexOf(p) === i).sort();
  const conflict = (t: IDlpTask): React.ReactNode => (
    <span className={styles.dlpChips}>
      {t.business && <span className={`${styles.dlpChip} ${styles.dlpChipBusiness}`} title="Business">{t.business}</span>}
      {t.nonBusiness && <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`} title="Non-Business">{t.nonBusiness}</span>}
      {t.blocked && <span className={`${styles.dlpChip} ${styles.dlpChipBlocked}`} title="Blockiert">{t.blocked}</span>}
    </span>
  );
  const taskCols: IColumn<IDlpTask>[] = [
    {
      key: 'name', label: 'App / Flow', sort: t => t.name.toLowerCase(),
      render: t => (
        <span className={styles.ulSku}>
          <Icon iconName={t.kind === 'Flow' ? 'Flow' : 'AppIconDefault'} />
          <a className={styles.link} href={props.service.recordUrl('admin_dlpimpactanalysis', t.id)} target="_blank" rel="noopener noreferrer">{t.name || '(ohne Namen)'}</a>
        </span>
      )
    },
    { key: 'policy', label: 'Richtlinie', sort: t => t.policyName.toLowerCase(), render: t => t.policyName || <span className={styles.muted}>–</span> },
    { key: 'env', label: 'Environment', sort: t => t.env.toLowerCase(), render: t => t.env || <span className={styles.muted}>–</span> },
    { key: 'owner', label: 'Verantwortlich', sort: t => t.owner.toLowerCase(), render: t => t.owner || <span className={styles.muted}>–</span> },
    { key: 'conflict', label: 'Konflikt', sort: t => `${t.blocked}${t.business}`, render: conflict },
    { key: 'decision', label: 'Entscheidung', sort: t => t.decision, render: t => t.decision || <span className={styles.muted}>offen</span> },
    { key: 'status', label: 'Status', sort: t => t.status, render: t => (t.status ? <Pill tone="info">{t.status}</Pill> : <span className={styles.muted}>–</span>) },
    { key: 'actionBy', label: 'Handlung bis', sort: t => t.actionBy, render: t => dateCell(t.actionBy) },
    { key: 'created', label: 'Erstellt', sort: t => t.created, render: t => dateCell(t.created) }
  ];
  const exportTasks = (): void => {
    const lines = ['Typ;Name;Richtlinie;Environment;Verantwortlich;Business;Non-Business;Blockiert;Entscheidung;Status;Handlung bis;Erstellt'].concat(visibleTasks.map(t => [
      t.kind, t.name, t.policyName, t.env, t.owner, t.business, t.nonBusiness, t.blocked, t.decision, t.status,
      t.actionBy ? formatDate(new Date(t.actionBy)) : '', t.created ? formatDate(new Date(t.created)) : ''
    ].map(csvEsc).join(';')));
    downloadText(`DLP-Aufgabenliste_${new Date().toISOString().slice(0, 10)}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };

  if (editing && base && props.canAct) {
    return (
      <DlpEditor
        working={editing}
        base={base}
        unblockable={unblockable}
        api={api}
        service={props.service}
        maxRecords={props.maxRecords}
        canAct={props.canAct}
        actHint={props.actHint}
        actions={props.actions}
        requestedBy={props.requestedBy}
        onOpen={props.onOpen}
        onActionsChanged={props.onActionsChanged}
        onClose={changed => { setEditing(undefined); if (changed) { load(false); } }}
      />
    );
  }

  const liveCount = (items || []).length;
  const tabs: { key: Tab; label: string; count: number | undefined }[] = [
    { key: 'live', label: source === 'coe' ? 'Richtlinien (CoE)' : 'Angewendete Richtlinien', count: items ? liveCount : undefined },
    { key: 'draft', label: 'Entwürfe', count: drafts ? drafts.length : undefined },
    { key: 'tasks', label: 'Aufgabenliste', count: tasks ? tasks.length : undefined }
  ];

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="ShieldAlert" /> Richtlinien live · Auswirkung aus dem CoE</div>
          <div className={styles.secHeroTitle}>DLP-Auswirkungsanalyse</div>
          <div className={styles.ovHeroLabel}>
            Welche Apps und Flows wären betroffen, wenn du eine DLP-Richtlinie änderst? Richtlinie wählen, Connector-Gruppen und Environments anpassen,
            Auswirkung prüfen – dann als Entwurf speichern oder direkt übernehmen. Wie die CoE-App „Data Policy Impact Analysis“.
          </div>
          <div className={styles.ovHeroChips}>
            <button type="button" className={styles.ovHeroChip} onClick={() => load(true)} disabled={loading}><Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}</button>
            <a className={styles.ovHeroChip} href={PPAC_DLP_URL} target="_blank" rel="noopener noreferrer"><Icon iconName="OpenInNewTab" /> Datenrichtlinien im Admin Center</a>
            {loadedAt && <span className={styles.ccStamp}>Stand: {formatDate(loadedAt, true)}</span>}
          </div>
        </div>
      </section>

      {liveError && (
        <div className={styles.ccWarnings}>
          <div className={styles.muted}>
            <Icon iconName="Info" /> Richtlinien live nicht lesbar – {source === 'coe' ? 'angezeigt wird der Stand aus dem CoE (letzter Sync, ohne Standardgruppe; ändern im Tenant nicht möglich). ' : ''}{liveError.text}
            {liveError.kind === 'forbidden' && <> <a className={styles.link} href={PIM_URL} target="_blank" rel="noopener noreferrer">Rolle in PIM aktivieren</a></>}
          </div>
        </div>
      )}
      {tab === 'live' && (items || []).some(i => !!i.loadError) && (
        <div className={styles.ccWarnings}>
          <div className={styles.muted}>
            <Icon iconName="Warning" /> {formatNumber((items || []).filter(i => !!i.loadError).length)} Richtlinie(n) aus dem CoE nicht vollständig lesbar – Connector-Gruppen bzw. Environments unbekannt, Auswirkung dafür nicht berechenbar. Bitte „Aktualisieren“.
          </div>
        </div>
      )}
      {notice && <MessageBar messageBarType={notice.type} isMultiline={true} onDismiss={() => setNotice(undefined)}>{notice.text}</MessageBar>}

      <section className={styles.ovCard}>
        <div className={styles.toolbar}>
          <div className={styles.ccSegments} role="tablist" aria-label="Ansicht">
            {tabs.map(t => (
              <button key={t.key} type="button" role="tab" aria-selected={tab === t.key} data-dlp-tab={t.key}
                className={`${styles.ccSegment} ${tab === t.key ? styles.ccSegmentActive : ''}`} onClick={() => { setTab(t.key); setSelected(''); }}>
                {t.label} {t.count !== undefined && <span className={styles.ccSegmentCount}>{formatNumber(t.count)}</span>}
              </button>
            ))}
          </div>
          <span className={styles.grow} />
          <SearchBox placeholder={tab === 'tasks' ? 'App, Flow, Environment oder Person' : 'Richtlinie oder Person'} value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 240 } }} />
          {tab !== 'tasks' && base && (
            <Dropdown
              placeholder="Alle Environments"
              ariaLabel="Richtlinien für Environment"
              selectedKey={envFilter || '__all'}
              options={[{ key: '__all', text: 'Alle Environments' }].concat(base.envs.map(e => ({ key: e.name, text: e.displayName })))}
              onChange={(_e, o) => setEnvFilter(o && o.key !== '__all' ? String(o.key) : '')}
              styles={{ root: { width: 240 }, dropdownItemsWrapper: { maxHeight: 360 } }}
            />
          )}
          {tab === 'tasks' && (
            <Dropdown
              ariaLabel="Richtlinie"
              selectedKey={taskPolicy || '__all'}
              options={[{ key: '__all', text: 'Alle Richtlinien' }].concat(taskPolicies.map(p => ({ key: p, text: p })))}
              onChange={(_e, o) => setTaskPolicy(o && o.key !== '__all' ? String(o.key) : '')}
              styles={{ root: { width: 220 } }}
            />
          )}
        </div>

        {tab !== 'tasks' && (
          <div className={styles.dlpCommands} role="toolbar" aria-label="Richtlinie">
            {props.canAct && (
              <>
                <button type="button" className={styles.dlpCommand} disabled={!current || opening} onClick={() => current && open(current, current.source === 'draft' ? 'draft' : 'edit').catch(() => setOpening(false))}>
                  <Icon iconName="Edit" /> Bearbeiten / Analysieren
                </button>
                <button type="button" className={styles.dlpCommand} disabled={!current || opening} onClick={() => current && open(current, 'copy').catch(() => setOpening(false))}>
                  <Icon iconName="Copy" /> Kopieren
                </button>
              </>
            )}
            <button type="button" className={styles.dlpCommand} disabled={!current || current.source === 'draft'} onClick={() => current && window.open(ppacPolicyUrl(current.id), '_blank', 'noopener')}>
              <Icon iconName="OpenInNewTab" /> Im Admin Center öffnen
            </button>
            {props.canAct && (
              <button type="button" className={`${styles.dlpCommand} ${styles.dlpCommandDanger}`} disabled={!current || current.source === 'coe'} onClick={() => current && setDeleting(current)}>
                <Icon iconName="Delete" /> Löschen
              </button>
            )}
            <span className={styles.grow} />
            {props.canAct && (
              <button type="button" className={styles.dlpCommand} disabled={!base} onClick={() => setEditing(emptyWorking())}>
                <Icon iconName="Add" /> Neue Richtlinie
              </button>
            )}
          </div>
        )}
        {!props.canAct && props.actHint && tab !== 'tasks' && <div className={styles.ccWarnings}><div className={styles.muted}><Icon iconName="Info" /> {props.actHint}</div></div>}
        {tab === 'draft' && draftError && <div className={styles.ccWarnings}><div className={styles.muted}><Icon iconName="Info" /> {draftError}</div></div>}
        {tab === 'tasks' && taskError && <div className={styles.ccWarnings}><div className={styles.muted}><Icon iconName="Info" /> {taskError}</div></div>}

        {loading && !items && <div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>Richtlinien werden gelesen …</span></div>}
        {tab !== 'tasks' && items && (
          <SortTable<IDlpListItem>
            columns={policyCols}
            rows={visible}
            rowKey={i => i.key}
            initial={{ key: tab === 'draft' ? 'modified' : 'name', desc: tab === 'draft' }}
            empty={list.length ? 'Keine Richtlinien für diesen Filter.' : tab === 'draft' ? 'Keine Entwürfe.' : 'Keine DLP-Richtlinien.'}
            onRow={i => setSelected(selected === i.key ? '' : i.key)}
            rowClass={i => (i.key === selected ? styles.ulRowSelected : undefined)}
          />
        )}
        {tab === 'tasks' && tasks && (
          <>
            <div className={styles.toolbar} style={{ paddingTop: 0 }}>
              <span className={styles.resultBadge} role="status"><strong>{formatNumber(visibleTasks.length)}</strong> Einträge</span>
              <span className={styles.grow} />
              <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportTasks} disabled={!visibleTasks.length} />
            </div>
            <SortTable<IDlpTask>
              columns={taskCols}
              rows={visibleTasks}
              rowKey={t => t.id}
              initial={{ key: 'created', desc: true }}
              empty={taskList.length ? 'Keine Einträge für diesen Filter.' : 'Die Aufgabenliste ist leer. Betroffene Apps und Flows kommen über „Auswirkung → Zur Aufgabenliste“ hinein.'}
              limit={500}
            />
          </>
        )}
        <div className={styles.ccFoot}>
          {tab === 'tasks'
            ? 'CoE-Tabelle „DLP Impact Analysis“ – Entscheidung, Status und Benachrichtigung der Maker pflegst du in der CoE-App (Name anklicken).'
            : tab === 'draft'
              ? 'Entwürfe aus der CoE-Tabelle „Data Policy Drafts“ – dieselben wie in der CoE-App. Sie ändern nichts im Tenant.'
              : source === 'live'
                ? `Live aus dem Tenant (Admin-Schnittstelle, deine Rechte). Zeile anklicken = auswählen${props.canAct ? ', Name anklicken = öffnen' : ''}.`
                : 'Stand aus dem CoE (letzter Sync). Für die Live-Ansicht die Rolle „Power Platform Administrator“ aktivieren.'}
        </div>
      </section>

      {deleting && (
        <DeleteDialog
          item={deleting}
          onCancel={() => setDeleting(undefined)}
          onConfirm={async reason => {
            const it = deleting;
            setDeleting(undefined);
            setNotice(undefined);
            if (it.source === 'draft') {
              try {
                await props.service.remove('admin_datapolicydraftses', it.id);
                setNotice({ type: MessageBarType.success, text: `Entwurf „${it.draftName || it.displayName}“ gelöscht.` });
                setSelected('');
                load(false);
              } catch (e) {
                setNotice({ type: MessageBarType.error, text: `Entwurf nicht gelöscht: ${errText(e)}` });
              }
              return;
            }
            if (!props.actions) {
              return;
            }
            const req: IActionRequest = { action: 'dlpdelete', resourceType: 'dlppolicy', resourceId: it.id, coeRecordId: '', resourceName: it.displayName, environmentId: '', reason: reason.trim() };
            let logId: number;
            try {
              logId = await props.actions.add(req, props.requestedBy, 'In Arbeit', 'Wird direkt im CoE Cockpit ausgeführt (Rechte der angemeldeten Person) …');
            } catch (e) {
              setNotice({ type: MessageBarType.error, text: `Protokolleintrag nicht möglich – nicht gelöscht: ${errText(e)}` });
              return;
            }
            try {
              await api.remove(it.id);
              const logged = await props.actions.complete(logId, 'Erledigt', `Richtlinie „${it.displayName}“ (${scopeText(it, envLabel)}) im Tenant gelöscht.`).then(() => true, () => false);
              setNotice({
                type: logged ? MessageBarType.success : MessageBarType.warning,
                text: `Die Richtlinie „${it.displayName}“ wurde im Tenant gelöscht${logged ? ' und im Aktionsprotokoll vermerkt.' : `.${LOGWARN}`}`
              });
              setSelected('');
              load(false);
            } catch (e) {
              const logged = await props.actions.complete(logId, e instanceof DirectActionError && e.kind === 'rejected' ? 'Abgelehnt' : 'Fehler', errText(e).slice(0, 4000)).then(() => true, () => false);
              setNotice({ type: MessageBarType.error, text: logged ? errText(e) : `${errText(e)}${LOGWARN}` });
            }
            props.onActionsChanged();
          }}
        />
      )}
    </div>
  );
};

const DeleteDialog: React.FC<{ item: IDlpListItem; onCancel: () => void; onConfirm: (reason: string) => Promise<void> }> = props => {
  const [reason, setReason] = React.useState<string>('');
  const [typed, setTyped] = React.useState<string>('');
  const draft = props.item.source === 'draft';
  return (
    <Dialog
      hidden={false}
      onDismiss={props.onCancel}
      minWidth={520}
      dialogContentProps={{ type: DialogType.normal, title: draft ? 'Entwurf löschen' : 'DLP-Richtlinie im Tenant löschen', subText: `„${draft ? props.item.draftName || props.item.displayName : props.item.displayName}“` }}
    >
      <div className={styles.notify}>
        {draft
          ? <MessageBar messageBarType={MessageBarType.info}>Gelöscht wird nur der Entwurf in der CoE-Tabelle – im Tenant ändert sich nichts.</MessageBar>
          : (
            <>
              <MessageBar messageBarType={MessageBarType.severeWarning} isMultiline={true}>
                Die Richtlinie wird sofort im Tenant gelöscht; ihre Regeln gelten danach in keinem Environment mehr. Das lässt sich nicht rückgängig machen.
                Tipp: vorher kopieren und als Entwurf speichern.
              </MessageBar>
              <TextField label="Begründung (für das Aktionsprotokoll)" required={true} multiline={true} rows={2} value={reason} onChange={(_e, v) => setReason(v || '')} />
              <TextField label="Zur Bestätigung den Namen der Richtlinie eingeben" value={typed} onChange={(_e, v) => setTyped(v || '')} />
            </>
          )}
      </div>
      <DialogFooter>
        <PrimaryButton text="Löschen" iconProps={{ iconName: 'Delete' }} disabled={!draft && (reason.trim().length < 5 || typed.trim() !== props.item.displayName.trim())}
          onClick={() => { props.onConfirm(reason).catch(() => undefined); }} />
        <DefaultButton text="Abbrechen" onClick={props.onCancel} />
      </DialogFooter>
    </Dialog>
  );
};
