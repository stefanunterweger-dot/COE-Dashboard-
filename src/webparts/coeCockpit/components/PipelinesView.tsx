import * as React from 'react';
import { AadHttpClientFactory } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable, Tile } from './CopilotCreditsView';
import {
  IHostData, IHostScan, IPipeline, IPipelineHost, IPlEnv, IPlRun, ICoeEnv, loadCoeEnvs, loadHost, mergeHosts, PPAC_DEPLOYMENT_URL, readHostScan,
  resetPipelineCache, RunState, scanHosts
} from '../models/pipelineHosts';
import { DataverseService } from '../services/DataverseService';
import { PIM_URL } from '../services/PowerPlatformApiService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * „Pipelines“ (ALM): alle Pipelines-Hosts des Tenants, ihre Pipelines und die Environments darin – live aus dem Dataverse
 * der Hosts, nur lesen. Ziel-Environments müssen Managed sein (Microsoft erzwingt das seit Oktober 2026).
 */
export interface IPipelinesViewProps {
  service: DataverseService;
  factory: AadHttpClientFactory;
  /** Pipelines-Host aus den Webpart-Einstellungen (optional) */
  configuredHost: string;
  reloadKey: string;
  onOpenEnvironment: (coeId: string) => void;
}

type Tab = 'pipelines' | 'envs';
const ALL = '';
const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));
const csvEsc = (v: string): string => (/[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
const RUN_TONE: { [s in RunState]: 'success' | 'danger' | 'info' | 'warning' | 'neutral' } = { ok: 'success', failed: 'danger', running: 'info', pending: 'warning', canceled: 'neutral', '': 'neutral' };

interface IPipelineRow extends IPipeline { host: IPipelineHost; data: IHostData }
interface IEnvRow extends IPlEnv { host: IPipelineHost }

export const PipelinesView: React.FC<IPipelinesViewProps> = props => {
  const [scan, setScan] = React.useState<IHostScan | undefined>(undefined);
  const [scanning, setScanning] = React.useState<{ done: number; total: number } | undefined>(undefined);
  const [hosts, setHosts] = React.useState<IPipelineHost[]>([]);
  const [data, setData] = React.useState<{ [url: string]: IHostData }>({});
  const [problem, setProblem] = React.useState<string>('');
  const [host, setHost] = React.useState<string>(ALL);
  const [tab, setTab] = React.useState<Tab>('pipelines');
  const [search, setSearch] = React.useState<string>('');
  const [onlyIssues, setOnlyIssues] = React.useState<boolean>(false);
  const [open, setOpen] = React.useState<string>('');
  const gen = React.useRef<number>(0);

  const load = async (fresh: boolean, rescan: boolean): Promise<void> => {
    const my = ++gen.current;
    const alive = (): boolean => gen.current === my;
    setProblem('');
    if (fresh || rescan) {
      resetPipelineCache(rescan);
      setData({});
    }
    let list: ICoeEnv[];
    try {
      list = await loadCoeEnvs(props.service, fresh || rescan);
    } catch (e) {
      if (alive()) {
        setProblem(`Environments aus dem CoE nicht lesbar: ${errText(e)}`);
      }
      return;
    }
    if (!alive()) {
      return;
    }
    let s = rescan ? undefined : readHostScan();
    if (!s) {
      setScanning({ done: 0, total: list.filter(e => /^https:\/\//i.test(e.url)).length });
      s = await scanHosts(props.factory, list, (done, total) => { if (alive()) { setScanning({ done, total }); } }, alive);
      if (!alive()) {
        return;
      }
      setScanning(undefined);
    }
    setScan(s);
    const h = mergeHosts(props.configuredHost, s, list);
    setHosts(h);
    h.forEach(x => {
      loadHost(props.factory, x, list, fresh || rescan).then(d => { if (alive()) { setData(prev => ({ ...prev, [x.url]: d })); } }, () => undefined);
    });
  };
  React.useEffect(() => { load(false, false).catch(e => setProblem(errText(e))); }, [props.service, props.configuredHost, props.reloadKey]);

  const loaded = hosts.map(h => data[h.url]).filter(d => !!d);
  const shown = loaded.filter(d => !host || d.host.url === host);
  const pipelines: IPipelineRow[] = [];
  shown.forEach(d => d.pipelines.forEach(p => pipelines.push({ ...p, host: d.host, data: d })));
  const envRows: IEnvRow[] = [];
  shown.forEach(d => d.envs.forEach(e => envRows.push({ ...e, host: d.host })));
  const envName = (d: IHostData, id: string): string => {
    const e = d.envs.filter(x => x.id === id)[0];
    return e ? (e.coe ? e.coe.displayName : e.name) : '–';
  };
  const envOf = (d: IHostData, id: string): IPlEnv | undefined => d.envs.filter(x => x.id === id)[0];
  const unmanagedTarget = (e: IPlEnv): boolean => e.role === 'target' && !!e.coe && e.coe.managed === false;
  const issue = (e: IPlEnv): boolean => unmanagedTarget(e) || e.validationOk === false;
  const pipelineIssue = (p: IPipelineRow): boolean => p.stages.some(s => { const e = envOf(p.data, s.targetId); return !!e && issue(e); }) || (!!p.lastRun && p.lastRun.state === 'failed');

  const q = search.trim().toLowerCase();
  const visiblePipelines = pipelines.filter(p => (!onlyIssues || pipelineIssue(p)) && (!q || p.name.toLowerCase().indexOf(q) >= 0 || p.description.toLowerCase().indexOf(q) >= 0
    || p.owner.toLowerCase().indexOf(q) >= 0 || p.devEnvIds.concat(p.stages.map(s => s.targetId)).some(id => envName(p.data, id).toLowerCase().indexOf(q) >= 0)));
  const visibleEnvs = envRows.filter(e => (!onlyIssues || issue(e)) && (!q || e.name.toLowerCase().indexOf(q) >= 0 || (e.coe && e.coe.displayName.toLowerCase().indexOf(q) >= 0)
    || e.envId.toLowerCase().indexOf(q) >= 0 || e.pipelines.some(p => p.name.toLowerCase().indexOf(q) >= 0)));
  const unmanagedCount = envRows.filter(unmanagedTarget).length;
  const failedValidation = envRows.filter(e => e.validationOk === false).length;

  const runCell = (r: IPlRun | undefined): React.ReactNode => (r
    ? (
      <span className={styles.ulUser}>
        <span><Pill tone={RUN_TONE[r.state]}>{r.status || '–'}</Pill></span>
        <span className={styles.muted}>{r.start ? formatDate(new Date(r.start), true) : ''} · {r.artifact} {r.version}</span>
      </span>
    )
    : <span className={styles.muted}>keine</span>);
  const envChip = (d: IHostData, id: string, target: boolean, label?: string): React.ReactNode => {
    const e = envOf(d, id);
    const bad = !!e && (unmanagedTarget(e) || e.validationOk === false);
    return (
      <span key={`${id}${label || ''}`} className={`${styles.dlpChip} ${bad ? styles.dlpChipBlocked : target ? styles.dlpChipBusiness : styles.dlpChipGeneral}`}
        title={e ? `${e.coe ? `${e.coe.type}${e.coe.managed === false ? ', nicht Managed' : e.coe.managed ? ', Managed' : ''}` : 'nicht im CoE'}${e.validationOk === false ? ' – Prüfung fehlgeschlagen' : ''}` : ''}>
        {label ? `${label}: ` : ''}{envName(d, id)}
      </span>
    );
  };

  const pipelineCols: IColumn<IPipelineRow>[] = ([] as IColumn<IPipelineRow>[]).concat(!host && hosts.length > 1 ? [{
    key: 'host', label: 'Host', sort: (p: IPipelineRow) => p.host.name.toLowerCase(), render: (p: IPipelineRow) => p.host.name
  }] : []).concat([
    {
      key: 'name', label: 'Pipeline', sort: p => p.name.toLowerCase(),
      render: p => (
        <span className={styles.ulUser}>
          <a className={styles.link} href={`${p.host.url}/main.aspx?pagetype=entityrecord&etn=deploymentpipeline&id=${p.id}`} target="_blank" rel="noopener noreferrer"
            onClick={e => e.stopPropagation()} title="In der App „Deployment Pipeline Configuration“ öffnen">{p.name || '(ohne Namen)'}</a>
          {p.description && <span className={styles.muted}>{p.description}</span>}
        </span>
      )
    },
    {
      key: 'dev', label: 'Entwicklung', sort: p => p.devEnvIds.map(id => envName(p.data, id)).join(', ').toLowerCase(),
      render: p => (p.devEnvIds.length ? <span className={styles.dlpChips}>{p.devEnvIds.map(id => envChip(p.data, id, false))}</span> : <span className={styles.muted}>–</span>)
    },
    {
      key: 'stages', label: 'Phasen → Ziel', sort: p => p.stages.length,
      render: p => (p.stages.length
        ? <span className={styles.dlpChips}>{p.stages.map((s, i) => <React.Fragment key={s.id}>{i > 0 && <Icon iconName="ChevronRight" />}{envChip(p.data, s.targetId, true, s.name)}</React.Fragment>)}</span>
        : <span className={styles.muted}>keine Phase</span>)
    },
    { key: 'owner', label: 'Besitzer', sort: p => p.owner.toLowerCase(), render: p => p.owner || <span className={styles.muted}>–</span> },
    { key: 'last', label: 'Letzte Bereitstellung', sort: p => (p.lastRun ? p.lastRun.start : ''), render: p => runCell(p.lastRun) },
    { key: 'runs', label: 'Bereitstellungen', numeric: true, title: 'unter den letzten 1.000 des Hosts', sort: p => p.runCount, render: p => formatNumber(p.runCount) }
  ]);

  const envCols: IColumn<IEnvRow>[] = ([] as IColumn<IEnvRow>[]).concat(!host && hosts.length > 1 ? [{
    key: 'host', label: 'Host', sort: (e: IEnvRow) => e.host.name.toLowerCase(), render: (e: IEnvRow) => e.host.name
  }] : []).concat([
    {
      key: 'name', label: 'Environment', sort: e => (e.coe ? e.coe.displayName : e.name).toLowerCase(),
      render: e => (
        <span className={styles.ulUser}>
          {e.coe
            ? <button type="button" className={styles.linkButton} onClick={() => props.onOpenEnvironment(e.coe!.coeId)}>{e.coe.displayName}</button>
            : <b>{e.name}</b>}
          <span className={styles.muted}>{e.envId}</span>
        </span>
      )
    },
    { key: 'role', label: 'Rolle', sort: e => e.role, render: e => (e.role === 'target' ? <Pill tone="info">Ziel</Pill> : e.role === 'dev' ? <Pill tone="neutral">Entwicklung</Pill> : <span className={styles.muted}>–</span>) },
    {
      key: 'pipelines', label: 'Pipelines', sort: e => e.pipelines.length,
      render: e => (e.pipelines.length
        ? <span className={styles.dlpChips}>{e.pipelines.map(p => <span key={`${p.id}${p.as}${p.stage || ''}`} className={`${styles.dlpChip} ${p.as === 'target' ? styles.dlpChipBusiness : styles.dlpChipGeneral}`}>{p.name}{p.stage ? ` · ${p.stage}` : ''}</span>)}</span>
        : <span className={styles.muted}>in keiner Pipeline</span>)
    },
    {
      key: 'managed', label: 'Managed', sort: e => (e.coe && e.coe.managed !== undefined ? (e.coe.managed ? 1 : 0) : -1),
      render: e => (!e.coe || e.coe.managed === undefined
        ? <span className={styles.muted}>–</span>
        : e.coe.managed
          ? <Pill tone="success">Ja</Pill>
          : unmanagedTarget(e)
            ? <Pill tone="danger" title="Ziel-Environments müssen Managed sein – sonst blockiert Microsoft Bereitstellungen nach der Übergangsfrist">Nein – Ziel!</Pill>
            : <Pill tone="neutral">Nein</Pill>)
    },
    {
      key: 'validation', label: 'Prüfstatus', sort: e => e.validation,
      render: e => (e.validation ? <Pill tone={e.validationOk ? 'success' : e.validationOk === false ? 'danger' : 'warning'} title={e.error || undefined}>{e.validation}</Pill> : <span className={styles.muted}>–</span>)
    },
    { key: 'type', label: 'Typ', sort: e => (e.coe ? e.coe.type : ''), render: e => (e.coe ? e.coe.type : <span className={styles.muted}>nicht im CoE</span>) }
  ]);

  const exportCsv = (): void => {
    const lines = tab === 'pipelines'
      ? ['Host;Pipeline;Beschreibung;Entwicklung;Phasen → Ziel;Besitzer;Letzte Bereitstellung;Status;Solution;Version'].concat(visiblePipelines.map(p => [
        p.host.name, p.name, p.description, p.devEnvIds.map(id => envName(p.data, id)).join(', '),
        p.stages.map(s => `${s.name}: ${envName(p.data, s.targetId)}`).join(' → '), p.owner,
        p.lastRun && p.lastRun.start ? formatDate(new Date(p.lastRun.start), true) : '', p.lastRun ? p.lastRun.status : '', p.lastRun ? p.lastRun.artifact : '', p.lastRun ? p.lastRun.version : ''
      ].map(csvEsc).join(';')))
      : ['Host;Environment;Environment-ID;Rolle;Pipelines;Managed;Prüfstatus;Typ'].concat(visibleEnvs.map(e => [
        e.host.name, e.coe ? e.coe.displayName : e.name, e.envId, e.role === 'target' ? 'Ziel' : e.role === 'dev' ? 'Entwicklung' : '',
        e.pipelines.map(p => `${p.name}${p.stage ? ` (${p.stage})` : ''}`).join(', '), e.coe && e.coe.managed !== undefined ? (e.coe.managed ? 'Ja' : 'Nein') : '',
        e.validation, e.coe ? e.coe.type : ''
      ].map(csvEsc).join(';')));
    downloadText(`Pipelines_${tab === 'pipelines' ? 'Pipelines' : 'Environments'}_${new Date().toISOString().slice(0, 10)}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };

  const rowKey = (p: IPipelineRow): string => `${p.host.url}|${p.id}`;
  const expanded = (p: IPipelineRow): React.ReactNode => {
    if (open !== rowKey(p)) {
      return undefined;
    }
    const runs = p.data.runs.filter(r => r.pipelineId === p.id).slice(0, 8);
    return (
      <div className={styles.dlpCols} style={{ padding: '4px 0 8px' }}>
        <div className={styles.dlpList}>
          <div className={styles.dlpListHead}>Environments der Pipeline</div>
          <div className={styles.dlpListBody}>
            {p.devEnvIds.map(id => { const e = envOf(p.data, id); return (
              <div key={`d${id}`} className={styles.dlpListRow}><Pill tone="neutral">Entwicklung</Pill><span><b>{envName(p.data, id)}</b></span><span className={styles.grow} />
                {e && e.validation && <Pill tone={e.validationOk ? 'success' : e.validationOk === false ? 'danger' : 'warning'} title={e.error || undefined}>{e.validation}</Pill>}</div>
            ); })}
            {p.stages.map(s => { const e = envOf(p.data, s.targetId); return (
              <div key={`s${s.id}`} className={styles.dlpListRow}><Pill tone="info">{s.name || 'Ziel'}</Pill><span><b>{envName(p.data, s.targetId)}</b></span><span className={styles.grow} />
                {e && e.coe && e.coe.managed === false && <Pill tone="danger">nicht Managed</Pill>}
                {e && e.coe && e.coe.managed && <Pill tone="success">Managed</Pill>}
                {e && e.validation && <Pill tone={e.validationOk ? 'success' : e.validationOk === false ? 'danger' : 'warning'} title={e.error || undefined}>{e.validation}</Pill>}</div>
            ); })}
            {!p.devEnvIds.length && !p.stages.length && <div className={styles.dlpListRow}><span className={styles.muted}>Keine Environments.</span></div>}
          </div>
        </div>
        <div className={styles.dlpList}>
          <div className={styles.dlpListHead}>Letzte Bereitstellungen</div>
          <div className={styles.dlpListBody}>
            {runs.map(r => (
              <div key={r.id} className={styles.dlpListRow}><Pill tone={RUN_TONE[r.state]}>{r.status || '–'}</Pill><span>{r.artifact} {r.version}</span><span className={styles.grow} />
                <span className={styles.muted}>{r.start ? formatDate(new Date(r.start), true) : ''}{r.by ? ` · ${r.by}` : ''}</span></div>
            ))}
            {!runs.length && <div className={styles.dlpListRow}><span className={styles.muted}>Keine Bereitstellungen.</span></div>}
          </div>
        </div>
      </div>
    );
  };

  const hostTile = (h: IPipelineHost): JSX.Element => {
    const d = data[h.url];
    const hint = !d ? 'wird gelesen …' : d.problem ? (d.problemKind === 'forbidden' ? 'kein Zugriff' : 'nicht lesbar')
      : `${formatNumber(d.envs.length)} Environments${h.configured ? ' · Standard-Host' : ''}`;
    return (
      <Tile key={h.url} label={h.name} hint={hint} icon="Rocket" tone={d && d.problem ? styles.ovToneSerious : styles.ovToneInfo}
        value={d && !d.problem ? formatNumber(d.pipelines.length) : '–'} active={host === h.url} onClick={() => setHost(host === h.url ? ALL : h.url)} />
    );
  };
  const problems = loaded.filter(d => !!d.problem);
  const devNotes = loaded.filter(d => !!d.devNote);
  const totalPipelines = loaded.reduce((s, d) => s + d.pipelines.length, 0);
  const loading = !!scanning || hosts.some(h => !data[h.url]);

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="Rocket" /> Live aus den Pipelines-Hosts · nur lesen</div>
          <div className={styles.secHeroTitle}>Pipelines</div>
          <div className={styles.ovHeroLabel}>
            Welche Pipelines-Hosts es gibt, welche Pipelines darin liegen und welche Environments sie verbinden – Entwicklung und Ziele je Phase.
            Ziel-Environments müssen Managed sein; Microsoft blockiert Bereitstellungen in nicht verwaltete Ziele seit Oktober 2026 nach einer Übergangsfrist.
          </div>
          <div className={styles.ovHeroChips}>
            <button type="button" className={styles.ovHeroChip} onClick={() => { load(true, false).catch(e => setProblem(errText(e))); }} disabled={loading}><Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}</button>
            <button type="button" className={styles.ovHeroChip} onClick={() => { load(true, true).catch(e => setProblem(errText(e))); }} disabled={!!scanning}><Icon iconName="Search" /> Hosts neu suchen</button>
            <a className={styles.ovHeroChip} href={PPAC_DEPLOYMENT_URL} target="_blank" rel="noopener noreferrer"><Icon iconName="OpenInNewTab" /> Admin Center (Deployment)</a>
            {scan && <span className={styles.ccStamp}>Hosts gesucht: {formatDate(new Date(scan.at), true)}</span>}
          </div>
        </div>
      </section>

      {problem && <div className={styles.ccWarnings}><div className={styles.muted}><Icon iconName="Warning" /> {problem}</div></div>}
      {scanning && (
        <section className={styles.ovCard} style={{ padding: '14px 24px' }}>
          <ProgressIndicator label="Pipelines-Hosts werden gesucht" description={`${formatNumber(scanning.done)} von ${formatNumber(scanning.total)} Dataverse-Environments geprüft (einmal je Tag, Ergebnis bleibt im Browser)`}
            percentComplete={scanning.total ? scanning.done / scanning.total : undefined} />
        </section>
      )}

      {!scanning && (
        <div className={styles.secGrid} data-pl-hosts="">
          <Tile label="Alle Hosts" hint={`${formatNumber(hosts.length)} Hosts gefunden`} icon="ViewAll" tone={styles.ovToneInfo} value={formatNumber(totalPipelines)} active={!host} onClick={() => setHost(ALL)} />
          {hosts.map(hostTile)}
          <Tile label="Platform host" hint="persönliche Pipelines der Maker – nur im Admin Center" icon="CloudWeather" tone={styles.ovToneInfo} value="–" />
        </div>
      )}

      {!scanning && (scan || problems.length || unmanagedCount || failedValidation || devNotes.length) ? (
        <div className={styles.ccWarnings}>
          {scan && (
            <div className={styles.muted}>
              <Icon iconName="Info" /> {formatNumber(scan.hosts.length)} Hosts in {formatNumber(scan.checked)} Dataverse-Environments aus dem CoE gefunden
              {scan.denied ? ` · ${formatNumber(scan.denied)} ohne Zugriff (nicht geprüft)` : ''}{scan.failed ? ` · ${formatNumber(scan.failed)} nicht erreichbar` : ''}.
              {scan.denied ? <> <a className={styles.link} href={PIM_URL} target="_blank" rel="noopener noreferrer">Rolle in PIM aktivieren</a> und „Hosts neu suchen“.</> : ''}
            </div>
          )}
          {problems.map(d => <div key={d.host.url} className={styles.muted}><Icon iconName="Warning" /> {d.host.name}: {d.problem}</div>)}
          {unmanagedCount > 0 && (
            <div className={styles.muted}>
              <Icon iconName="Warning" /> {formatNumber(unmanagedCount)} Ziel-Environment{unmanagedCount === 1 ? ' ist' : 's sind'} nicht Managed – Bereitstellungen dorthin werden nach der Übergangsfrist blockiert
              {' '}<button type="button" className={styles.textButton} onClick={() => { setTab('envs'); setOnlyIssues(true); }}>anzeigen</button>
            </div>
          )}
          {failedValidation > 0 && <div className={styles.muted}><Icon iconName="Warning" /> {formatNumber(failedValidation)} Environment{failedValidation === 1 ? '' : 's'} mit fehlgeschlagener Prüfung (Prüfstatus).</div>}
          {devNotes.map(d => <div key={`n${d.host.url}`} className={styles.muted}><Icon iconName="Info" /> {d.host.name}: {d.devNote}</div>)}
        </div>
      ) : null}

      {!scanning && (
        <section className={styles.ovCard}>
          <div className={styles.toolbar}>
            <div className={styles.ccSegments} role="tablist" aria-label="Ansicht">
              <button type="button" role="tab" aria-selected={tab === 'pipelines'} data-pl-tab="pipelines" className={`${styles.ccSegment} ${tab === 'pipelines' ? styles.ccSegmentActive : ''}`} onClick={() => setTab('pipelines')}>
                Pipelines <span className={styles.ccSegmentCount}>{formatNumber(pipelines.length)}</span>
              </button>
              <button type="button" role="tab" aria-selected={tab === 'envs'} data-pl-tab="envs" className={`${styles.ccSegment} ${tab === 'envs' ? styles.ccSegmentActive : ''}`} onClick={() => setTab('envs')}>
                Environments <span className={styles.ccSegmentCount}>{formatNumber(envRows.length)}</span>
              </button>
            </div>
            <button type="button" className={`${styles.ccSegment} ${onlyIssues ? styles.ccSegmentActive : ''}`} aria-pressed={onlyIssues} onClick={() => setOnlyIssues(!onlyIssues)}>
              <Icon iconName="Warning" /> Nur mit Hinweis
            </button>
            <SearchBox placeholder={tab === 'pipelines' ? 'Pipeline, Environment oder Person' : 'Environment oder Pipeline'} value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 260 } }} />
            <span className={styles.resultBadge} role="status"><strong>{formatNumber(tab === 'pipelines' ? visiblePipelines.length : visibleEnvs.length)}</strong> {tab === 'pipelines' ? 'Pipelines' : 'Environments'}</span>
            <span className={styles.grow} />
            <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportCsv} disabled={tab === 'pipelines' ? !visiblePipelines.length : !visibleEnvs.length} />
          </div>
          {loading && !loaded.length && <div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>Pipelines werden gelesen …</span></div>}
          {tab === 'pipelines' && (loaded.length > 0 || !loading) && (
            <SortTable<IPipelineRow>
              columns={pipelineCols}
              rows={visiblePipelines}
              rowKey={rowKey}
              initial={{ key: 'name', desc: false }}
              empty={hosts.length ? 'Keine Pipelines für diese Auswahl.' : 'Keine Pipelines-Hosts gefunden.'}
              expanded={expanded}
              onRow={p => setOpen(open === rowKey(p) ? '' : rowKey(p))}
              rowClass={p => (open === rowKey(p) ? styles.ulRowSelected : undefined)}
            />
          )}
          {tab === 'envs' && (loaded.length > 0 || !loading) && (
            <SortTable<IEnvRow>
              columns={envCols}
              rows={visibleEnvs}
              rowKey={e => `${e.host.url}|${e.id}`}
              initial={{ key: 'role', desc: true }}
              empty={hosts.length ? 'Keine Environments für diese Auswahl.' : 'Keine Pipelines-Hosts gefunden.'}
            />
          )}
          <div className={styles.ccFoot}>
            Live aus dem Dataverse der Hosts (Tabellen „Deployment Pipeline“, „Deployment Stage“, „Deployment Environment“, „Deployment Stage Run“) mit deinen Rechten;
            Managed-Status aus dem CoE. Zeile anklicken = Details, Pipeline-Name öffnet die App „Deployment Pipeline Configuration“ im Host. Ändern nur im Host bzw. im Admin Center.
          </div>
        </section>
      )}
    </div>
  );
};
