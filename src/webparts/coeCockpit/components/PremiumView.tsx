import * as React from 'react';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton } from '@fluentui/react/lib/Button';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable, Tile } from './CopilotCreditsView';
import { Tone } from '../models/ITableConfig';
import {
  IPremiumApp, IPremiumData, IPremiumEnv, IPremiumUser, LICENSE_LABEL, LicenseStatus, loadPremiumImpact, PREMIUM_DEADLINE, PREMIUM_INFO_URL, USAGE_DAYS
} from '../models/premiumLicensing';
import { DataverseService } from '../services/DataverseService';
import { IPipelineIndex, pipelineInfo } from '../models/pipelineHosts';
import { PowerPlatformApiService, PPAPI_VERSIONS } from '../services/PowerPlatformApiService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * Ansicht „Betroffen ab Februar 2027“: Apps in Managed Environments öffnen dann nur noch mit Premium-Lizenz.
 * Zeigt alle Berechtigten der Apps (Besitzer, Mitbesitzer, Benutzer, Mitglieder – auch verschachtelter – Gruppen),
 * ihre Lizenzen (live über Microsoft Graph) und wer die Apps in den letzten 30 Tagen geöffnet hat.
 */
export interface IPremiumViewProps {
  service: DataverseService;
  graphFactory?: MSGraphClientFactory;
  ppApi?: PowerPlatformApiService;
  managedField: string;
  reloadKey: string;
  onOpenApp: (appId: string) => void;
  /** Pipelines aller Hosts (Environments, die Ziel einer Pipeline sind, müssen Managed bleiben) */
  loadPipelines?: () => Promise<IPipelineIndex>;
}

// Zwischenspeicher über Ansichtswechsel hinweg
let cache: { key: string; data: IPremiumData } | undefined;

/** Nur für Tests/Vorschau: Zwischenspeicher leeren. */
export function resetPremiumCache(): void {
  cache = undefined;
}

const LICENSE_TONE: { [k in LicenseStatus]: Tone } = { premium: 'success', trial: 'warning', perApp: 'warning', none: 'danger', notfound: 'neutral', unknown: 'neutral' };
const isBlocked = (u: IPremiumUser): boolean => u.license === 'none' || u.license === 'perApp';

type Mode = 'users' | 'envs' | 'apps';
type UserFilter = 'all' | 'blocked' | 'premium' | 'other' | 'used';

/** Pay-as-you-go für Power Apps je Environment (über die Power Platform API, falls freigegeben). */
async function payGoEnvironments(api: PowerPlatformApiService): Promise<{ [envId: string]: string }> {
  const out: { [envId: string]: string } = {};
  const plans = await api.getAll<{ id?: string; name?: string; payGoEntitlements?: { productCategory?: string; payAsYouGoState?: boolean }[] }>('/licensing/billingPolicies', PPAPI_VERSIONS.billing);
  await Promise.all(plans.filter(p => (p.payGoEntitlements || []).some(e => !!e.payAsYouGoState && /powerapps/i.test(e.productCategory || ''))).map(async p => {
    const envs = await api.getAll<{ environmentId?: string }>(`/licensing/billingPolicies/${encodeURIComponent(String(p.id))}/environments`, PPAPI_VERSIONS.billing).catch(() => [] as { environmentId?: string }[]);
    envs.forEach(e => { if (e.environmentId) { out[String(e.environmentId).toLowerCase()] = String(p.name || p.id); } });
  }));
  return out;
}

export const PremiumView: React.FC<IPremiumViewProps> = props => {
  const key = `${props.service.baseUrl}|${props.managedField}|${props.reloadKey}`;
  const [data, setData] = React.useState<IPremiumData | undefined>(cache && cache.key === key ? cache.data : undefined);
  const [error, setError] = React.useState<string>('');
  const [progress, setProgress] = React.useState<string>('');
  const [loading, setLoading] = React.useState<boolean>(false);
  const [mode, setMode] = React.useState<Mode>('users');
  const [filter, setFilter] = React.useState<UserFilter>('all');
  const [search, setSearch] = React.useState<string>('');
  const [open, setOpen] = React.useState<string>('');
  const [pipes, setPipes] = React.useState<IPipelineIndex | undefined>(undefined);
  const [pipeState, setPipeState] = React.useState<'off' | 'loading' | 'ready' | 'error'>(props.loadPipelines ? 'loading' : 'off');
  const [pipeProblem, setPipeProblem] = React.useState<string>('');
  const [onlyPipelineEnvs, setOnlyPipelineEnvs] = React.useState<boolean>(false);
  React.useEffect(() => {
    if (!props.loadPipelines) {
      setPipeState('off');
      return;
    }
    let alive = true;
    setPipeState('loading');
    props.loadPipelines().then(ix => { if (alive) { setPipes(ix); setPipeState('ready'); } },
      (e: Error) => { if (alive) { setPipeProblem(e.message); setPipeState('error'); } });
    return () => { alive = false; };
  }, [props.loadPipelines, props.reloadKey]);

  const load = (force: boolean): void => {
    if (!force && cache && cache.key === key) {
      setData(cache.data);
      return;
    }
    setLoading(true);
    setError('');
    loadPremiumImpact(props.service, props.graphFactory, {
      managedField: props.managedField,
      onProgress: setProgress,
      payGo: props.ppApi ? () => payGoEnvironments(props.ppApi!) : undefined
    }).then(d => {
      cache = { key, data: d };
      setData(d);
      setLoading(false);
    }, (e: Error) => { setError(e.message); setLoading(false); });
  };
  React.useEffect(() => { load(false); }, [key]);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const days = Math.round((PREMIUM_DEADLINE.getTime() - today.getTime()) / 86400000);

  const hero = (
    <section className={`${styles.ovHero} ${styles.secHero}`}>
      <div className={styles.ovHeroGlow} aria-hidden="true" />
      <div className={styles.ovHeroMain}>
        <div className={styles.ovHeroEyebrow}><Icon iconName="Permissions" /> Lizenzpflicht in Managed Environments</div>
        <div className={styles.secHeroTitle}>Betroffen ab Februar 2027</div>
        <div className={styles.ovHeroLabel}>
          Ab Februar 2027 öffnen Apps in <b>Managed Environments</b> nur noch für Personen mit Premium-Lizenz (Power Apps Premium, Power Automate Premium
          oder Dynamics 365 Enterprise) – auch Standard-Apps. Angezeigt werden alle Berechtigten der Apps und ihre Lizenzen.
        </div>
        <div className={styles.ovHeroChips}>
          <button type="button" className={styles.ovHeroChip} onClick={() => load(true)} disabled={loading}><Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}</button>
          <a className={styles.ovHeroChip} href={PREMIUM_INFO_URL} target="_blank" rel="noopener noreferrer"><Icon iconName="OpenInNewWindow" /> Information von Microsoft</a>
          {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
        </div>
      </div>
      <div className={styles.ovCountdown}>
        <span className={styles.ovCountdownTop}><Icon iconName="Clock" /> Sperre ohne Premium-Lizenz</span>
        <span className={styles.ovCountdownDays}>{days > 0 ? formatNumber(days) : '0'}</span>
        <span className={styles.ovCountdownUnit}>{days === 1 ? 'Tag' : 'Tage'} bis Februar 2027</span>
        <span className={styles.ovCountdownFoot}>{data
          ? data.licenseState !== 'ok' ? 'Lizenzen nicht geprüft'
            : <><strong>{data.warnings.some(w => /Personen nicht geprüft/.test(w)) ? 'mind. ' : ''}{formatNumber(data.users.filter(isBlocked).length)}</strong>&nbsp;Personen ohne Premium</>
          : 'laut Microsoft „ab Februar 2027“'}</span>
      </div>
    </section>
  );

  if (error || (!data && loading) || !data) {
    return (
      <div className={styles.ov}>
        {hero}
        <div className={styles.ovCard}>
          {error
            ? <div className={styles.ovFailed}><Icon iconName="Error" /> {error}</div>
            : <div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>{progress || 'Wird geladen …'}</span></div>}
        </div>
      </div>
    );
  }

  const appById: { [id: string]: IPremiumApp } = {};
  data.apps.forEach(a => { appById[a.id] = a; });
  const blocked = data.users.filter(isBlocked);
  const usedApps = data.apps.filter(a => a.users.length > 0);
  const envsWithImpact = data.envs.filter(e => e.unlicensed > 0);
  const q = search.trim().toLowerCase();

  const notices = (
    <>
      {data.licenseState === 'error' && (
        <div className={styles.accessNotice} role="alert">
          <span className={styles.accessIcon}><Icon iconName="Warning" /></span>
          <div className={styles.ccNoticeText}>
            <strong>Lizenzprüfung abgebrochen</strong>
            <span>„Ohne Premium“ ist deshalb nicht aussagekräftig – bitte später „Aktualisieren“.</span>
            {data.licenseProblem && <span className={styles.ccDetail}>Technisch: {data.licenseProblem}</span>}
          </div>
        </div>
      )}
      {(data.licenseState === 'forbidden' || data.licenseState === 'consent') && (
        <div className={styles.accessNotice} role="alert">
          <span className={styles.accessIcon}><Icon iconName="Lock" /></span>
          <div className={styles.ccNoticeText}>
            <strong>Lizenzen konnten nicht geprüft werden</strong>
            <span>Das Webpart braucht die Berechtigung „Microsoft Graph – User.Read.All“ (für Gruppen zusätzlich „GroupMember.Read.All“). Ein Admin setzt sie mit dem Skript <code>scripts/Grant-CoeCockpitPermissions.ps1 -Licenses</code>.</span>
            {data.licenseProblem && <span className={styles.ccDetail}>Technisch: {data.licenseProblem}</span>}
          </div>
        </div>
      )}
      {data.unresolvedGroups.length > 0 && (
        <div className={styles.ccWarnings}>
          <div className={styles.muted}><Icon iconName="Info" /> {formatNumber(data.unresolvedGroups.length)} Gruppe(n) mit Freigaben konnten nicht aufgelöst werden
            {data.groupState === 'forbidden' ? ' (Berechtigung „GroupMember.Read.All“ fehlt)' : ''}: {data.unresolvedGroups.slice(0, 6).map(g => `${g.name}${g.size ? ` (${formatNumber(g.size)})` : ''}`).join(', ')}{data.unresolvedGroups.length > 6 ? ' …' : ''}</div>
        </div>
      )}
      {(!data.usageAvailable || data.warnings.length > 0) && (
        <div className={styles.ccWarnings}>
          {!data.usageAvailable && <div className={styles.muted}><Icon iconName="Info" /> App-Starts sind ohne CoE-Audit-Log nicht verfügbar – angezeigt werden die Berechtigungen.</div>}
          {data.warnings.map((w, i) => <div key={i} className={styles.muted}><Icon iconName="Info" /> {w}</div>)}
        </div>
      )}
    </>
  );

  const appNames = (u: IPremiumUser): string[] => u.access.map(a => (appById[a.appId] ? appById[a.appId].name : a.appId)).filter((x, i, arr) => arr.indexOf(x) === i);
  const envNames = (u: IPremiumUser): string[] => u.access.map(a => (appById[a.appId] ? appById[a.appId].envName : '')).filter((x, i, arr) => !!x && arr.indexOf(x) === i);
  const viaLabel = (a: { role: string; via: string }): string => (a.via === 'direkt' ? a.role : a.via === 'geöffnet' ? 'ohne Freigabe geöffnet' : a.via === 'Organisation' ? 'Organisation (für alle freigegeben)' : `${a.role} über ${a.via}`);
  const viaText = (u: IPremiumUser): string => u.access.map(viaLabel).filter((x, i, arr) => arr.indexOf(x) === i).join(', ');

  const userFilters: { key: UserFilter; label: string; test: (u: IPremiumUser) => boolean }[] = [
    { key: 'all', label: 'Alle', test: () => true },
    { key: 'blocked', label: 'Ohne Premium', test: isBlocked },
    { key: 'premium', label: 'Premium', test: u => u.license === 'premium' },
    { key: 'other', label: 'Testversion / ungeprüft', test: u => u.license === 'trial' || u.license === 'unknown' || u.license === 'notfound' },
    { key: 'used', label: `Geöffnet (${USAGE_DAYS} Tage)`, test: u => u.launches > 0 }
  ];
  const uf = userFilters.filter(f => f.key === filter)[0];
  const visibleUsers = data.users.filter(u => uf.test(u) && (!q || `${u.name} ${u.upn} ${appNames(u).join(' ')} ${envNames(u).join(' ')}`.toLowerCase().indexOf(q) >= 0));

  const userCols: IColumn<IPremiumUser>[] = [
    { key: 'name', label: 'Person', sort: u => u.name.toLowerCase(), render: u => <span className={styles.ccName} title={u.upn}><b>{u.name}</b>{u.guest ? ' (Gast)' : ''}<br /><span className={styles.muted}>{u.upn}</span></span> },
    { key: 'license', label: 'Lizenz', sort: u => LICENSE_LABEL[u.license], render: u => <Pill tone={LICENSE_TONE[u.license]}>{LICENSE_LABEL[u.license]}</Pill> },
    { key: 'licenses', label: 'Zugewiesene Lizenzen', sort: u => u.licenses.join(', '), render: u => <span className={styles.ccName} title={u.licenses.join(', ')}>{u.licenses.length ? u.licenses.join(', ') : '–'}</span> },
    { key: 'apps', label: 'Apps', numeric: true, sort: u => appNames(u).length, render: u => <span title={appNames(u).join('\n')}>{formatNumber(appNames(u).length)}</span> },
    { key: 'envs', label: 'Environments', sort: u => envNames(u).join(', '), render: u => <span className={styles.ccName} title={envNames(u).join('\n')}>{envNames(u).join(', ')}</span> },
    { key: 'launches', label: `Geöffnet (${USAGE_DAYS} T.)`, numeric: true, sort: u => u.launches, render: u => (data.usageAvailable ? formatNumber(u.launches) : '–') }
  ];
  const pipeOf = (e: IPremiumEnv): ReturnType<typeof pipelineInfo> => pipelineInfo(pipes, e.id);
  /** Nicht alle Hosts geprüft: „keine Pipeline“ ist dann nicht sicher */
  const pipesIncomplete = !!pipes && (pipes.notChecked > 0 || pipes.unreadable.length > 0);
  const pipelineTargets = data.envs.filter(e => { const i = pipeOf(e); return !!i && i.target.length > 0; });
  const isPipelineTarget = (e: IPremiumEnv): boolean => { const i = pipeOf(e); return !!i && i.target.length > 0; };
  const envCols: IColumn<IPremiumEnv>[] = ([
    { key: 'name', label: 'Environment', sort: e => e.name.toLowerCase(), render: e => <a className={`${styles.link} ${styles.ccName}`} href={`https://admin.powerplatform.microsoft.com/environments/${encodeURIComponent(e.id)}/hub`} target="_blank" rel="noopener noreferrer">{e.name}</a> },
    { key: 'type', label: 'Typ', sort: e => e.type, render: e => e.type || '–' },
    { key: 'apps', label: 'Apps', numeric: true, sort: e => e.apps, render: e => formatNumber(e.apps) },
    { key: 'users', label: 'Berechtigte', numeric: true, sort: e => e.users.length, render: e => formatNumber(e.users.length) },
    {
      key: 'unlicensed', label: 'Ohne Premium', numeric: true, sort: e => e.unlicensed,
      render: e => (e.payGo ? <span className={styles.muted} title={`Pay-as-you-go über „${e.payGo}“`}>gedeckt (PAYG)</span>
        : data.licenseState !== 'ok' ? <span className={styles.muted} title="Lizenzen nicht geprüft">–</span>
        : <span className={styles.ccConsumed}><b>{formatNumber(e.unlicensed)}</b>
          {e.users.length > 0 && <span className={styles.ccMeter} title={`${formatNumber(Math.round((e.unlicensed / e.users.length) * 100))} % ohne Premium`}><span className={`${styles.ccMeterFill} ${styles.ccMeterFull}`} style={{ width: `${(e.unlicensed / e.users.length) * 100}%` }} /></span>}
        </span>)
    }
  ] as IColumn<IPremiumEnv>[]).concat(pipeState === 'off' ? [] : [{
    key: 'pipelines', label: 'Pipelines', title: 'Ziel einer Pipeline: Managed lässt sich nicht zurücknehmen (Microsoft verlangt Managed für Pipeline-Ziele)',
    sort: (e: IPremiumEnv) => { const i = pipeOf(e); return i ? i.target.length * 100 + i.dev.length + (i.host ? 1 : 0) : -1; },
    render: (e: IPremiumEnv) => {
      if (pipeState === 'loading') {
        return <span className={styles.muted}>…</span>;
      }
      const i = pipeOf(e);
      if (!i) {
        return pipeState === 'error' ? <span className={styles.muted}>–</span>
          : pipesIncomplete ? <span className={styles.muted} title="Nicht alle Pipelines-Hosts konnten geprüft werden – siehe Hinweis unter der Tabelle">keine gefunden (?)</span>
            : <span className={styles.muted}>keine</span>;
      }
      return (
        <span className={styles.dlpChips}>
          {i.target.length > 0 && <Pill tone="danger" title={`Ziel in: ${i.target.join(', ')} – muss Managed bleiben`}>Ziel in {formatNumber(i.target.length)} Pipeline{i.target.length === 1 ? '' : 's'} – bleibt Managed</Pill>}
          {i.dev.length > 0 && <Pill tone="neutral" title={`Entwicklung in: ${i.dev.join(', ')}`}>Entwicklung ({formatNumber(i.dev.length)})</Pill>}
          {i.host && <Pill tone="info">Pipelines-Host</Pill>}
        </span>
      );
    }
  }]);
  const appCols: IColumn<IPremiumApp>[] = [
    { key: 'name', label: 'App', sort: a => a.name.toLowerCase(), render: a => <button type="button" className={`${styles.textButton} ${styles.ccName}`} onClick={e => { e.stopPropagation(); props.onOpenApp(a.id); }}>{a.name}</button> },
    { key: 'env', label: 'Environment', sort: a => a.envName, render: a => <span className={styles.ccName}>{a.envName}</span> },
    { key: 'owner', label: 'Besitzer', sort: a => a.owner, render: a => a.owner || '–' },
    { key: 'users', label: 'Berechtigte', numeric: true, sort: a => a.users.length, render: a => (a.tenantShared ? <span title="Für alle in der Organisation freigegeben">{formatNumber(a.users.length)} + Organisation</span> : formatNumber(a.users.length)) },
    { key: 'unlicensed', label: 'Ohne Premium', numeric: true, sort: a => a.unlicensed, render: a => (data.licenseState === 'ok' ? <b>{formatNumber(a.unlicensed)}</b> : <span className={styles.muted} title="Lizenzen nicht geprüft">–</span>) },
    { key: 'launches', label: `Starts (${USAGE_DAYS} T.)`, numeric: true, sort: a => a.launches, render: a => (data.usageAvailable ? formatNumber(a.launches) : '–') }
  ];

  const exportUsers = (): void => {
    const rows = [['Person', 'UPN', 'Lizenzstatus', 'Zugewiesene Lizenzen', 'Apps', 'Environments', 'Berechtigt als', `Geöffnet (${USAGE_DAYS} Tage)`]]
      .concat(visibleUsers.map(u => [u.name, u.upn, LICENSE_LABEL[u.license], u.licenses.join(', '), appNames(u).join(', '), envNames(u).join(', '), viaText(u), String(u.launches)]));
    const csv = '﻿' + rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(';')).join('\r\n');
    downloadText(`Betroffen_Februar_2027_${new Date().toISOString().substring(0, 10)}.csv`, csv, 'text/csv;charset=utf-8');
  };

  return (
    <div className={styles.ov}>
      {hero}
      {notices}
      <div className={styles.ccGrid}>
        <Tile label="Managed Environments" hint={`davon ${formatNumber(data.envs.filter(e => e.apps > 0).length)} mit Apps`} icon="Shield" tone={styles.ovToneInfo} value={formatNumber(data.envs.length)} onClick={() => { setMode('envs'); setOnlyPipelineEnvs(false); }} active={mode === 'envs' && !onlyPipelineEnvs} />
        <Tile label="Apps" hint={`${formatNumber(usedApps.length)} mit Berechtigten`} icon="AppIconDefault" tone={styles.ovToneInfo} value={formatNumber(data.apps.length)} onClick={() => setMode('apps')} active={mode === 'apps'} />
        <Tile label="Berechtigte Personen" hint={data.usageAvailable ? `${formatNumber(data.users.filter(u => u.launches > 0).length)} haben in ${USAGE_DAYS} Tagen eine App geöffnet` : 'Besitzer, Mitbesitzer, Benutzer, Gruppen'} icon="People" tone={styles.ovToneInfo} value={formatNumber(data.users.length)} onClick={() => { setMode('users'); setFilter('all'); }} active={mode === 'users' && filter === 'all'} />
        {pipeState !== 'off' && (
          <Tile label="Ziel einer Pipeline" hint={pipeState === 'loading' ? 'Pipelines werden gelesen …' : pipeState === 'error' ? 'Pipelines nicht lesbar' : 'Managed lässt sich hier nicht zurücknehmen'}
            icon="Rocket" tone={pipelineTargets.length ? styles.ovToneSerious : styles.ovToneInfo} value={pipeState === 'ready' ? formatNumber(pipelineTargets.length) : '–'}
            onClick={pipeState === 'ready' ? () => { setMode('envs'); setOnlyPipelineEnvs(true); } : undefined} active={mode === 'envs' && onlyPipelineEnvs} />
        )}
        <Tile label="Ohne Premium-Lizenz" hint={data.licenseState === 'ok' ? `in ${formatNumber(envsWithImpact.length)} Environment(s) – wären ab Februar 2027 gesperrt` : 'Lizenzen nicht geprüft'} icon="Blocked2" tone={blocked.length ? styles.ovToneCritical : styles.ovToneInfo} value={data.licenseState === 'ok' ? formatNumber(blocked.length) : '–'} onClick={() => { setMode('users'); setFilter('blocked'); }} active={mode === 'users' && filter === 'blocked'} />
      </div>
      <section className={styles.ovCard}>
        <div className={styles.toolbar}>
          <div className={styles.ccSegments} role="group" aria-label="Ansicht">
            {([['users', 'Personen'], ['envs', 'Environments'], ['apps', 'Apps']] as [Mode, string][]).map(([k, l]) => (
              <button key={k} type="button" className={`${styles.ccSegment} ${mode === k ? styles.ccSegmentActive : ''}`} aria-pressed={mode === k} onClick={() => setMode(k)}>{l}</button>
            ))}
          </div>
          <SearchBox placeholder="Person, App oder Environment suchen" value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 280, borderRadius: 6 } }} />
          {mode === 'users' && (
            <span className={`${styles.resultBadge} ${filter !== 'all' || q ? styles.resultBadgeActive : ''}`} role="status">
              <strong>{formatNumber(visibleUsers.length)}</strong> von {formatNumber(data.users.length)} Personen
            </span>
          )}
          <span className={styles.grow} />
          {mode === 'users' && <DefaultButton text="Liste exportieren (CSV)" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportUsers} disabled={!visibleUsers.length} />}
        </div>
        {mode === 'users' && (
          <div className={styles.toolbar} style={{ paddingTop: 0 }}>
            <div className={styles.ccSegments} role="group" aria-label="Filter">
              {userFilters.map(f => (
                <button key={f.key} type="button" className={`${styles.ccSegment} ${filter === f.key ? styles.ccSegmentActive : ''}`} aria-pressed={filter === f.key} onClick={() => setFilter(f.key)}>
                  {f.label} <span className={styles.ccSegmentCount}>{formatNumber(data.users.filter(f.test).length)}</span>
                </button>
              ))}
            </div>
          </div>
        )}
        {mode === 'users' && (
          <SortTable<IPremiumUser>
            columns={userCols}
            rows={visibleUsers}
            rowKey={u => u.upn || u.id}
            initial={{ key: 'apps', desc: true }}
            empty="Keine Personen für diesen Filter."
            onRow={u => setOpen(open === (u.upn || u.id) ? '' : (u.upn || u.id))}
            expanded={u => open === (u.upn || u.id) && (
              <ul className={styles.ccEnvList}>
                {u.access.map((a, i) => (
                  <li key={i}>
                    <button type="button" className={styles.textButton} onClick={e => { e.stopPropagation(); props.onOpenApp(a.appId); }}>{appById[a.appId] ? appById[a.appId].name : a.appId}</button>
                    <span className={styles.muted}> · {appById[a.appId] ? appById[a.appId].envName : ''} · {viaLabel(a)}{u.launchedApps.indexOf(a.appId) >= 0 && a.via !== 'geöffnet' ? ' · geöffnet' : ''}</span>
                  </li>
                ))}
              </ul>
            )}
          />
        )}
        {mode === 'envs' && (
          <SortTable<IPremiumEnv>
            columns={envCols}
            rows={data.envs.filter(e => (!q || e.name.toLowerCase().indexOf(q) >= 0) && (!onlyPipelineEnvs || isPipelineTarget(e)))}
            rowKey={e => e.id}
            initial={{ key: 'unlicensed', desc: true }}
            empty={onlyPipelineEnvs ? 'Kein Managed Environment ist Ziel einer Pipeline.' : 'Keine Managed Environments gefunden.'}
          />
        )}
        {mode === 'envs' && pipeState !== 'off' && (
          <div className={styles.ccFoot}>
            {pipeState === 'error'
              ? <span><Icon iconName="Warning" /> Pipelines nicht lesbar: {pipeProblem}</span>
              : <span>
                {pipesIncomplete && pipes && <><Icon iconName="Warning" /> <b>Ergebnis unvollständig:</b>
                  {pipes.notChecked > 0 ? ` ${formatNumber(pipes.notChecked)} Environments konnten bei der Host-Suche nicht geprüft werden (ohne Zugriff oder nicht erreichbar)` : ''}
                  {pipes.unreadable.length ? `${pipes.notChecked > 0 ? ';' : ''} nicht lesbar: ${pipes.unreadable.join(', ')}` : ''}
                  {' '}– „keine gefunden (?)“ heißt dann nicht sicher „keine Pipeline“. Details und „Hosts neu suchen“ in der Ansicht „Pipelines“.<br /></>}
                „Pipelines“: live aus {pipesIncomplete ? 'den lesbaren' : 'allen'} Pipelines-Hosts{pipes ? ` (${formatNumber(pipes.hosts)} Hosts)` : ''}.
                Ziel-Environments einer Pipeline müssen Managed sein – Managed lässt sich dort nicht zurücknehmen. Entwicklungs-Environments müssen nicht Managed sein.</span>}
            {onlyPipelineEnvs && <button type="button" className={styles.textButton} onClick={() => setOnlyPipelineEnvs(false)}>Alle Managed Environments zeigen</button>}
          </div>
        )}
        {mode === 'apps' && (
          <SortTable<IPremiumApp>
            columns={appCols}
            rows={data.apps.filter(a => !q || `${a.name} ${a.envName} ${a.owner}`.toLowerCase().indexOf(q) >= 0)}
            rowKey={a => a.id}
            initial={{ key: 'unlicensed', desc: true }}
            empty="Keine Apps in Managed Environments."
          />
        )}
        <div className={styles.ccFoot}>
          Managed laut Spalte „{data.managedField}“ im CoE. Berechtigte: Besitzer, Mitbesitzer, Benutzer und Mitglieder freigegebener Gruppen (inkl. verschachtelter Gruppen).
          Lizenzen live über Microsoft Graph; als ausreichend gelten Power Apps Premium, Power Automate Premium und Dynamics 365 Enterprise mit Power-Apps-Rechten.
          Per-App-Pässe am Environment werden nicht erkannt; Pay-as-you-go (Power Apps) nur, wenn die Power Platform API freigegeben ist.
          Microsofts eigene Liste: Admin Center → Lizenzierung → Power Apps → „Bericht herunterladen“ → „Users requiring licenses in Managed Environments“.
        </div>
      </section>
    </div>
  );
};
