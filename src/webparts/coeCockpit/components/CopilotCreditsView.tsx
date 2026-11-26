import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { Tone } from '../models/ITableConfig';
import { PIM_URL, PowerPlatformApiService, PpApiError } from '../services/PowerPlatformApiService';
import {
  BILLING_KEY, CREDITS_CACHE_MS, CreditStatus, IBillingPlan, ICreditRow, ICreditsData, loadCopilotCredits, planStatusLabel, STATUS_HINT, STATUS_LABEL
} from '../models/copilotCredits';
import { formatDate, formatNumber } from '../utils/format';

/**
 * Ansichten „Copilot-Credits“ und „Abrechnungspläne“ – live aus der Power Platform API (kein CoE, kein Dataverse).
 * Der angemeldete Benutzer braucht die aktive Rolle „Power Platform Administrator“.
 */
export type CreditsMode = 'credits' | 'billing';

export interface ICopilotCreditsProps {
  api: PowerPlatformApiService;
  mode: CreditsMode;
  /** Ändert sich beim „Aktualisieren“ des Cockpits → Daten neu laden. */
  reloadKey: string;
  onSwitch: (mode: CreditsMode) => void;
}

// Geteilter Zwischenspeicher: beim Wechsel zwischen den beiden Ansichten nicht neu laden.
interface IShared { data?: ICreditsData; at: number; reloadKey?: string; pending?: Promise<ICreditsData>; pendingAll?: boolean; progress?: { done: number; total: number } }
const shared: IShared = { at: 0 };
const listeners: Array<() => void> = [];
function notify(): void {
  listeners.slice().forEach(l => l());
}

/** Nur für Tests/Vorschau: Zwischenspeicher leeren. */
export function resetCreditsCache(): void {
  shared.data = undefined;
  shared.at = 0;
  shared.pending = undefined;
  shared.progress = undefined;
}

function load(api: PowerPlatformApiService, all: boolean, reloadKey: string): Promise<ICreditsData> {
  if (shared.pending && (shared.pendingAll || !all)) {
    return shared.pending;
  }
  shared.progress = undefined;
  shared.pendingAll = all;
  const p = loadCopilotCredits(api, {
    all,
    onProgress: (done, total) => { shared.progress = { done, total }; notify(); }
  }).then(d => {
    if (shared.pending === p) {
      shared.data = d;
      shared.at = Date.now();
      shared.reloadKey = reloadKey;
      shared.pending = undefined;
      notify();
    }
    return d;
  }, (e: PpApiError) => {
    if (shared.pending === p) {
      shared.pending = undefined;
      notify();
    }
    throw e;
  });
  shared.pending = p;
  notify();
  return p;
}

const STATUS_TONE: { [k in CreditStatus]: Tone } = { within: 'success', overage: 'danger', covered: 'warning', '': 'neutral' };
const PPAC_ENV = (id: string): string => `https://admin.powerplatform.microsoft.com/environments/${encodeURIComponent(id)}/hub`;
const AZURE_RG = (sub: string, rg: string): string => `https://portal.azure.com/#resource/subscriptions/${encodeURIComponent(sub)}/resourceGroups/${encodeURIComponent(rg)}/overview`;

function n(v: number | undefined): string {
  return v === undefined ? '–' : formatNumber(Math.round(v * 100) / 100);
}

function yesNo(v: boolean | undefined): string {
  return v === undefined ? '–' : v ? 'Ja' : 'Nein';
}

type CreditFilter = 'all' | 'allocated' | 'overage' | 'paygo' | 'none';
const FILTERS: { key: CreditFilter; label: string; test: (r: ICreditRow) => boolean }[] = [
  { key: 'all', label: 'Alle', test: () => true },
  { key: 'allocated', label: 'Mit Zuweisung', test: r => r.allocated !== undefined },
  { key: 'overage', label: 'Über Kapazität', test: r => r.status === 'overage' || r.status === 'covered' },
  { key: 'paygo', label: 'Pay-as-you-go', test: r => r.payGo },
  { key: 'none', label: 'Ohne Zuweisung', test: r => r.allocated === undefined }
];

interface IColumn<T> { key: string; label: string; numeric?: boolean; sort: (r: T) => string | number; render: (r: T) => React.ReactNode; title?: string }

function SortTable<T>(props: { columns: IColumn<T>[]; rows: T[]; rowKey: (r: T) => string; initial: { key: string; desc: boolean }; empty: string; expanded?: (r: T) => React.ReactNode; onRow?: (r: T) => void }): JSX.Element {
  const [sort, setSort] = React.useState(props.initial);
  const col = props.columns.filter(c => c.key === sort.key)[0] || props.columns[0];
  const rows = props.rows.slice().sort((a, b) => {
    const x = col.sort(a);
    const y = col.sort(b);
    const r = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'de');
    return sort.desc ? -r : r;
  });
  return (
    <div className={styles.tableWrap}>
      <table className={`${styles.table} ${styles.ccTable}`}>
        <thead>
          <tr>
            {props.columns.map(c => {
              const sorted = c.key === sort.key;
              return (
                <th key={c.key} className={styles.th} scope="col" aria-sort={sorted ? (sort.desc ? 'descending' : 'ascending') : 'none'} title={c.title}>
                  <div className={styles.thInner}>
                    <button
                      type="button"
                      className={`${styles.thButton} ${sorted ? styles.thSorted : ''}`}
                      onClick={() => setSort({ key: c.key, desc: sorted ? !sort.desc : !!c.numeric })}
                    >
                      <span className={styles.ellipsis}>{c.label}</span>
                      <Icon className={styles.sortIcon} iconName={sorted && sort.desc ? 'SortDown' : 'SortUp'} />
                    </button>
                  </div>
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && (
            <tr><td className={styles.td} colSpan={props.columns.length}><span className={styles.muted}>{props.empty}</span></td></tr>
          )}
          {rows.map(r => {
            const extra = props.expanded ? props.expanded(r) : undefined;
            return (
              <React.Fragment key={props.rowKey(r)}>
                <tr className={props.onRow ? styles.ccRowClickable : undefined} onClick={props.onRow ? () => props.onRow!(r) : undefined}>
                  {props.columns.map(c => <td key={c.key} className={`${styles.td} ${c.numeric ? styles.ccNum : ''}`}>{c.render(r)}</td>)}
                </tr>
                {extra && <tr className={styles.ccExpanded}><td className={styles.td} colSpan={props.columns.length}>{extra}</td></tr>}
              </React.Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

const Tile: React.FC<{ label: string; hint: string; icon: string; tone: string; value: React.ReactNode; active?: boolean; onClick?: () => void }> = p => (
  <button
    type="button"
    className={`${styles.secTile} ${p.active ? styles.ccTileActive : ''}`}
    disabled={!p.onClick}
    onClick={p.onClick}
    aria-pressed={p.onClick ? !!p.active : undefined}
  >
    <span className={`${styles.govIcon} ${p.tone}`}><Icon iconName={p.icon} /></span>
    <span className={styles.govText}>
      <span className={styles.govLabel}>{p.label}</span>
      <span className={styles.govHint}>{p.hint}</span>
    </span>
    <span className={styles.govValue}>{p.value}</span>
  </button>
);

export const CopilotCreditsView: React.FC<ICopilotCreditsProps> = props => {
  const { api, mode } = props;
  const [, setTick] = React.useState(0);
  const [error, setError] = React.useState<PpApiError | undefined>(undefined);
  const [filter, setFilter] = React.useState<CreditFilter>('all');
  const [search, setSearch] = React.useState('');
  const [openPlan, setOpenPlan] = React.useState<string>('');

  React.useEffect(() => {
    const l = (): void => setTick(t => t + 1);
    listeners.push(l);
    return () => { listeners.splice(listeners.indexOf(l), 1); };
  }, []);

  const start = (all: boolean, force: boolean): void => {
    const fresh = shared.data && Date.now() - shared.at < CREDITS_CACHE_MS && shared.reloadKey === props.reloadKey;
    if (!force && fresh && (!all || shared.data!.allChecked)) {
      return;
    }
    setError(undefined);
    load(api, all || (!!shared.data && shared.data.allChecked && !force), props.reloadKey).catch((e: PpApiError) => setError(e));
  };

  React.useEffect(() => { start(false, false); }, [api, props.reloadKey]);

  const data = shared.data;
  const loading = !!shared.pending;
  const progress = shared.progress;
  const forbidden = error && (error.kind === 'forbidden' || error.kind === 'consent');

  const hero = (
    <section className={`${styles.ovHero} ${styles.secHero}`}>
      <div className={styles.ovHeroGlow} aria-hidden="true" />
      <div className={styles.ovHeroMain}>
        <div className={styles.ovHeroEyebrow}><Icon iconName={mode === 'credits' ? 'AllCurrency' : 'Bank'} /> Live aus der Power Platform API</div>
        <div className={styles.secHeroTitle}>{mode === 'credits' ? 'Copilot-Credits je Environment' : 'Abrechnungspläne (Pay-as-you-go)'}</div>
        <div className={styles.ovHeroLabel}>
          {mode === 'credits'
            ? 'Zuweisung, Verbrauch und Pay-as-you-go für Copilot Studio – direkt vom Dienst, ohne CoE und Dataverse.'
            : 'Abrechnungspläne verrechnen Verbrauch über ein Azure-Abonnement. Hier siehst du alle Pläne und die verknüpften Environments.'}
          {' '}Benötigt die aktive Rolle „Power Platform Administrator“.
        </div>
        <div className={styles.ovHeroChips}>
          <button type="button" className={styles.ovHeroChip} onClick={() => start(false, true)} disabled={loading}>
            <Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}
          </button>
          <button type="button" className={styles.ovHeroChip} onClick={() => props.onSwitch(mode === 'credits' ? 'billing' : 'credits')}>
            <Icon iconName={mode === 'credits' ? 'Bank' : 'AllCurrency'} /> {mode === 'credits' ? 'Abrechnungspläne' : 'Copilot-Credits'}
          </button>
          {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
        </div>
      </div>
    </section>
  );

  const notice = error && (
    <div className={styles.accessNotice} role="alert">
      <span className={styles.accessIcon}><Icon iconName={forbidden ? 'Lock' : 'Warning'} /></span>
      <div className={styles.ccNoticeText}>
        <strong>{error.kind === 'consent' ? 'Power Platform API noch nicht freigegeben' : error.kind === 'forbidden' ? 'Keine Berechtigung' : 'Daten konnten nicht geladen werden'}</strong>
        <span>{error.message}</span>
        {error.detail && error.detail !== error.message && <span className={styles.ccDetail}>Technisch: {error.detail}</span>}
      </div>
      <div className={styles.accessActions}>
        {error.kind === 'forbidden' && <a className={styles.link} href={PIM_URL} target="_blank" rel="noopener noreferrer"><Icon iconName="OpenInNewWindow" /> Rolle in PIM aktivieren</a>}
        <button type="button" className={styles.textButton} onClick={() => start(false, true)}>Erneut versuchen</button>
      </div>
    </div>
  );

  const loadingBlock = loading && !data && (
    <div className={styles.ovCard}>
      <div className={styles.ccLoading}>
        <span className={styles.skeleton} style={{ width: '40%', height: 14 }} />
        <span className={styles.muted}>
          {progress && progress.total ? `Environments werden abgefragt: ${formatNumber(progress.done)} von ${formatNumber(progress.total)}` : 'Lizenzdaten werden abgefragt …'}
        </span>
      </div>
    </div>
  );

  const warnings = data && data.warnings.length > 0 && (
    <div className={styles.ccWarnings}>
      {data.warnings.map((w, i) => <div key={i} className={styles.muted}><Icon iconName="Info" /> {w}</div>)}
    </div>
  );

  if (mode === 'billing') {
    const plans = data ? data.plans : [];
    const linked = plans.reduce((s, p) => s + p.environments.length, 0);
    const cols: IColumn<IBillingPlan>[] = [
      { key: 'name', label: 'Name', sort: p => p.name.toLowerCase(), render: p => <span className={styles.titleCell}><Icon iconName={openPlan === p.id ? 'ChevronDown' : 'ChevronRight'} /> {p.name}</span> },
      { key: 'status', label: 'Status', sort: p => planStatusLabel(p.status), render: p => <Pill tone={/^enabled$/i.test(p.status) ? 'success' : 'neutral'}>{planStatusLabel(p.status)}</Pill> },
      { key: 'envs', label: 'Environments', numeric: true, sort: p => p.environments.length, render: p => formatNumber(p.environments.length) },
      { key: 'products', label: 'Produkte', sort: p => p.products.join(', '), render: p => <span className={styles.ccName} title={p.products.join(', ')}>{p.products.length ? p.products.join(', ') : '–'}</span> },
      { key: 'sub', label: 'Azure-Abonnement', sort: p => p.subscriptionId, render: p => <span className={styles.mono} title={p.subscriptionId}>{p.subscriptionId ? `${p.subscriptionId.slice(0, 8)}…` : '–'}</span> },
      {
        key: 'rg', label: 'Ressourcengruppe', sort: p => p.resourceGroup.toLowerCase(),
        render: p => p.subscriptionId && p.resourceGroup
          ? <a className={styles.link} href={AZURE_RG(p.subscriptionId, p.resourceGroup)} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()}>{p.resourceGroup} <Icon iconName="OpenInNewWindow" /></a>
          : '–'
      },
      { key: 'loc', label: 'Region', sort: p => p.location, render: p => p.location || '–' },
      { key: 'created', label: 'Erstellt am', sort: p => p.createdOn, render: p => p.createdOn ? formatDate(new Date(p.createdOn)) : '–' }
    ];
    const rowByEnv: { [id: string]: ICreditRow } = {};
    (data ? data.rows : []).forEach(r => { rowByEnv[r.environmentId] = r; });
    return (
      <div className={styles.ov}>
        {hero}
        {notice}
        {loadingBlock}
        {data && (
          <>
            <div className={styles.ccGrid}>
              <Tile label="Abrechnungspläne" hint="im Tenant" icon="Bank" tone={styles.ovToneInfo} value={formatNumber(plans.length)} />
              <Tile label="Aktiv" hint="Status „Aktiv“" icon="CheckMark" tone={styles.ovToneInfo} value={formatNumber(plans.filter(p => /^enabled$/i.test(p.status)).length)} />
              <Tile label="Verknüpfte Environments" hint="verrechnen über einen Plan" icon="Globe" tone={styles.ovToneInfo} value={formatNumber(linked)} />
              <Tile label="Mit Copilot Studio" hint="Pläne mit Copilot-Studio-Abrechnung" icon="AllCurrency" tone={styles.ovToneWarning} value={formatNumber(plans.filter(p => p.copilot).length)} />
            </div>
            {warnings}
            <section className={styles.ovCard}>
              {data.plansError
                ? <div className={styles.ovFailed}><Icon iconName="Info" /> {data.plansError}</div>
                : (
                  <SortTable<IBillingPlan>
                    columns={cols}
                    rows={plans}
                    rowKey={p => p.id}
                    initial={{ key: 'name', desc: false }}
                    empty="Keine Abrechnungspläne im Tenant."
                    onRow={p => setOpenPlan(openPlan === p.id ? '' : p.id)}
                    expanded={p => openPlan === p.id && (
                      p.environments.length === 0
                        ? <span className={styles.muted}>Mit diesem Plan ist kein Environment verknüpft.</span>
                        : (
                          <ul className={styles.ccEnvList}>
                            {p.environments.map(e => {
                              const r = rowByEnv[e.id];
                              return (
                                <li key={e.id}>
                                  <a className={styles.link} href={PPAC_ENV(e.id)} target="_blank" rel="noopener noreferrer" onClick={ev => ev.stopPropagation()}>{e.name}</a>
                                  {r && r.payGoConsumed !== undefined && <span className={styles.muted}> · Copilot-Credits Pay-as-you-go: {n(r.payGoConsumed)}</span>}
                                </li>
                              );
                            })}
                          </ul>
                        )
                    )}
                  />
                )}
            </section>
          </>
        )}
      </div>
    );
  }

  // ---------- Copilot-Credits ----------
  const rows = data ? data.rows : [];
  const s = data ? data.summary : {};
  const overage = rows.filter(FILTERS[2].test).length;
  const q = search.trim().toLowerCase();
  const f = FILTERS.filter(x => x.key === filter)[0];
  const narrowed = filter !== 'all' || !!q;
  const visible = rows.filter(r => f.test(r) && (!q || r.name.toLowerCase().indexOf(q) >= 0 || r.environmentId.indexOf(q) >= 0 || (r.payGoPlan || '').toLowerCase().indexOf(q) >= 0));
  const pct = (a?: number, b?: number): string => a !== undefined && b ? ` · ${formatNumber(Math.round((a / b) * 100))} %` : '';
  const copilotPlans = data ? data.plans.filter(p => p.copilot).length : 0;
  const cols: IColumn<ICreditRow>[] = [
    {
      key: 'name', label: 'Environment', sort: r => r.name.toLowerCase(),
      render: r => <a className={`${styles.link} ${styles.ccName}`} href={PPAC_ENV(r.environmentId)} target="_blank" rel="noopener noreferrer" title={`${r.name}\n${r.environmentId}`}>{r.name}</a>
    },
    { key: 'status', label: 'Status', sort: r => STATUS_LABEL[r.status], render: r => r.status ? <Pill tone={STATUS_TONE[r.status]} title={STATUS_HINT[r.status]}>{STATUS_LABEL[r.status]}</Pill> : <span className={styles.empty}>–</span> },
    { key: 'allocated', label: 'Zugewiesen', numeric: true, sort: r => r.allocated || 0, render: r => n(r.allocated) },
    {
      key: 'consumed', label: 'Prepaid verbraucht', numeric: true, sort: r => r.consumed || 0,
      render: r => {
        const share = r.allocated && r.consumed !== undefined ? Math.min(100, (r.consumed / r.allocated) * 100) : undefined;
        return (
          <span className={styles.ccConsumed}>
            <span>{n(r.consumed)}</span>
            {share !== undefined && (
              <span className={styles.ccMeter} title={`${formatNumber(Math.round(share))} % der Zuweisung`}>
                <span className={`${styles.ccMeterFill} ${share >= 100 ? styles.ccMeterFull : share >= 80 ? styles.ccMeterHigh : ''}`} style={{ width: `${share}%` }} />
              </span>
            )}
          </span>
        );
      }
    },
    { key: 'pool', label: 'Tenant-Kapazität', sort: r => yesNo(r.tenantPool), render: r => yesNo(r.tenantPool), title: '„Draw from the available capacity in my tenant“ – nach Verbrauch der Zuweisung weiter aus der freien Tenant-Kapazität' },
    {
      key: 'paygo', label: 'Pay-as-you-go', sort: r => (r.payGo ? `1${r.payGoPlan}` : r.payGoPlan ? `0${r.payGoPlan}` : '0'),
      render: r => r.payGo
        ? <span className={styles.ccName} title={`Abrechnungsplan: ${r.payGoPlan}`}>Ja · {r.payGoPlan}</span>
        : r.payGoPlan ? <span className={styles.muted} title={`Abrechnungsplan „${r.payGoPlan}“ verrechnet kein Copilot Studio`}>Nein (Plan ohne Copilot)</span> : 'Nein'
    },
    { key: 'paygoConsumed', label: 'PAYG verbraucht', title: 'Pay-as-you-go verbraucht', numeric: true, sort: r => (r.payGo ? r.payGoConsumed || 0 : -1), render: r => n(r.payGo ? r.payGoConsumed : undefined) },
    {
      key: 'type', label: 'Typ', sort: r => `${r.type}${r.managed ? '1' : '0'}`,
      render: r => <span title={r.managed === undefined ? undefined : r.managed ? 'Managed Environment' : 'Kein Managed Environment'}>{r.type || '–'}{r.managed ? ' · Managed' : ''}</span>
    }
  ];

  return (
    <div className={styles.ov}>
      {hero}
      {notice}
      {loadingBlock}
      {data && (
        <>
          <div className={styles.ccGrid}>
            <Tile label="Gekauft (Prepaid)" hint="Kapazität im Tenant" icon="AllCurrency" tone={styles.ovToneInfo} value={n(s.purchased)} />
            <Tile label="Zugewiesen" hint={`an ${formatNumber(rows.filter(r => r.allocated !== undefined).length)} Environments`} icon="Org" tone={styles.ovToneInfo} value={n(s.allocated)} />
            <Tile label="Noch zuweisbar" hint="gekauft minus zugewiesen" icon="Add" tone={s.assignable === 0 ? styles.ovToneWarning : styles.ovToneInfo} value={n(s.assignable)} />
            <Tile label="Prepaid verbraucht" hint={`aktueller Zeitraum${pct(s.consumed, s.purchased)}`} icon="BarChartVertical" tone={styles.ovToneInfo} value={n(s.consumed)} />
            <Tile label="Pay-as-you-go verbraucht" hint={`über ${formatNumber(copilotPlans)} ${copilotPlans === 1 ? 'Abrechnungsplan' : 'Abrechnungspläne'} (Azure)`} icon="Bank" tone={styles.ovToneSerious} value={n(s.payGoConsumed)} onClick={() => setFilter(filter === 'paygo' ? 'all' : 'paygo')} active={filter === 'paygo'} />
            <Tile label="Über Kapazität" hint="Zuweisung aufgebraucht" icon="Warning" tone={overage ? styles.ovToneCritical : styles.ovToneInfo} value={formatNumber(overage)} onClick={() => setFilter(filter === 'overage' ? 'all' : 'overage')} active={filter === 'overage'} />
          </div>
          {warnings}
          <section className={styles.ovCard}>
            <div className={styles.toolbar}>
              <SearchBox
                placeholder="Environment oder Abrechnungsplan suchen"
                value={search}
                onChange={(_e, v) => setSearch(v || '')}
                styles={{ root: { width: 280, borderRadius: 6 } }}
              />
              <span className={`${styles.resultBadge} ${narrowed ? styles.resultBadgeActive : ''}`} role="status" aria-live="polite">
                {narrowed && <Icon iconName="Filter" />}
                {narrowed
                  ? <><strong>{formatNumber(visible.length)}</strong> von {formatNumber(rows.length)} Environments</>
                  : <><strong>{formatNumber(rows.length)}</strong> {rows.length === 1 ? 'Environment' : 'Environments'}</>}
              </span>
              <div className={styles.ccSegments} role="group" aria-label="Filter">
                {FILTERS.map(x => (
                  <button
                    key={x.key}
                    type="button"
                    className={`${styles.ccSegment} ${filter === x.key ? styles.ccSegmentActive : ''}`}
                    aria-pressed={filter === x.key}
                    onClick={() => setFilter(x.key)}
                  >
                    {x.label} <span className={styles.ccSegmentCount}>{formatNumber(rows.filter(x.test).length)}</span>
                  </button>
                ))}
              </div>
              <span className={styles.grow} />
              <span className={styles.resultInfo}>
                {loading && progress && progress.total ? `Prüfe ${formatNumber(progress.done)} von ${formatNumber(progress.total)} …` : data.allChecked
                  ? `Alle ${formatNumber(data.envTotal || rows.length)} Environments geprüft`
                  : `${formatNumber(rows.length)} Environments mit Zuweisung, Verbrauch oder Abrechnungsplan`}
              </span>
              {!data.allChecked && (
                <button type="button" className={styles.textButton} onClick={() => start(true, false)} disabled={loading} title="Fragt jedes Environment des Tenants einzeln ab – dauert bei vielen Environments etwas.">
                  Alle Environments prüfen
                </button>
              )}
            </div>
            <SortTable<ICreditRow>
              columns={cols}
              rows={visible}
              rowKey={r => r.environmentId}
              initial={{ key: 'allocated', desc: true }}
              empty={rows.length ? 'Keine Environments für diesen Filter.' : 'Keine Environments mit Copilot-Credits gefunden.'}
            />
            <div className={styles.ccFoot}>
              Werte laut Power Platform API{rows.some(r => r.updatedOn) ? ` (Verbrauch aktualisiert ${formatDate(new Date(rows.filter(r => r.updatedOn).map(r => r.updatedOn!).sort().reverse()[0]), true)})` : ''}.
              Zuweisen und Überziehungsregeln änderst du im Power Platform Admin Center unter Lizenzierung → Copilot Studio.
            </div>
          </section>
        </>
      )}
    </div>
  );
};

export const CREDITS_NAV = { credits: 'copilot-credits', billing: BILLING_KEY };
