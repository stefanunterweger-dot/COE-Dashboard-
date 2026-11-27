import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable, Tile } from './CopilotCreditsView';
import { IOptData, IOptRow, loadLicenseOptimization, PREMIUM_SKUS } from '../models/licenseOptimization';
import { LicenseLoadError } from '../models/userLicenses';
import { DataverseService } from '../services/DataverseService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * „Lizenz-Optimierung“ (Lizenzen & Kapazität): Premium-Lizenzen ohne erkennbare Nutzung – Kandidaten zum Einsparen.
 * Lizenzen live aus Microsoft Graph, Nutzung aus dem CoE (App-Starts im Audit-Log, eingeschaltete Flows).
 */
export interface ILicenseOptimizationViewProps {
  service: DataverseService;
  graphFactory: MSGraphClientFactory;
  reloadKey: string;
}

const PRICE_KEY = 'coe-cockpit:license-prices';
const STATUS_LABEL: { [s in IOptRow['status']]: string } = { unused: 'Ungenutzt', partial: 'Teilweise genutzt', used: 'Genutzt', unknown: 'Nicht prüfbar' };
const STATUS_TONE: { [s in IOptRow['status']]: 'danger' | 'warning' | 'success' | 'neutral' } = { unused: 'danger', partial: 'warning', used: 'success', unknown: 'neutral' };
const PAGE = 300;

function readPrices(): { [part: string]: number } {
  try {
    return JSON.parse(window.localStorage.getItem(PRICE_KEY) || '{}') as { [part: string]: number };
  } catch {
    return {};
  }
}

export const LicenseOptimizationView: React.FC<ILicenseOptimizationViewProps> = props => {
  const [days, setDays] = React.useState<number>(90);
  const [data, setData] = React.useState<IOptData | undefined>(undefined);
  const [error, setError] = React.useState<string>('');
  const [loading, setLoading] = React.useState<boolean>(true);
  const [search, setSearch] = React.useState<string>('');
  const [status, setStatus] = React.useState<'all' | IOptRow['status']>('all');
  const [prices, setPrices] = React.useState<{ [part: string]: number }>(readPrices);
  const [limit, setLimit] = React.useState<number>(PAGE);

  React.useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError('');
    loadLicenseOptimization(props.service, props.graphFactory, days).then(
      d => { if (!cancelled) { setData(d); setLoading(false); } },
      (e: unknown) => { if (!cancelled) { setError(e instanceof LicenseLoadError || e instanceof Error ? e.message : String(e)); setLoading(false); } }
    );
    return () => { cancelled = true; };
  }, [props.service, props.graphFactory, props.reloadKey, days]);
  React.useEffect(() => { setLimit(PAGE); }, [search, status, days]);

  const price = (part: string): number => (prices[part] !== undefined ? prices[part] : PREMIUM_SKUS[part] ? PREMIUM_SKUS[part].price : 0);
  const setPrice = (part: string, v: number): void => {
    const next = { ...prices, [part]: v };
    setPrices(next);
    try {
      window.localStorage.setItem(PRICE_KEY, JSON.stringify(next));
    } catch {
      // nur für diese Sitzung
    }
  };
  /** Einsparung je Person: Summe der ungenutzten Lizenzen */
  const saving = (r: IOptRow): number => r.licenses.filter(l => l.used === false).reduce((s, l) => s + price(l.part), 0);

  const rows = data ? data.rows : [];
  const q = search.trim().toLowerCase();
  const visible = rows.filter(r => (status === 'all' || r.status === status)
    && (!q || `${r.user.name} ${r.user.upn} ${r.user.department} ${r.licenses.map(l => l.name).join(' ')}`.toLowerCase().indexOf(q) >= 0));
  const assigned = rows.reduce((s, r) => s + r.licenses.length, 0);
  const unusedLicenses = rows.reduce((s, r) => s + r.licenses.filter(l => l.used === false).length, 0);
  const monthly = rows.reduce((s, r) => s + saving(r), 0);
  const count = (st: IOptRow['status']): number => rows.filter(r => r.status === st).length;
  const euro = (v: number): string => `${formatNumber(Math.round(v))} €`;

  const exportCsv = (): void => {
    const esc = (v: string): string => (/[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
    const lines = ['Name;Benutzername;Abteilung;Premium-Lizenzen;Ungenutzt;Premium-App-Starts;Premium-Apps;Zuletzt gestartet;Aktive Flows;Status;Einsparung €/Monat'].concat(visible.map(r => [
      r.user.name, r.user.upn, r.user.department, r.licenses.map(l => l.name).join(', '), r.licenses.filter(l => l.used === false).map(l => l.name).join(', '),
      String(r.premiumLaunches), String(r.premiumApps), r.lastLaunch ? formatDate(new Date(r.lastLaunch)) : '', String(r.activeFlows), STATUS_LABEL[r.status], String(saving(r))
    ].map(esc).join(';')));
    downloadText(`Lizenz-Optimierung_${new Date().toISOString().slice(0, 10)}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };

  const cols: IColumn<IOptRow>[] = [
    { key: 'name', label: 'Benutzer', sort: r => r.user.name.toLowerCase(), render: r => <span className={styles.ulUser}><b>{r.user.name}</b><span className={styles.ccSub}>{r.user.upn}{r.user.enabled ? '' : ' · deaktiviert'}</span></span> },
    { key: 'dept', label: 'Abteilung', sort: r => r.user.department.toLowerCase(), render: r => r.user.department || <span className={styles.muted}>–</span> },
    {
      key: 'lic', label: 'Premium-Lizenzen', sort: r => r.licenses.map(l => l.name).join(', '),
      render: r => (
        <span className={styles.ulChips}>
          {r.licenses.map(l => (
            <span key={l.id} className={`${styles.ulChip} ${l.used === false ? styles.loUnused : l.used ? styles.loUsed : ''}`}
              title={l.used === false ? 'Keine Nutzung erkannt' : l.used ? 'Genutzt' : 'Nicht prüfbar'}>{l.name}</span>
          ))}
        </span>
      )
    },
    { key: 'launches', label: 'Premium-App-Starts', numeric: true, sort: r => r.premiumLaunches, render: r => (r.premiumLaunches ? `${formatNumber(r.premiumLaunches)} (${formatNumber(r.premiumApps)} Apps)` : <span className={styles.muted}>0</span>) },
    { key: 'last', label: 'Zuletzt gestartet', sort: r => r.lastLaunch, render: r => (r.lastLaunch ? formatDate(new Date(r.lastLaunch)) : <span className={styles.muted}>nie</span>) },
    { key: 'flows', label: 'Aktive Flows', numeric: true, sort: r => r.activeFlows, render: r => (r.activeFlows ? formatNumber(r.activeFlows) : <span className={styles.muted}>0</span>) },
    { key: 'status', label: 'Status', sort: r => ['unused', 'partial', 'unknown', 'used'].indexOf(r.status), render: r => <Pill tone={STATUS_TONE[r.status]}>{STATUS_LABEL[r.status]}</Pill> },
    { key: 'saving', label: 'Einsparung / Monat', numeric: true, sort: r => saving(r), render: r => (saving(r) ? <b>{euro(saving(r))}</b> : <span className={styles.muted}>–</span>) }
  ];

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="Savings" /> Lizenzen live (Microsoft Graph) · Nutzung aus dem CoE</div>
          <div className={styles.secHeroTitle}>Lizenz-Optimierung</div>
          <div className={styles.ovHeroLabel}>
            Wer hat eine Power-Platform-Premium-Lizenz, nutzt sie aber nicht? Power Apps Premium gilt als genutzt, wenn die Person im Zeitraum eine
            Premium-App gestartet hat; Power Automate Premium, wenn ihr ein eingeschalteter Flow gehört. Kandidaten zum Nachfragen – keine sichere Aussage.
          </div>
          <div className={styles.ovHeroChips}>
            {[30, 60, 90, 180].map(d => (
              <button key={d} type="button" className={`${styles.ovHeroChip} ${days === d ? styles.loDayActive : ''}`} aria-pressed={days === d} onClick={() => setDays(d)} disabled={loading}>
                {d} Tage
              </button>
            ))}
            {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
          </div>
        </div>
      </section>
      {error && (
        <div className={styles.accessNotice} role="alert">
          <span className={styles.accessIcon}><Icon iconName="Lock" /></span>
          <div className={styles.ccNoticeText}><strong>Lizenzen konnten nicht geladen werden</strong><span>{error}</span></div>
        </div>
      )}
      {loading && !data && <div className={styles.ovCard}><div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>Lizenzen und Nutzung werden ausgewertet …</span></div></div>}
      {data && (
        <>
          <div className={styles.ccGrid}>
            <Tile label="Premium-Lizenzen" hint={`zugewiesen an ${formatNumber(rows.length)} Personen`} icon="ContactCard" tone={styles.ovToneInfo} value={formatNumber(assigned)} active={status === 'all'} onClick={() => setStatus('all')} />
            <Tile label="Ungenutzt" hint={`keine Nutzung in ${days} Tagen`} icon="Blocked2" tone={count('unused') ? styles.ovToneCritical : styles.ovToneInfo} value={formatNumber(count('unused'))}
              active={status === 'unused'} onClick={() => setStatus(status === 'unused' ? 'all' : 'unused')} />
            <Tile label="Teilweise genutzt" hint="eine von mehreren Lizenzen ungenutzt" icon="Warning" tone={count('partial') ? styles.ovToneWarning : styles.ovToneInfo} value={formatNumber(count('partial'))}
              active={status === 'partial'} onClick={() => setStatus(status === 'partial' ? 'all' : 'partial')} />
            <Tile label="Mögliche Einsparung" hint={`${formatNumber(unusedLicenses)} Lizenzen · ${euro(monthly * 12)} im Jahr`} icon="Savings" tone={monthly ? styles.ovToneSerious : styles.ovToneInfo} value={euro(monthly)} />
          </div>
          {data.problems.length > 0 && (
            <div className={styles.ccWarnings}>{data.problems.map((p, i) => <div key={i} className={styles.muted}><Icon iconName="Info" /> {p}</div>)}</div>
          )}
          {data.skus.length > 0 && (
            <div className={styles.loPrices}>
              <span className={styles.muted}>Listenpreis pro Monat (anpassbar, nur für diesen Browser):</span>
              {data.skus.map(s => (
                <label key={s.id} className={styles.loPrice}>
                  {s.name}
                  <input type="number" min={0} step={1} value={price(s.part)} onChange={e => setPrice(s.part, Math.max(0, Number(e.target.value) || 0))} aria-label={`Preis ${s.name}`} />
                  €
                </label>
              ))}
            </div>
          )}
          <section className={styles.ovCard}>
            <div className={styles.toolbar}>
              <SearchBox placeholder="Name, Benutzername, Abteilung oder Lizenz" value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 280, borderRadius: 6 } }} />
              <div className={styles.ccSegments} role="group" aria-label="Status">
                {(['all', 'unused', 'partial', 'used'] as ('all' | IOptRow['status'])[]).map(s => (
                  <button key={s} type="button" className={`${styles.ccSegment} ${status === s ? styles.ccSegmentActive : ''}`} aria-pressed={status === s} onClick={() => setStatus(s)}>
                    {s === 'all' ? 'Alle' : STATUS_LABEL[s]} <span className={styles.ccSegmentCount}>{formatNumber(s === 'all' ? rows.length : count(s))}</span>
                  </button>
                ))}
              </div>
              <span className={`${styles.resultBadge} ${visible.length !== rows.length ? styles.resultBadgeActive : ''}`} role="status" aria-live="polite">
                <strong>{formatNumber(visible.length)}</strong>{visible.length !== rows.length ? ` von ${formatNumber(rows.length)}` : ''} Personen
              </span>
              <span className={styles.grow} />
              <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportCsv} disabled={!visible.length} />
            </div>
            <SortTable<IOptRow>
              columns={cols}
              rows={visible}
              rowKey={r => r.user.id}
              initial={{ key: 'saving', desc: true }}
              empty={rows.length ? 'Keine Personen für diesen Filter.' : 'Niemand hat eine Power-Platform-Premium-Lizenz.'}
              limit={limit}
            />
            {visible.length > limit && (
              <div className={styles.ccFoot}><span>{formatNumber(limit)} von {formatNumber(visible.length)} angezeigt.</span>
                <button type="button" className={styles.textButton} onClick={() => setLimit(limit + PAGE)}>Weitere anzeigen</button></div>
            )}
            <div className={styles.ccFoot}>
              Nicht erkannt werden: Nutzung über Dynamics-365-Lizenzen, Pay-as-you-go und Pro-App-Pässe, Flows, die die Person nur nutzt (nicht besitzt),
              sowie Starts, die das CoE-Audit-Log nicht erfasst. Vor dem Entziehen einer Lizenz bitte bei der Person nachfragen.
            </div>
          </section>
        </>
      )}
    </div>
  );
};
