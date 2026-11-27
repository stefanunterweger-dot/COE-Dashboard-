import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton } from '@fluentui/react/lib/Button';
import styles from './CoeCockpit.module.scss';
import { IColumn, SortTable } from './CopilotCreditsView';
import { gb, IEnvStorage, IStorageData, loadStorage, STORAGE_LABEL, STORAGE_TYPES, StorageType } from '../models/storage';
import { DataverseService } from '../services/DataverseService';
import { PIM_URL, PowerPlatformApiService } from '../services/PowerPlatformApiService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * „Dataverse-Speicher“ (Lizenzen & Kapazität): Kontingent des Tenants (Datenbank, Datei, Log) live aus der
 * Power Platform API, Verbrauch je Environment aus dem CoE – mit Ampel ab 80 % bzw. 100 %.
 */
export interface IStorageViewProps {
  service: DataverseService;
  ppApi: PowerPlatformApiService | undefined;
  reloadKey: string;
  onOpenEnvironment: (id: string) => void;
}

const ICON: { [t in StorageType]: string } = { Database: 'Database', File: 'Page', Log: 'History' };

export const StorageView: React.FC<IStorageViewProps> = props => {
  const [data, setData] = React.useState<IStorageData | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [search, setSearch] = React.useState<string>('');
  const [type, setType] = React.useState<string>('');

  const load = (): void => {
    setLoading(true);
    loadStorage(props.service, props.ppApi).then(d => { setData(d); setLoading(false); }, () => setLoading(false));
  };
  React.useEffect(() => { load(); }, [props.service, props.ppApi, props.reloadKey]);

  const envs = data ? data.envs : [];
  const types = envs.map(e => e.type).filter((t, i, a) => !!t && a.indexOf(t) === i).sort();
  const q = search.trim().toLowerCase();
  const visible = envs.filter(e => (!type || e.type === type) && (!q || e.name.toLowerCase().indexOf(q) >= 0));
  const sumEnv = (t: StorageType): number => envs.reduce((s, e) => s + (t === 'Database' ? e.db : t === 'File' ? e.file : e.log), 0);
  const tenantDb = data ? data.tenant.filter(t => t.type === 'Database')[0] : undefined;

  const exportCsv = (): void => {
    const lines = ['Environment;Typ;Datenbank (GB);Datei (GB);Log (GB);Gesamt (GB)'].concat(visible.map(e => [
      e.name, e.type, gb(e.db, 2), gb(e.file, 2), gb(e.log, 2), gb(e.db + e.file + e.log, 2)
    ].map(v => (/[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(';')));
    downloadText(`Dataverse-Speicher_${new Date().toISOString().slice(0, 10)}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };

  const bar = (v: number, max: number): JSX.Element => (
    <span className={styles.stBar} aria-hidden="true"><span style={{ width: `${max ? Math.min(100, (v / max) * 100) : 0}%` }} /></span>
  );
  const maxDb = Math.max(1, ...envs.map(e => e.db));
  const cols: IColumn<IEnvStorage>[] = [
    { key: 'name', label: 'Environment', sort: e => e.name.toLowerCase(), render: e => <button type="button" className={styles.linkButton} onClick={() => props.onOpenEnvironment(e.id)}>{e.name}</button> },
    { key: 'type', label: 'Typ', sort: e => e.type, render: e => e.type || <span className={styles.muted}>–</span> },
    { key: 'db', label: 'Datenbank (GB)', numeric: true, sort: e => e.db, render: e => <>{bar(e.db, maxDb)}{gb(e.db, 2)}</> },
    { key: 'file', label: 'Datei (GB)', numeric: true, sort: e => e.file, render: e => gb(e.file, 2) },
    { key: 'log', label: 'Log (GB)', numeric: true, sort: e => e.log, render: e => gb(e.log, 2) },
    { key: 'sum', label: 'Gesamt (GB)', numeric: true, sort: e => e.db + e.file + e.log, render: e => <b>{gb(e.db + e.file + e.log, 2)}</b> },
    {
      key: 'share', label: 'Anteil Datenbank', numeric: true, sort: e => e.db,
      render: e => (tenantDb && tenantDb.total ? `${formatNumber(Math.round((e.db / tenantDb.total) * 1000) / 10)} %` : <span className={styles.muted}>–</span>)
    }
  ];

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="Database" /> Kontingent live (Power Platform API) · Verbrauch aus dem CoE</div>
          <div className={styles.secHeroTitle}>Dataverse-Speicher</div>
          <div className={styles.ovHeroLabel}>
            Wie viel Datenbank-, Datei- und Log-Kapazität der Tenant hat und welche Environments sie verbrauchen. Über dem Kontingent sperrt Microsoft
            u. a. das Anlegen und Kopieren von Environments.
          </div>
          <div className={styles.ovHeroChips}>
            <button type="button" className={styles.ovHeroChip} onClick={load} disabled={loading}><Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}</button>
            {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
          </div>
        </div>
      </section>
      {loading && !data && <div className={styles.ovCard}><div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>Kapazität wird gelesen …</span></div></div>}
      {data && (
        <>
          <div className={styles.stGrid}>
            {STORAGE_TYPES.map(t => {
              const c = data.tenant.filter(x => x.type === t)[0];
              const used = c && c.used !== undefined ? c.used : sumEnv(t);
              const total = c ? c.total : undefined;
              const pct = total ? (used / total) * 100 : undefined;
              return (
                <div key={t} className={styles.stCard} data-storage={t}>
                  <div className={styles.stTitle}><Icon iconName={ICON[t]} /> {STORAGE_LABEL[t]}</div>
                  <div className={styles.stValue}>{gb(used)} GB <span>{total !== undefined ? `von ${gb(total)} GB` : 'verbraucht'}</span></div>
                  {pct !== undefined && (
                    <div className={`${styles.stMeter} ${pct >= 100 ? styles.stMeterFull : pct >= 80 ? styles.stMeterWarn : ''}`} role="img" aria-label={`${Math.round(pct)} % belegt`}>
                      <span style={{ width: `${Math.min(100, pct)}%` }} />
                    </div>
                  )}
                  <div className={styles.stHint}>
                    {pct !== undefined ? `${formatNumber(Math.round(pct))} % belegt` : 'Kontingent nicht lesbar'}
                    {total !== undefined && used < total ? ` · frei ${gb(total - used)} GB` : ''}
                    {total !== undefined && used > total ? ` · ${gb(used - total)} GB darüber` : ''}
                    {c && c.overflow ? ` · ${gb(c.overflow)} GB über andere Kapazität gedeckt` : ''}
                  </div>
                </div>
              );
            })}
          </div>
          {(data.tenantStatus || data.tenantProblem || data.envProblem) && (
            <div className={styles.ccWarnings}>
              {data.tenantStatus && <div className={styles.muted}><Icon iconName="Info" /> Microsoft: {data.tenantStatus}</div>}
              {data.tenantProblem && (
                <div className={styles.muted}>
                  <Icon iconName="Info" /> Kontingent des Tenants nicht lesbar – angezeigt wird nur der Verbrauch laut CoE. {data.tenantProblem}
                  {data.tenantProblemKind === 'forbidden' && <> <a className={styles.link} href={PIM_URL} target="_blank" rel="noopener noreferrer">Rolle in PIM aktivieren</a></>}
                </div>
              )}
              {data.envProblem && <div className={styles.muted}><Icon iconName="Info" /> {data.envProblem}</div>}
            </div>
          )}
          <section className={styles.ovCard}>
            <div className={styles.toolbar}>
              <SearchBox placeholder="Environment suchen" value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 240, borderRadius: 6 } }} />
              <div className={styles.ccSegments} role="group" aria-label="Typ">
                <button type="button" className={`${styles.ccSegment} ${!type ? styles.ccSegmentActive : ''}`} aria-pressed={!type} onClick={() => setType('')}>Alle <span className={styles.ccSegmentCount}>{formatNumber(envs.length)}</span></button>
                {types.map(t => (
                  <button key={t} type="button" className={`${styles.ccSegment} ${type === t ? styles.ccSegmentActive : ''}`} aria-pressed={type === t} onClick={() => setType(type === t ? '' : t)}>
                    {t} <span className={styles.ccSegmentCount}>{formatNumber(envs.filter(e => e.type === t).length)}</span>
                  </button>
                ))}
              </div>
              <span className={styles.resultBadge} role="status"><strong>{formatNumber(visible.length)}</strong> Environments · {gb(visible.reduce((s, e) => s + e.db + e.file + e.log, 0))} GB</span>
              <span className={styles.grow} />
              <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportCsv} disabled={!visible.length} />
            </div>
            <SortTable<IEnvStorage>
              columns={cols}
              rows={visible}
              rowKey={e => e.id}
              initial={{ key: 'sum', desc: true }}
              empty={envs.length ? 'Keine Environments für diesen Filter.' : 'Keine Kapazitätsdaten im CoE.'}
            />
            <div className={styles.ccFoot}>Verbrauch je Environment laut CoE (letzter Sync); das Kontingent des Tenants live aus der Power Platform API.</div>
          </section>
        </>
      )}
    </div>
  );
};
