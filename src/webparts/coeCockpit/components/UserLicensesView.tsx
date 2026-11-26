import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable, Tile } from './CopilotCreditsView';
import { ILicUser, ISku, IUserLicenseData, LicenseLoadError, loadUserLicenses } from '../models/userLicenses';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * „Lizenzen je Benutzer“ (Navigation → Lizenzen & Kapazität): live aus Microsoft Graph.
 * Filter nach Lizenz (eine oder mehrere, „mindestens eine“ oder „alle“, auch „ohne Lizenz“), Suche, Kacheln als
 * Schnellfilter (ohne Lizenz, deaktiviert mit Lizenz, Gäste mit Lizenz), Lizenzübersicht mit gekauft/verbraucht, CSV-Export.
 */
export interface IUserLicensesViewProps {
  graphFactory: MSGraphClientFactory;
  reloadKey: string;
}

const NONE = '__none';
type Quick = 'all' | 'unlicensed' | 'disabledLicensed' | 'guestLicensed' | 'licensed';
const PAGE = 300;

export const UserLicensesView: React.FC<IUserLicensesViewProps> = props => {
  const [data, setData] = React.useState<IUserLicenseData | undefined>(undefined);
  const [error, setError] = React.useState<LicenseLoadError | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [progress, setProgress] = React.useState<number>(0);
  const [search, setSearch] = React.useState<string>('');
  const [selected, setSelected] = React.useState<string[]>([]);
  const [mode, setMode] = React.useState<'any' | 'all'>('any');
  const [quick, setQuick] = React.useState<Quick>('all');
  const [limit, setLimit] = React.useState<number>(PAGE);

  const start = (fresh: boolean): void => {
    setLoading(true);
    setError(undefined);
    setProgress(0);
    loadUserLicenses(props.graphFactory, n => setProgress(n), fresh).then(
      d => { setData(d); setLoading(false); },
      (e: unknown) => {
        setError(e instanceof LicenseLoadError ? e : new LicenseLoadError(e instanceof Error ? e.message : String(e), 'failed'));
        setLoading(false);
      }
    );
  };
  React.useEffect(() => { start(false); }, [props.graphFactory, props.reloadKey]);
  React.useEffect(() => { setLimit(PAGE); }, [search, selected.join('|'), mode, quick]);

  const users = data ? data.users : [];
  const skus: ISku[] = data ? Object.keys(data.skus).map(k => data.skus[k]).sort((a, b) => b.assigned - a.assigned || a.name.localeCompare(b.name, 'de')) : [];
  const skuName = (id: string): string => (data && data.skus[id] ? data.skus[id].name : id);

  const quickTest: { [q in Quick]: (u: ILicUser) => boolean } = {
    all: () => true,
    licensed: u => u.skus.length > 0,
    unlicensed: u => u.skus.length === 0,
    disabledLicensed: u => !u.enabled && u.skus.length > 0,
    guestLicensed: u => u.guest && u.skus.length > 0
  };

  const q = search.trim().toLowerCase();
  const visible = users.filter(u => {
    if (!quickTest[quick](u)) {
      return false;
    }
    if (selected.length) {
      const wantNone = selected.indexOf(NONE) >= 0;
      const lic = selected.filter(s => s !== NONE);
      const hasNone = u.skus.length === 0;
      const match = mode === 'all'
        ? (wantNone ? hasNone && lic.length === 0 : lic.every(s => u.skus.indexOf(s) >= 0))
        : (wantNone && hasNone) || lic.some(s => u.skus.indexOf(s) >= 0);
      if (!match) {
        return false;
      }
    }
    if (q) {
      const hay = `${u.name} ${u.upn} ${u.department} ${u.jobTitle} ${u.skus.map(skuName).join(' ')}`.toLowerCase();
      if (hay.indexOf(q) < 0) {
        return false;
      }
    }
    return true;
  });
  const narrowed = visible.length !== users.length;

  const toggleSku = (id: string): void => {
    setSelected(prev => (prev.indexOf(id) >= 0 ? prev.filter(x => x !== id) : prev.concat([id])));
  };

  const exportCsv = (): void => {
    const esc = (v: string): string => (/[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
    const lines = ['Name;Benutzername;Abteilung;Position;Status;Typ;Anzahl Lizenzen;Lizenzen'].concat(visible.map(u => [
      u.name, u.upn, u.department, u.jobTitle, u.enabled ? 'Aktiv' : 'Deaktiviert', u.guest ? 'Gast' : 'Mitglied', String(u.skus.length),
      u.skus.map(skuName).sort().join(', ')
    ].map(esc).join(';')));
    const stamp = new Date().toISOString().slice(0, 10);
    downloadText(`Lizenzen je Benutzer_${stamp}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };

  const hero = (
    <section className={`${styles.ovHero} ${styles.secHero}`}>
      <div className={styles.ovHeroGlow} aria-hidden="true" />
      <div className={styles.ovHeroMain}>
        <div className={styles.ovHeroEyebrow}><Icon iconName="ContactCard" /> Live aus Microsoft Graph</div>
        <div className={styles.secHeroTitle}>Lizenzen je Benutzer</div>
        <div className={styles.ovHeroLabel}>
          Welche Lizenzen jede Person hat – filterbar nach Lizenz, Abteilung und Status. Die Daten kommen direkt aus Entra ID, nicht aus dem CoE.
        </div>
        <div className={styles.ovHeroChips}>
          <button type="button" className={styles.ovHeroChip} onClick={() => start(true)} disabled={loading}>
            <Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}
          </button>
          {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
        </div>
      </div>
    </section>
  );

  const notice = error && (
    <div className={styles.accessNotice} role="alert">
      <span className={styles.accessIcon}><Icon iconName={error.kind === 'failed' ? 'Warning' : 'Lock'} /></span>
      <div className={styles.ccNoticeText}>
        <strong>{error.kind === 'failed' ? 'Lizenzen konnten nicht geladen werden' : 'Keine Berechtigung'}</strong>
        <span>{error.message}</span>
        {error.detail && error.detail !== error.message && <span className={styles.ccDetail}>Technisch: {error.detail}</span>}
      </div>
      <div className={styles.accessActions}>
        <button type="button" className={styles.textButton} onClick={() => start(true)}>Erneut versuchen</button>
      </div>
    </div>
  );

  if (!data) {
    return (
      <div className={styles.ov}>
        {hero}
        {notice}
        {loading && (
          <div className={styles.ovCard}>
            <div className={styles.ccLoading}>
              <span className={styles.skeleton} style={{ width: '40%', height: 14 }} />
              <span className={styles.muted}>{progress ? `${formatNumber(progress)} Benutzer geladen …` : 'Benutzer und Lizenzen werden gelesen …'}</span>
            </div>
          </div>
        )}
      </div>
    );
  }

  const licensed = users.filter(quickTest.licensed).length;
  const unlicensed = users.length - licensed;
  const disabledLicensed = users.filter(quickTest.disabledLicensed).length;
  const guestLicensed = users.filter(quickTest.guestLicensed).length;
  const quickTile = (k: Quick): { active: boolean; onClick: () => void } => ({ active: quick === k, onClick: () => setQuick(quick === k ? 'all' : k) });

  const skuCols: IColumn<ISku>[] = [
    {
      key: 'name', label: 'Lizenz', sort: s => s.name.toLowerCase(),
      render: s => (
        <span className={styles.ulSku}>
          <Icon iconName={selected.indexOf(s.id) >= 0 ? 'CheckboxComposite' : 'Checkbox'} />
          <span><b>{s.name}</b>{s.part && s.part !== s.name && <span className={styles.ccSub}> {s.part}</span>}</span>
        </span>
      )
    },
    { key: 'assigned', label: 'Zugewiesen', numeric: true, sort: s => s.assigned, render: s => formatNumber(s.assigned) },
    { key: 'enabled', label: 'Gekauft', numeric: true, sort: s => s.enabled === undefined ? -1 : s.enabled, render: s => (s.enabled === undefined ? '–' : formatNumber(s.enabled)) },
    {
      key: 'free', label: 'Frei', numeric: true, sort: s => (s.enabled === undefined || s.consumed === undefined ? -1 : s.enabled - s.consumed),
      render: s => (s.enabled === undefined || s.consumed === undefined ? '–' : formatNumber(Math.max(0, s.enabled - s.consumed)))
    }
  ];

  const userCols: IColumn<ILicUser>[] = [
    {
      key: 'name', label: 'Benutzer', sort: u => u.name.toLowerCase(),
      render: u => <span className={styles.ulUser}><b>{u.name}</b><span className={styles.ccSub}>{u.upn}</span></span>
    },
    { key: 'dept', label: 'Abteilung', sort: u => u.department.toLowerCase(), render: u => u.department || <span className={styles.muted}>–</span> },
    { key: 'status', label: 'Status', sort: u => (u.enabled ? 0 : 1), render: u => <Pill tone={u.enabled ? 'success' : 'danger'}>{u.enabled ? 'Aktiv' : 'Deaktiviert'}</Pill> },
    { key: 'type', label: 'Typ', sort: u => (u.guest ? 1 : 0), render: u => (u.guest ? <Pill tone="warning">Gast</Pill> : <span className={styles.muted}>Mitglied</span>) },
    {
      key: 'licenses', label: 'Lizenzen', sort: u => u.skus.map(skuName).sort().join(', '),
      render: u => (u.skus.length
        ? <span className={styles.ulChips}>{u.skus.map(skuName).sort().map(n => <span key={n} className={styles.ulChip}>{n}</span>)}</span>
        : <span className={styles.muted}>keine</span>)
    },
    { key: 'count', label: 'Anzahl', numeric: true, sort: u => u.skus.length, render: u => formatNumber(u.skus.length) }
  ];

  const options: IDropdownOption[] = [{ key: NONE, text: `Ohne Lizenz (${formatNumber(unlicensed)})` }]
    .concat(skus.map(s => ({ key: s.id, text: `${s.name} (${formatNumber(s.assigned)})` })));

  return (
    <div className={styles.ov}>
      {hero}
      {notice}
      <div className={styles.ccGrid}>
        <Tile label="Benutzer" hint={`davon ${formatNumber(users.filter(u => u.guest).length)} Gäste`} icon="People" tone={styles.ovToneInfo} value={formatNumber(users.length)} {...quickTile('all')} />
        <Tile label="Mit Lizenz" hint={`${formatNumber(skus.length)} verschiedene Lizenzen`} icon="ContactCard" tone={styles.ovToneInfo} value={formatNumber(licensed)} {...quickTile('licensed')} />
        <Tile label="Ohne Lizenz" hint="keine zugewiesene Lizenz" icon="Contact" tone={styles.ovToneInfo} value={formatNumber(unlicensed)} {...quickTile('unlicensed')} />
        <Tile label="Deaktiviert mit Lizenz" hint="Lizenz freigeben?" icon="Blocked2" tone={disabledLicensed ? styles.ovToneWarning : styles.ovToneInfo} value={formatNumber(disabledLicensed)} {...quickTile('disabledLicensed')} />
        <Tile label="Gäste mit Lizenz" hint="selten gewollt" icon="UserWarning" tone={guestLicensed ? styles.ovToneWarning : styles.ovToneInfo} value={formatNumber(guestLicensed)} {...quickTile('guestLicensed')} />
      </div>
      {(data.skuProblem || data.truncated) && (
        <div className={styles.ccWarnings}>
          {data.skuProblem && <div className={styles.muted}><Icon iconName="Info" /> {data.skuProblem}</div>}
          {data.truncated && <div className={styles.muted}><Icon iconName="Info" /> Sehr viele Benutzer – es wurden nur die ersten {formatNumber(users.length)} geladen.</div>}
        </div>
      )}

      <section className={styles.ovCard}>
        <div className={styles.ulHead}>
          <b>Lizenzübersicht</b>
          <span className={styles.muted}>Zeile anklicken = nach dieser Lizenz filtern</span>
        </div>
        <SortTable<ISku>
          columns={skuCols}
          rows={skus}
          rowKey={s => s.id}
          initial={{ key: 'assigned', desc: true }}
          empty="Keine Lizenzen zugewiesen."
          onRow={s => toggleSku(s.id)}
          rowClass={s => (selected.indexOf(s.id) >= 0 ? styles.ulRowSelected : undefined)}
        />
      </section>

      <section className={styles.ovCard}>
        <div className={styles.toolbar}>
          <SearchBox
            placeholder="Name, Benutzername oder Abteilung"
            value={search}
            onChange={(_e, v) => setSearch(v || '')}
            styles={{ root: { width: 260, borderRadius: 6 } }}
          />
          <Dropdown
            placeholder="Nach Lizenz filtern"
            multiSelect
            options={options}
            selectedKeys={selected}
            onChange={(_e, o) => o && toggleSku(String(o.key))}
            styles={{ root: { width: 300 }, dropdown: { borderRadius: 6 } }}
            ariaLabel="Nach Lizenz filtern"
          />
          <div className={styles.ccSegments} role="group" aria-label="Verknüpfung">
            {(['any', 'all'] as ('any' | 'all')[]).map(m => (
              <button
                key={m}
                type="button"
                className={`${styles.ccSegment} ${mode === m ? styles.ccSegmentActive : ''}`}
                aria-pressed={mode === m}
                disabled={selected.length < 2}
                title={m === 'any' ? 'Benutzer mit mindestens einer der gewählten Lizenzen' : 'Benutzer mit allen gewählten Lizenzen'}
                onClick={() => setMode(m)}
              >
                {m === 'any' ? 'mind. eine' : 'alle'}
              </button>
            ))}
          </div>
          <span className={`${styles.resultBadge} ${narrowed ? styles.resultBadgeActive : ''}`} role="status" aria-live="polite">
            {narrowed && <Icon iconName="Filter" />}
            {narrowed
              ? <><strong>{formatNumber(visible.length)}</strong> von {formatNumber(users.length)} Benutzern</>
              : <><strong>{formatNumber(users.length)}</strong> Benutzer</>}
          </span>
          {(selected.length > 0 || quick !== 'all' || !!q) && (
            <button type="button" className={styles.textButton} onClick={() => { setSelected([]); setQuick('all'); setSearch(''); }}>Filter zurücksetzen</button>
          )}
          <span className={styles.grow} />
          <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportCsv} disabled={!visible.length} title="Gefilterte Benutzer als CSV (Excel)" />
        </div>
        <SortTable<ILicUser>
          columns={userCols}
          rows={visible}
          rowKey={u => u.id}
          initial={{ key: 'name', desc: false }}
          empty={users.length ? 'Keine Benutzer für diesen Filter.' : 'Keine Benutzer gefunden.'}
          limit={limit}
        />
        {visible.length > limit && (
          <div className={styles.ccFoot}>
            <span>{formatNumber(limit)} von {formatNumber(visible.length)} angezeigt.</span>
            <button type="button" className={styles.textButton} onClick={() => setLimit(limit + PAGE)}>Weitere {formatNumber(Math.min(PAGE, visible.length - limit))} anzeigen</button>
          </div>
        )}
        <div className={styles.ccFoot}>
          Quelle: Microsoft Graph (zugewiesene Lizenzen je Benutzer{data.skuSource === 'subscribedSkus' ? ', gekaufte Mengen aus „subscribedSkus“' : ''}) – mit den Rechten der angemeldeten Person.
        </div>
      </section>
    </div>
  );
};
