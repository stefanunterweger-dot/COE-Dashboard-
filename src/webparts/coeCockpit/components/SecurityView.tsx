import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord, ITableConfig } from '../models/ITableConfig';
import { DataverseService, ICountResult } from '../services/DataverseService';
import { IColumnFilter } from '../utils/filters';
import { displayValue, formatNumber } from '../utils/format';
import { BarList, IBarItem } from './Charts';
import { Card, Failed, Skeleton } from './OverviewView';

/**
 * Ansicht „Sicherheit & Freigaben“ – Risiken aus den CoE-Tabellen auf einen Blick (nur lesend):
 * Oversharing, Gäste, unverschlüsselte Custom Connectors, Environments ohne Sicherheitsgruppe/DLP,
 * deaktivierte Konten mit Ressourcen, System-Administratoren je Environment, DLP-Verstöße.
 */
export const SECURITY_KEY = 'security';
const SYSADMIN_TEMPLATE = '627090ff-40a3-4053-8790-584edc5be201';

export interface ISecurityProps {
  service: DataverseService;
  tables: ITableConfig[];
  reloadKey: string;
  onNavigate: (tableKey: string, filters?: { [field: string]: IColumnFilter }) => void;
  onOpenRecord: (table: ITableConfig, record: IRecord, primaryId: string) => void;
}

interface ISection<T> { status: 'loading' | 'ready' | 'error'; data?: T; error?: string }

interface IRisk {
  key: string;
  label: string;
  hint: string;
  icon: string;
  tone: string;
  logical: string;
  odata: string;
  nav?: { [field: string]: IColumnFilter };
}

export const SecurityView: React.FC<ISecurityProps> = props => {
  const { service, tables, onNavigate, onOpenRecord } = props;
  const byLogical = (logical: string): ITableConfig | undefined => tables.filter(t => t.logicalName === logical)[0];

  const risks: IRisk[] = [
    { key: 'tenantApps', label: 'Apps für die gesamte Organisation freigegeben', hint: 'Jede Person im Tenant kann die App öffnen', icon: 'People', tone: styles.ovToneSerious, logical: 'admin_app', odata: 'admin_appsharedwithtenant eq true', nav: { admin_appsharedwithtenant: { values: ['Ja'] } } },
    { key: 'guestApps', label: 'Apps an Gäste freigegeben', hint: 'Externe Konten haben Zugriff', icon: 'Contact', tone: styles.ovToneWarning, logical: 'admin_app', odata: 'admin_appsharedguestuser eq true', nav: { admin_appsharedguestuser: { values: ['Ja'] } } },
    { key: 'guestFlows', label: 'Flows an Gäste freigegeben', hint: 'Externe Konten als Mitbesitzer/Benutzer', icon: 'Contact', tone: styles.ovToneWarning, logical: 'admin_flow', odata: 'admin_flowsharedguestuser eq true', nav: { admin_flowsharedguestuser: { values: ['Ja'] } } },
    { key: 'http', label: 'Custom Connectors über unverschlüsseltes HTTP', hint: 'Daten werden ohne TLS übertragen', icon: 'Unlock', tone: styles.ovToneCritical, logical: 'admin_connector', odata: 'admin_customconnectorhostishttp eq true', nav: { admin_customconnectorhostishttp: { values: ['Ja'] } } },
    { key: 'noSecGroup', label: 'Environments ohne Sicherheitsgruppe', hint: 'Zugriff nicht auf eine Entra-Gruppe beschränkt', icon: 'Globe', tone: styles.ovToneWarning, logical: 'admin_environment', odata: "(admin_securitygroupname eq 'Not assigned' or admin_securitygroupname eq null)", nav: { admin_securitygroupname: { values: ['Keine', ''] } } },
    { key: 'noDlp', label: 'Environments ohne DLP-Richtlinie', hint: 'Keine Datenverlust-Richtlinie wirksam', icon: 'Shield', tone: styles.ovToneSerious, logical: 'admin_environment', odata: '(admin_hasdlppolicy eq false or admin_hasdlppolicy eq null)', nav: { admin_hasdlppolicy: { values: ['Nein', ''] } } },
    { key: 'dlpApps', label: 'Apps mit DLP-Verstößen', hint: 'Laut DLP-Auswertung des CoE', icon: 'Blocked2', tone: styles.ovToneCritical, logical: 'admin_app', odata: 'admin_dlppoliciesviolated ne null' },
    { key: 'disabled', label: 'Deaktivierte Konten im Inventar', hint: 'Personen nicht mehr aktiv – ihre Apps/Flows übernehmen', icon: 'UserWarning', tone: styles.ovToneSerious, logical: 'admin_maker', odata: 'admin_accountenabled eq false', nav: { admin_accountenabled: { values: ['Nein'] } } },
    { key: 'guests', label: 'Gast-Maker', hint: 'Externe Konten, die Ressourcen erstellen', icon: 'Contact', tone: styles.ovToneInfo, logical: 'admin_maker', odata: 'admin_userisguest eq true', nav: { admin_userisguest: { values: ['Ja'] } } }
  ];

  const [counts, setCounts] = React.useState<{ [k: string]: ISection<ICountResult> }>({});
  const [admins, setAdmins] = React.useState<ISection<{ records: IRecord[]; counts: { [envId: string]: number }; primaryId: string }>>({ status: 'loading' });
  const [dlpApps, setDlpApps] = React.useState<ISection<{ records: IRecord[]; primaryId: string }>>({ status: 'loading' });
  const [policies, setPolicies] = React.useState<ISection<IRecord[]>>({ status: 'loading' });

  React.useEffect(() => {
    let cancelled = false;
    risks.forEach(r => {
      const t = byLogical(r.logical);
      if (!t) {
        setCounts(prev => ({ ...prev, [r.key]: { status: 'error' } }));
        return;
      }
      setCounts(prev => ({ ...prev, [r.key]: { status: 'loading' } }));
      service.getCountWithBase(t, r.odata).then(
        c => { if (!cancelled) { setCounts(prev => ({ ...prev, [r.key]: { status: 'ready', data: c } })); } },
        (e: Error) => { if (!cancelled) { setCounts(prev => ({ ...prev, [r.key]: { status: 'error', error: e.message } })); } }
      );
    });

    // Environments mit den meisten System-Administratoren
    const envT = byLogical('admin_environment');
    setAdmins({ status: 'loading' });
    // Ohne SYSTEM und App-Benutzer (wie im Reiter „Betroffen ab 15.12.“); ältere CoE-Versionen ohne Benutzertyp → einfacher Filter
    const sysAdmin = `(admin_securityroletemplateid eq '${SYSADMIN_TEMPLATE}' or admin_securityrolename eq 'System Administrator')`;
    const adminFilters = [
      `${sysAdmin} and admin_name ne 'SYSTEM' and (admin_usertype ne 597910002 or admin_usertype eq null)`,
      `${sysAdmin} and admin_name ne 'SYSTEM'`
    ];
    const groupedAdmins = (i: number): Promise<{ [value: string]: number }> =>
      service.getGroupedCounts('admin_environmentsecurityrolepermission', '_admin_environment_value', adminFilters[i])
        .catch(e => (i + 1 < adminFilters.length ? groupedAdmins(i + 1) : Promise.reject(e)));
    groupedAdmins(0)
      .then(async c => {
        const ids = Object.keys(c).filter(id => /^[0-9a-f-]{36}$/i.test(id)).sort((a, b) => c[b] - c[a]).slice(0, 8);
        if (!ids.length || !envT) {
          return { records: [], counts: c, primaryId: 'admin_environmentid' };
        }
        const res = await service.getRecords({ ...envT, key: '_secenv', strictFilter: true, filter: ids.map(id => `admin_environmentid eq ${id}`).join(' or ') }, 8);
        return { records: res.records, counts: c, primaryId: res.meta.primaryId };
      })
      .then(d => { if (!cancelled) { setAdmins({ status: 'ready', data: d }); } }, (e: Error) => { if (!cancelled) { setAdmins({ status: 'error', error: e.message }); } });

    // Apps mit DLP-Verstößen (Liste)
    const appsT = byLogical('admin_app');
    if (appsT) {
      service.getRecords({ ...appsT, key: '_dlpapps', filter: `(${appsT.filter || 'statecode eq 0'}) and (admin_dlppoliciesviolated ne null)`, strictFilter: true }, 8)
        .then(res => { if (!cancelled) { setDlpApps({ status: 'ready', data: { records: res.records, primaryId: res.meta.primaryId } }); } },
          (e: Error) => { if (!cancelled) { setDlpApps({ status: 'error', error: e.message }); } });
    } else {
      setDlpApps({ status: 'error' });
    }

    // DLP-Richtlinien
    const dlpT = byLogical('admin_dlppolicy');
    if (dlpT) {
      service.getRecords(dlpT, 500).then(res => { if (!cancelled) { setPolicies({ status: 'ready', data: res.records }); } },
        (e: Error) => { if (!cancelled) { setPolicies({ status: 'error', error: e.message }); } });
    } else {
      setPolicies({ status: 'error' });
    }
    return () => { cancelled = true; };
  }, [service, props.reloadKey]);

  const envT = byLogical('admin_environment');
  const appsT = byLogical('admin_app');
  const dlpT = byLogical('admin_dlppolicy');

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="Shield" /> Sicherheit &amp; Freigaben</div>
          <div className={styles.secHeroTitle}>Risiken im Tenant auf einen Blick</div>
          <div className={styles.ovHeroLabel}>
            Alle Werte stammen aus dem CoE-Inventar. Ein Klick führt in die gefilterte Liste – dort lassen sich die Einträge
            prüfen, exportieren oder im Admin Center öffnen.
          </div>
        </div>
      </section>

      <div className={styles.secGrid}>
        {risks.map(r => {
          const c = counts[r.key];
          const t = byLogical(r.logical);
          const clickable = !!t && !!r.nav;
          return (
            <button
              key={r.key}
              type="button"
              className={styles.secTile}
              disabled={!clickable}
              onClick={() => t && r.nav && onNavigate(t.key, r.nav)}
              title={clickable ? 'In der Liste anzeigen' : r.hint}
            >
              <span className={`${styles.govIcon} ${r.tone}`}><Icon iconName={r.icon} /></span>
              <span className={styles.govText}>
                <span className={styles.govLabel}>{r.label}</span>
                <span className={styles.govHint}>{r.hint}</span>
              </span>
              <span className={styles.govValue}>
                {!c || c.status === 'loading'
                  ? <span className={styles.skeleton} style={{ width: 36, height: 18 }} />
                  : c.status === 'error' || !c.data ? <span title={c.error || 'Nicht verfügbar'}>–</span> : `${formatNumber(c.data.value)}${c.data.atLeast ? '+' : ''}`}
              </span>
            </button>
          );
        })}
      </div>

      <div className={styles.ovRow3}>
        <Card title="Meiste System-Administratoren" subtitle="Environments nach Anzahl Personen mit dieser Rolle" icon="Shield">
          {admins.status === 'loading' && <Skeleton rows={6} />}
          {admins.status === 'error' && <Failed text="das CoE erfasst keine Sicherheitsrollen (Tabelle „Tenant Security Roles“ → „Track SR“)" />}
          {admins.data && admins.data.records.length === 0 && admins.status === 'ready' && <Failed text="keine Daten – Sicherheitsrollen-Sync im CoE einschalten" />}
          {admins.data && (
            <BarList unit="System-Administratoren" items={admins.data.records
              .map(r => ({ r, n: admins.data!.counts[String(r[admins.data!.primaryId]).toLowerCase()] || 0 }))
              .sort((a, b) => b.n - a.n)
              .map(({ r, n }, i): IBarItem => ({
                key: String(r[admins.data!.primaryId] || i),
                label: displayValue(r, 'admin_displayname'),
                sub: displayValue(r, 'admin_environmentsku'),
                value: n,
                onClick: envT ? () => onOpenRecord(envT, r, admins.data!.primaryId) : undefined
              }))} />
          )}
        </Card>

        <Card title="Apps mit DLP-Verstößen" subtitle="laut DLP-Auswertung im CoE" icon="Blocked2">
          {dlpApps.status === 'loading' && <Skeleton rows={6} />}
          {dlpApps.status === 'error' && <Failed text={dlpApps.error} />}
          {dlpApps.data && dlpApps.data.records.length === 0 && <div className={styles.ovFailed}><Icon iconName="CheckMark" /> Keine Verstöße gemeldet</div>}
          {dlpApps.data && dlpApps.data.records.length > 0 && (
            <ul className={styles.govList}>
              {dlpApps.data.records.map((r, i) => (
                <li key={String(r[dlpApps.data!.primaryId] || i)}>
                  <button type="button" className={styles.govItem} onClick={() => appsT && onOpenRecord(appsT, r, dlpApps.data!.primaryId)}>
                    <span className={`${styles.govIcon} ${styles.ovToneCritical}`}><Icon iconName="AppIconDefault" /></span>
                    <span className={styles.govText}>
                      <span className={styles.govLabel}>{displayValue(r, 'admin_displayname')}</span>
                      <span className={styles.govHint}>{displayValue(r, 'admin_dlppoliciesviolated')} · {displayValue(r, '_admin_appenvironment_value')}</span>
                    </span>
                    <Icon iconName="ChevronRight" className={styles.govChevron} />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </Card>

        <Card title="DLP-Richtlinien" subtitle="aus dem CoE-Inventar" icon="Shield" action={dlpT ? { label: 'Alle', onClick: () => onNavigate(dlpT.key) } : undefined}>
          {policies.status === 'loading' && <Skeleton rows={6} />}
          {policies.status === 'error' && <Failed text="DLP-Richtlinien werden vom CoE nicht erfasst" />}
          {policies.data && policies.data.length === 0 && <Failed text="keine DLP-Richtlinien im CoE-Inventar" />}
          {policies.data && policies.data.length > 0 && (
            <ul className={styles.govList}>
              {policies.data.slice(0, 8).map((r, i) => (
                <li key={String(r.admin_dlppolicyid || i)}>
                  <button type="button" className={styles.govItem} onClick={() => dlpT && onOpenRecord(dlpT, r, 'admin_dlppolicyid')}>
                    <span className={`${styles.govIcon} ${styles.ovToneInfo}`}><Icon iconName="Shield" /></span>
                    <span className={styles.govText}>
                      <span className={styles.govLabel}>{displayValue(r, 'admin_name')}</span>
                      <span className={styles.govHint}>{displayValue(r, 'admin_scope')}</span>
                    </span>
                    <Icon iconName="ChevronRight" className={styles.govChevron} />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </Card>
      </div>
    </div>
  );
};
