import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { Tile } from './CopilotCreditsView';
import { CATEGORY_LABEL, IMcData, IMcMessage, loadMessages, MessageLoadError } from '../models/messageCenter';
import { formatDate, formatNumber } from '../utils/format';

/**
 * „Was ändert Microsoft?“ (Start): Meldungen des Message Centers zur Power Platform – live aus Microsoft Graph.
 * Handlung nötig (mit Frist), wichtige Änderungen, neue Meldungen; Text sicher in einem abgeschotteten Rahmen.
 */
export interface IMessageCenterViewProps {
  graphFactory: MSGraphClientFactory;
  reloadKey: string;
}

type Filter = 'all' | 'action' | 'major' | 'new';
const DAY = 86400000;

/** Skripte, eingebettete Rahmen und Ereignis-Attribute entfernen (zusätzlich zur Abschottung des Rahmens). */
const clean = (html: string): string => html
  .replace(/<(script|iframe|object|embed)[\s\S]*?<\/\1>/gi, '')
  .replace(/<(script|iframe|object|embed)[^>]*>/gi, '')
  .replace(/\son[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
  .replace(/javascript:/gi, '');

/** Meldungstext in einem Rahmen ohne Skripte (sandbox) – Links öffnen in einem neuen Fenster. */
const frameDoc = (html: string): string => `<!doctype html><html><head><meta charset="utf-8"><base target="_blank">
<style>body{font:14px/1.5 "Segoe UI",system-ui,sans-serif;color:#242424;margin:12px;}a{color:#0f6cbd}img{max-width:100%}table{border-collapse:collapse}td,th{border:1px solid #ddd;padding:4px 8px}</style>
</head><body>${clean(html)}</body></html>`;

export const MessageCenterView: React.FC<IMessageCenterViewProps> = props => {
  const [data, setData] = React.useState<IMcData | undefined>(undefined);
  const [error, setError] = React.useState<MessageLoadError | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [filter, setFilter] = React.useState<Filter>('all');
  const [service, setService] = React.useState<string>('');
  const [search, setSearch] = React.useState<string>('');
  const [sort, setSort] = React.useState<'modified' | 'due'>('modified');
  const [open, setOpen] = React.useState<string>('');

  const load = (fresh: boolean): void => {
    setLoading(true);
    setError(undefined);
    loadMessages(props.graphFactory, fresh).then(
      d => { setData(d); setLoading(false); },
      (e: unknown) => { setError(e instanceof MessageLoadError ? e : new MessageLoadError(e instanceof Error ? e.message : String(e), 'failed')); setLoading(false); }
    );
  };
  React.useEffect(() => { load(false); }, [props.graphFactory, props.reloadKey]);

  const now = Date.now();
  const due = (m: IMcMessage): number | undefined => {
    const t = Date.parse(m.actionBy);
    return isNaN(t) ? undefined : t;
  };
  const needsAction = (m: IMcMessage): boolean => { const t = due(m); return t !== undefined && t >= now - DAY; };
  const isNew = (m: IMcMessage): boolean => { const t = Date.parse(m.start || m.modified); return !isNaN(t) && now - t <= 7 * DAY; };
  const msgs = data ? data.messages : [];
  const services = msgs.reduce((acc, m) => { m.services.forEach(s => { acc[s] = (acc[s] || 0) + 1; }); return acc; }, {} as { [s: string]: number });
  const q = search.trim().toLowerCase();
  const visible = msgs.filter(m => (filter === 'all' || (filter === 'action' && needsAction(m)) || (filter === 'major' && m.major) || (filter === 'new' && isNew(m)))
    && (!service || m.services.indexOf(service) >= 0)
    && (!q || `${m.id} ${m.title} ${m.services.join(' ')} ${m.tags.join(' ')}`.toLowerCase().indexOf(q) >= 0))
    .sort((a, b) => (sort === 'due'
      ? (due(a) === undefined ? 1 : 0) - (due(b) === undefined ? 1 : 0) || (due(a) || 0) - (due(b) || 0)
      : (Date.parse(b.modified) || 0) - (Date.parse(a.modified) || 0)));
  const actionList = msgs.filter(needsAction).sort((a, b) => (due(a) || 0) - (due(b) || 0));
  const next = actionList[0];

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="Megaphone" /> Live aus dem Microsoft 365 Message Center</div>
          <div className={styles.secHeroTitle}>Was ändert Microsoft?</div>
          <div className={styles.ovHeroLabel}>
            Ankündigungen von Microsoft zur Power Platform (Power Apps, Power Automate, Copilot Studio, Dataverse, Power Pages, Power BI) –
            mit Frist, wenn Handlung nötig ist. So fallen Umstellungen wie „Februar 2027“ früh auf.
          </div>
          <div className={styles.ovHeroChips}>
            <button type="button" className={styles.ovHeroChip} onClick={() => load(true)} disabled={loading}><Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}</button>
            <a className={styles.ovHeroChip} href="https://admin.microsoft.com/Adminportal/Home#/MessageCenter" target="_blank" rel="noopener noreferrer"><Icon iconName="OpenInNewWindow" /> Message Center</a>
            {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
          </div>
        </div>
      </section>
      {error && (
        <div className={styles.accessNotice} role="alert">
          <span className={styles.accessIcon}><Icon iconName={error.kind === 'failed' ? 'Warning' : 'Lock'} /></span>
          <div className={styles.ccNoticeText}>
            <strong>{error.kind === 'failed' ? 'Meldungen konnten nicht geladen werden' : 'Keine Berechtigung'}</strong>
            <span>{error.message}</span>
            {error.detail && error.detail !== error.message && <span className={styles.ccDetail}>Technisch: {error.detail}</span>}
          </div>
          <div className={styles.accessActions}><button type="button" className={styles.textButton} onClick={() => load(true)}>Erneut versuchen</button></div>
        </div>
      )}
      {loading && !data && <div className={styles.ovCard}><div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>Meldungen werden gelesen …</span></div></div>}
      {data && (
        <>
          <div className={styles.ccGrid}>
            <Tile label="Power-Platform-Meldungen" hint={`von ${formatNumber(data.totalAll)} im Message Center`} icon="Megaphone" tone={styles.ovToneInfo} value={formatNumber(msgs.length)}
              active={filter === 'all'} onClick={() => setFilter('all')} />
            <Tile label="Handlung nötig" hint={next ? `nächste Frist: ${formatDate(new Date(next.actionBy))}` : 'keine offene Frist'} icon="Ringer"
              tone={actionList.length ? styles.ovToneCritical : styles.ovToneInfo} value={formatNumber(actionList.length)} active={filter === 'action'} onClick={() => setFilter(filter === 'action' ? 'all' : 'action')} />
            <Tile label="Wichtige Änderungen" hint="von Microsoft als „major“ markiert" icon="AlertSolid" tone={styles.ovToneWarning} value={formatNumber(msgs.filter(m => m.major).length)}
              active={filter === 'major'} onClick={() => setFilter(filter === 'major' ? 'all' : 'major')} />
            <Tile label="Neu" hint="in den letzten 7 Tagen" icon="News" tone={styles.ovToneInfo} value={formatNumber(msgs.filter(isNew).length)}
              active={filter === 'new'} onClick={() => setFilter(filter === 'new' ? 'all' : 'new')} />
          </div>
          <section className={styles.ovCard}>
            <div className={styles.toolbar}>
              <SearchBox placeholder="Titel, MC-Nummer oder Stichwort" value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 260, borderRadius: 6 } }} />
              <div className={styles.ccSegments} role="group" aria-label="Dienst">
                <button type="button" className={`${styles.ccSegment} ${!service ? styles.ccSegmentActive : ''}`} aria-pressed={!service} onClick={() => setService('')}>Alle Dienste</button>
                {Object.keys(services).sort((a, b) => services[b] - services[a]).slice(0, 7).map(s => (
                  <button key={s} type="button" className={`${styles.ccSegment} ${service === s ? styles.ccSegmentActive : ''}`} aria-pressed={service === s} onClick={() => setService(service === s ? '' : s)}>
                    {s} <span className={styles.ccSegmentCount}>{formatNumber(services[s])}</span>
                  </button>
                ))}
              </div>
              <div className={styles.ccSegments} role="group" aria-label="Sortierung">
                <button type="button" className={`${styles.ccSegment} ${sort === 'modified' ? styles.ccSegmentActive : ''}`} aria-pressed={sort === 'modified'} onClick={() => setSort('modified')}>Neueste zuerst</button>
                <button type="button" className={`${styles.ccSegment} ${sort === 'due' ? styles.ccSegmentActive : ''}`} aria-pressed={sort === 'due'} onClick={() => setSort('due')}>Nach Frist</button>
              </div>
              <span className={`${styles.resultBadge} ${visible.length !== msgs.length ? styles.resultBadgeActive : ''}`} role="status"><strong>{formatNumber(visible.length)}</strong>{visible.length !== msgs.length ? ` von ${formatNumber(msgs.length)}` : ''} Meldungen</span>
            </div>
            <div className={styles.mcList}>
              {visible.length === 0 && <div className={styles.solEmpty}><Icon iconName="Info" /><span>Keine Meldungen für diesen Filter.</span></div>}
              {visible.map(m => {
                const d = due(m);
                const soon = d !== undefined && d - now < 30 * DAY;
                return (
                  <div key={m.id} className={styles.mcItem} data-mc={m.id}>
                    <button type="button" className={styles.mcHead} aria-expanded={open === m.id} onClick={() => setOpen(open === m.id ? '' : m.id)}>
                      <Icon iconName={open === m.id ? 'ChevronDown' : 'ChevronRight'} />
                      <span className={styles.mcMain}>
                        <span className={styles.mcTitle}>{m.title}</span>
                        <span className={styles.mcMeta}>
                          <b>{m.id}</b>
                          <span>{m.services.join(', ')}</span>
                          {m.category && <span>· {CATEGORY_LABEL[m.category] || m.category}</span>}
                          {m.start && <span>· veröffentlicht {formatDate(new Date(m.start))}</span>}
                          {m.modified && m.modified !== m.start && <span>· aktualisiert {formatDate(new Date(m.modified))}</span>}
                          {m.major && <Pill tone="warning">Wichtige Änderung</Pill>}
                          {isNew(m) && <Pill tone="info">Neu</Pill>}
                        </span>
                      </span>
                      {needsAction(m) && <span className={`${styles.mcDue} ${soon ? styles.mcDueSoon : ''}`}>Handlung bis {formatDate(new Date(m.actionBy))}</span>}
                    </button>
                    {open === m.id && (
                      <div className={styles.mcBody}>
                        <iframe title={`Meldung ${m.id}`} sandbox="allow-popups allow-popups-to-escape-sandbox" srcDoc={frameDoc(m.html || '<p>(kein Text)</p>')} />
                        <div className={styles.mcLinks}>
                          <a className={styles.link} href={`https://admin.microsoft.com/Adminportal/Home#/MessageCenter/:/messages/${encodeURIComponent(m.id)}`} target="_blank" rel="noopener noreferrer">
                            <Icon iconName="OpenInNewWindow" /> Im Message Center öffnen
                          </a>
                          {m.links.map(l => <a key={l.url} className={styles.link} href={l.url} target="_blank" rel="noopener noreferrer"><Icon iconName="Link" /> {l.label}</a>)}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
            <div className={styles.ccFoot}>Quelle: Microsoft Graph (Message Center) – gefiltert auf Dienste der Power Platform. Gelesen/archiviert setzt du weiterhin im Message Center.</div>
          </section>
        </>
      )}
    </div>
  );
};
