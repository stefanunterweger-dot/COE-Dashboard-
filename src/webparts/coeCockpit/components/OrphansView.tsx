import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { TextField } from '@fluentui/react/lib/TextField';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { IColumn, SortTable, Tile } from './CopilotCreditsView';
import { IPickedPerson, PersonPicker } from './PersonPicker';
import { IActionRequest, resourceOf } from '../models/actions';
import { IOrphan, IOrphanData, loadOrphans } from '../models/orphans';
import { ActionService } from '../services/ActionService';
import { DataverseService } from '../services/DataverseService';
import { DirectActionError, DirectActionService } from '../services/DirectActionService';
import { downloadText, formatDate, formatNumber } from '../utils/format';

/**
 * „Verwaiste Ressourcen“ (Auswertungen): Apps und Flows, deren Besitzer laut CoE nicht mehr im Unternehmen ist –
 * gruppiert nach ehemaligem Besitzer. Mitglieder der Admin-Gruppe mit aktiver Rolle weisen ausgewählte Ressourcen
 * gesammelt einer neuen Person zu (direkt über die Admin-Schnittstellen, jede Zuweisung im Aktionsprotokoll).
 */
export interface IOrphansViewProps {
  service: DataverseService;
  maxRecords: number;
  reloadKey: string;
  /** Darf neu zuweisen (Gruppe + aktive Rolle + Aktionsliste) */
  canAct: boolean;
  /** Warum nicht (Hinweis für alle anderen) */
  actHint: string;
  actions: ActionService | undefined;
  direct: DirectActionService;
  requestedBy: string;
  onOpen: (logicalName: string, id: string) => void;
  onActionsChanged: () => void;
}

const PAGE = 300;
const MAX_RUN = 100;
type TypeFilter = 'all' | 'app' | 'flow';
interface IRunResult { key: string; ok: boolean; text: string }

export const OrphansView: React.FC<IOrphansViewProps> = props => {
  const [data, setData] = React.useState<IOrphanData | undefined>(undefined);
  const [error, setError] = React.useState<string>('');
  const [loading, setLoading] = React.useState<boolean>(true);
  const [search, setSearch] = React.useState<string>('');
  const [type, setType] = React.useState<TypeFilter>('all');
  const [owner, setOwner] = React.useState<string>('');
  const [quick, setQuick] = React.useState<'' | 'solution' | 'active'>('');
  const [selected, setSelected] = React.useState<{ [key: string]: true }>({});
  const [done, setDone] = React.useState<{ [key: string]: IRunResult }>({});
  const [limit, setLimit] = React.useState<number>(PAGE);
  const [dialog, setDialog] = React.useState<boolean>(false);

  const load = (): void => {
    setLoading(true);
    setError('');
    loadOrphans(props.service, props.maxRecords).then(
      d => { setData(d); setLoading(false); },
      (e: unknown) => { setError(e instanceof Error ? e.message : String(e)); setLoading(false); }
    );
  };
  React.useEffect(() => { load(); }, [props.service, props.reloadKey]);
  React.useEffect(() => { setLimit(PAGE); }, [search, type, owner, quick]);

  const items = data ? data.items : [];
  const q = search.trim().toLowerCase();
  const visible = items.filter(i => (type === 'all' || i.type === type)
    && (!owner || i.ownerId === owner)
    && (quick !== 'solution' || i.inSolution === true)
    && (quick !== 'active' || (i.type === 'flow' && /aktiv/i.test(i.state)))
    && (!q || `${i.name} ${i.environment} ${i.ownerName}`.toLowerCase().indexOf(q) >= 0));
  const selectable = (i: IOrphan): boolean => props.canAct && !(done[i.key] && done[i.key].ok);
  const selectedItems = items.filter(i => selected[i.key] && selectable(i));
  const allVisibleSelected = visible.filter(selectable).length > 0 && visible.filter(selectable).every(i => selected[i.key]);

  // Ehemalige Besitzer
  const owners: { id: string; name: string; apps: number; flows: number }[] = [];
  const byOwner: { [id: string]: { id: string; name: string; apps: number; flows: number } } = {};
  items.forEach(i => {
    const id = i.ownerId || '—';
    if (!byOwner[id]) {
      byOwner[id] = { id, name: i.ownerName, apps: 0, flows: 0 };
      owners.push(byOwner[id]);
    }
    if (i.type === 'app') { byOwner[id].apps++; } else { byOwner[id].flows++; }
  });

  const toggleAll = (): void => {
    setSelected(prev => {
      const next = { ...prev };
      visible.filter(selectable).forEach(i => {
        if (allVisibleSelected) { delete next[i.key]; } else { next[i.key] = true; }
      });
      return next;
    });
  };

  const exportCsv = (): void => {
    const esc = (v: string): string => (/[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
    const lines = ['Typ;Name;Environment;Ehemaliger Besitzer;Zuletzt;In Solution;Status;Neu zugewiesen'].concat(visible.map(i => [
      i.type === 'app' ? 'App' : 'Flow', i.name, i.environment, i.ownerName, i.last ? formatDate(new Date(i.last)) : '',
      i.inSolution === undefined ? '' : i.inSolution ? 'Ja' : 'Nein', i.state, done[i.key] ? done[i.key].text : ''
    ].map(esc).join(';')));
    downloadText(`Verwaiste Ressourcen_${new Date().toISOString().slice(0, 10)}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  };

  const cols: IColumn<IOrphan>[] = ([] as IColumn<IOrphan>[]).concat(props.canAct ? [{
    key: 'sel', label: '', sort: (i: IOrphan) => (selected[i.key] ? 0 : 1),
    render: (i: IOrphan) => (
      <input type="checkbox" aria-label={`${i.name} auswählen`} disabled={!selectable(i)} checked={!!selected[i.key]}
        onClick={e => e.stopPropagation()}
        onChange={() => setSelected(prev => { const n = { ...prev }; if (n[i.key]) { delete n[i.key]; } else { n[i.key] = true; } return n; })} />
    )
  }] : []).concat([
    { key: 'type', label: 'Typ', sort: i => i.type, render: i => <span className={styles.ulSku}><Icon iconName={i.type === 'app' ? 'AppIconDefault' : 'Flow'} />{i.type === 'app' ? 'App' : 'Flow'}</span> },
    {
      key: 'name', label: 'Name', sort: i => i.name.toLowerCase(),
      render: i => <button type="button" className={styles.linkButton} onClick={e => { e.stopPropagation(); props.onOpen(i.logicalName, i.key.split(':')[1]); }}>{i.name || '(ohne Namen)'}</button>
    },
    { key: 'env', label: 'Environment', sort: i => i.environment.toLowerCase(), render: i => i.environment || <span className={styles.muted}>–</span> },
    { key: 'owner', label: 'Ehemaliger Besitzer', sort: i => i.ownerName.toLowerCase(), render: i => i.ownerName },
    { key: 'last', label: 'Zuletzt gestartet / geändert', sort: i => i.last, render: i => (i.last ? formatDate(new Date(i.last)) : <span className={styles.muted}>nie</span>) },
    {
      key: 'flags', label: 'Hinweis', sort: i => `${i.inSolution ? 1 : 0}${i.state}`,
      render: i => (
        <span className={styles.ulChips}>
          {i.inSolution && <Pill tone="warning" title="Teil einer Solution – Besitzerwechsel kann scheitern (dann im Environment über die Solution)">In Solution</Pill>}
          {i.type === 'flow' && i.state && <Pill tone={/aktiv/i.test(i.state) ? 'info' : /gesperrt/i.test(i.state) ? 'danger' : 'neutral'}>{i.state}</Pill>}
        </span>
      )
    },
    {
      key: 'result', label: 'Neu zugewiesen', sort: i => (done[i.key] ? (done[i.key].ok ? 0 : 1) : 2),
      render: i => (done[i.key] ? <Pill tone={done[i.key].ok ? 'success' : 'danger'} title={done[i.key].text}>{done[i.key].ok ? 'Erledigt' : 'Fehler'}</Pill> : <span className={styles.muted}>–</span>)
    }
  ]);

  const ownerCols: IColumn<{ id: string; name: string; apps: number; flows: number }>[] = [
    { key: 'name', label: 'Ehemaliger Besitzer', sort: o => o.name.toLowerCase(), render: o => <span className={styles.ulSku}><Icon iconName={owner === o.id ? 'CheckboxComposite' : 'Checkbox'} /><b>{o.name}</b></span> },
    { key: 'apps', label: 'Apps', numeric: true, sort: o => o.apps, render: o => formatNumber(o.apps) },
    { key: 'flows', label: 'Flows', numeric: true, sort: o => o.flows, render: o => formatNumber(o.flows) },
    { key: 'sum', label: 'Gesamt', numeric: true, sort: o => o.apps + o.flows, render: o => <b>{formatNumber(o.apps + o.flows)}</b> }
  ];

  const apps = items.filter(i => i.type === 'app').length;
  const flows = items.length - apps;
  const inSolution = items.filter(i => i.inSolution === true).length;
  const activeFlows = items.filter(i => i.type === 'flow' && /aktiv/i.test(i.state)).length;
  const narrowed = visible.length !== items.length;

  return (
    <div className={styles.ov}>
      <section className={`${styles.ovHero} ${styles.secHero}`}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}><Icon iconName="UserRemove" /> Aus dem CoE</div>
          <div className={styles.secHeroTitle}>Verwaiste Ressourcen</div>
          <div className={styles.ovHeroLabel}>
            Apps und Flows, deren Besitzer nicht mehr im Unternehmen ist. Wählen, einer neuen Person zuweisen – das Cockpit führt das direkt aus
            und vermerkt jede Zuweisung im Aktionsprotokoll.
          </div>
          <div className={styles.ovHeroChips}>
            <button type="button" className={styles.ovHeroChip} onClick={load} disabled={loading}><Icon iconName="Refresh" /> {loading ? 'Lädt …' : 'Aktualisieren'}</button>
            {data && <span className={styles.ccStamp}>Stand: {formatDate(data.loadedAt, true)}</span>}
          </div>
        </div>
      </section>
      {error && <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>Verwaiste Ressourcen konnten nicht geladen werden: {error}</MessageBar>}
      {loading && !data && <div className={styles.ovCard}><div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '40%', height: 14 }} /><span className={styles.muted}>Wird aus dem CoE geladen …</span></div></div>}
      {data && (
        <>
          <div className={styles.ccGrid}>
            <Tile label="Verwaiste Apps" hint="Besitzer nicht mehr im Unternehmen" icon="AppIconDefault" tone={apps ? styles.ovToneWarning : styles.ovToneInfo} value={formatNumber(apps)}
              active={type === 'app'} onClick={() => setType(type === 'app' ? 'all' : 'app')} />
            <Tile label="Verwaiste Flows" hint="laufen mit den Verbindungen des alten Besitzers" icon="Flow" tone={flows ? styles.ovToneWarning : styles.ovToneInfo} value={formatNumber(flows)}
              active={type === 'flow'} onClick={() => setType(type === 'flow' ? 'all' : 'flow')} />
            <Tile label="Davon Flows aktiv" hint="brechen ab, sobald Verbindungen ablaufen" icon="Play" tone={activeFlows ? styles.ovToneCritical : styles.ovToneInfo} value={formatNumber(activeFlows)}
              active={quick === 'active'} onClick={() => setQuick(quick === 'active' ? '' : 'active')} />
            <Tile label="In Solutions" hint="Besitzerwechsel über die Solution" icon="Package" tone={styles.ovToneInfo} value={formatNumber(inSolution)}
              active={quick === 'solution'} onClick={() => setQuick(quick === 'solution' ? '' : 'solution')} />
            <Tile label="Ehemalige Besitzer" hint="Zeile unten anklicken = filtern" icon="People" tone={styles.ovToneInfo} value={formatNumber(owners.length)} />
          </div>
          {(data.warnings.length > 0 || data.truncated) && (
            <div className={styles.ccWarnings}>
              {data.truncated && <div className={styles.muted}><Icon iconName="Info" /> Sehr viele Einträge – es wurden nur die ersten {formatNumber(props.maxRecords)} je Typ geladen.</div>}
              {data.warnings.map((w, i) => <div key={i} className={styles.muted}><Icon iconName="Info" /> {w}</div>)}
            </div>
          )}

          <section className={styles.ovCard}>
            <div className={styles.ulHead}><b>Nach ehemaligem Besitzer</b><span className={styles.muted}>Zeile anklicken = nur diese Person</span></div>
            <SortTable
              columns={ownerCols}
              rows={owners}
              rowKey={o => o.id}
              initial={{ key: 'sum', desc: true }}
              empty="Keine verwaisten Ressourcen – sehr gut."
              onRow={o => setOwner(owner === o.id ? '' : o.id)}
              rowClass={o => (owner === o.id ? styles.ulRowSelected : undefined)}
              limit={12}
            />
            {owners.length > 12 && <div className={styles.ccFoot}><span>Die 12 Personen mit den meisten Ressourcen – alle findest du über die Suche unten.</span></div>}
          </section>

          <section className={styles.ovCard}>
            <div className={styles.toolbar}>
              <SearchBox placeholder="App, Flow, Environment oder Person" value={search} onChange={(_e, v) => setSearch(v || '')} styles={{ root: { width: 260, borderRadius: 6 } }} />
              <span className={`${styles.resultBadge} ${narrowed ? styles.resultBadgeActive : ''}`} role="status" aria-live="polite">
                {narrowed && <Icon iconName="Filter" />}
                {narrowed ? <><strong>{formatNumber(visible.length)}</strong> von {formatNumber(items.length)} Einträgen</> : <><strong>{formatNumber(items.length)}</strong> Einträge</>}
              </span>
              {(narrowed || !!owner) && (
                <button type="button" className={styles.textButton} onClick={() => { setSearch(''); setType('all'); setOwner(''); setQuick(''); }}>Filter zurücksetzen</button>
              )}
              <span className={styles.grow} />
              {props.canAct && (
                <>
                  <button type="button" className={styles.textButton} onClick={toggleAll} disabled={!visible.some(selectable)}>
                    {allVisibleSelected ? 'Auswahl aufheben' : `Alle gefilterten auswählen (${formatNumber(Math.min(visible.filter(selectable).length, MAX_RUN))})`}
                  </button>
                  <PrimaryButton
                    text={`Neuen Besitzer zuweisen (${formatNumber(selectedItems.length)})`}
                    iconProps={{ iconName: 'Contact' }}
                    disabled={!selectedItems.length || !props.actions}
                    onClick={() => setDialog(true)}
                  />
                </>
              )}
              <DefaultButton text="CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportCsv} disabled={!visible.length} />
            </div>
            {!props.canAct && props.actHint && <div className={styles.ccWarnings}><div className={styles.muted}><Icon iconName="Info" /> {props.actHint}</div></div>}
            <SortTable<IOrphan>
              columns={cols}
              rows={visible}
              rowKey={i => i.key}
              initial={{ key: 'last', desc: false }}
              empty={items.length ? 'Keine Einträge für diesen Filter.' : 'Keine verwaisten Apps oder Flows – sehr gut.'}
              limit={limit}
            />
            {visible.length > limit && (
              <div className={styles.ccFoot}>
                <span>{formatNumber(limit)} von {formatNumber(visible.length)} angezeigt.</span>
                <button type="button" className={styles.textButton} onClick={() => setLimit(limit + PAGE)}>Weitere anzeigen</button>
              </div>
            )}
            <div className={styles.ccFoot}>
              Grundlage: CoE-Felder „App is Orphaned“ / „Flow is Orphaned“ (letzter CoE-Sync). Bei Flows wird die neue Person Mitbesitzer –
              den Ersteller kann Power Automate nicht ändern; danach am besten die Verbindungen des Flows auf die neue Person umstellen.
            </div>
          </section>
        </>
      )}
      {dialog && props.actions && (
        <ReassignDialog
          items={selectedItems.slice(0, MAX_RUN)}
          total={selectedItems.length}
          dataverse={props.service}
          actions={props.actions}
          direct={props.direct}
          requestedBy={props.requestedBy}
          onResult={r => setDone(prev => ({ ...prev, [r.key]: r }))}
          onClose={changed => {
            setDialog(false);
            if (changed) {
              setSelected({});
              props.onActionsChanged();
            }
          }}
        />
      )}
    </div>
  );
};

/** Sammel-Zuweisung mit Begründung, Fortschritt und Ergebnisliste. */
const ReassignDialog: React.FC<{
  items: IOrphan[]; total: number; dataverse: DataverseService; actions: ActionService; direct: DirectActionService; requestedBy: string;
  onResult: (r: IRunResult) => void; onClose: (changed: boolean) => void;
}> = props => {
  const [person, setPerson] = React.useState<IPickedPerson | undefined>(undefined);
  const [reason, setReason] = React.useState<string>('Besitzer hat das Unternehmen verlassen – Ressource wird übernommen.');
  const [phase, setPhase] = React.useState<'edit' | 'running' | 'done'>('edit');
  const [progress, setProgress] = React.useState<number>(0);
  const [results, setResults] = React.useState<{ item: IOrphan; ok: boolean; text: string }[]>([]);
  const stopRef = React.useRef<boolean>(false);

  const apps = props.items.filter(i => i.type === 'app').length;
  const flows = props.items.length - apps;
  const solutions = props.items.filter(i => i.inSolution === true).length;

  const run = async (): Promise<void> => {
    setPhase('running');
    stopRef.current = false;
    const out: { item: IOrphan; ok: boolean; text: string }[] = [];
    for (let n = 0; n < props.items.length && !stopRef.current; n++) {
      const item = props.items[n];
      const res = resourceOf(item.logicalName, item.record);
      if (!res) {
        continue;
      }
      const req: IActionRequest = {
        action: 'setowner', resourceType: res.type, resourceId: res.id, coeRecordId: res.coeId, resourceName: res.name, environmentId: res.env,
        targetUser: person!.id, targetName: person!.name, targetIsGroup: false, reason: reason.trim()
      };
      let ok = false;
      let text = '';
      try {
        const id = await props.actions.add(req, props.requestedBy, 'In Arbeit', 'Sammel-Zuweisung „Verwaiste Ressourcen“ – wird direkt ausgeführt …');
        try {
          const r = await props.direct.run(req);
          ok = true;
          text = `${r.text}${r.note}`;
        } catch (e) {
          text = e instanceof Error ? e.message : String(e);
          await props.actions.complete(id, e instanceof DirectActionError && e.kind === 'rejected' ? 'Abgelehnt' : 'Fehler', text.slice(0, 4000)).catch(() => undefined);
        }
        if (ok) {
          await props.actions.complete(id, 'Erledigt', text.slice(0, 4000)).catch(() => undefined);
        }
      } catch (e) {
        text = `Protokolleintrag nicht möglich – nicht ausgeführt: ${e instanceof Error ? e.message : String(e)}`;
      }
      out.push({ item, ok, text });
      props.onResult({ key: item.key, ok, text });
      setResults(out.slice());
      setProgress(n + 1);
    }
    setPhase('done');
  };

  const okCount = results.filter(r => r.ok).length;
  const failed = results.filter(r => !r.ok);

  return (
    <Dialog
      hidden={false}
      onDismiss={() => phase !== 'running' && props.onClose(results.length > 0)}
      minWidth={560}
      maxWidth={680}
      dialogContentProps={{ type: DialogType.normal, title: 'Neuen Besitzer zuweisen', subText: `${formatNumber(props.items.length)} ausgewählt: ${formatNumber(apps)} Apps, ${formatNumber(flows)} Flows` }}
      modalProps={{ isBlocking: phase === 'running' }}
    >
      <div className={styles.notify}>
        {props.total > props.items.length && (
          <MessageBar messageBarType={MessageBarType.info}>Pro Durchlauf höchstens {MAX_RUN} – die übrigen {formatNumber(props.total - props.items.length)} bleiben ausgewählt.</MessageBar>
        )}
        {phase === 'edit' && (
          <>
            <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
              Apps: Die neue Person wird Besitzer. Flows: Die neue Person wird Mitbesitzer (den Ersteller kann Power Automate nicht ändern).
              {solutions > 0 && ` ${formatNumber(solutions)} Einträge sind Teil einer Solution – dort kann der Wechsel scheitern.`}
              {' '}Ausgeführt wird direkt mit deinen Rechten; jede Zuweisung steht im Aktionsprotokoll.
            </MessageBar>
            <PersonPicker label="Neuer Besitzer" dataverse={props.dataverse} value={person} onChange={setPerson} />
            <TextField label="Begründung" required={true} multiline={true} rows={2} value={reason} onChange={(_e, v) => setReason(v || '')} />
          </>
        )}
        {phase !== 'edit' && (
          <ProgressIndicator
            label={phase === 'running' ? `Weise zu … ${formatNumber(progress)} von ${formatNumber(props.items.length)}` : `Fertig: ${formatNumber(okCount)} zugewiesen, ${formatNumber(failed.length)} Fehler`}
            percentComplete={props.items.length ? progress / props.items.length : 1}
          />
        )}
        {phase === 'done' && failed.length > 0 && (
          <div className={styles.bulkList}>
            {failed.slice(0, 20).map(f => <div key={f.item.key} className={styles.muted}><b>{f.item.name}:</b> {f.text}</div>)}
          </div>
        )}
      </div>
      <DialogFooter>
        {phase === 'edit' && (
          <PrimaryButton text="Zuweisen" iconProps={{ iconName: 'LightningBolt' }} disabled={!person || reason.trim().length < 5} onClick={() => { run().catch(() => setPhase('done')); }} />
        )}
        {phase === 'running' && <DefaultButton text="Anhalten" iconProps={{ iconName: 'Stop' }} onClick={() => { stopRef.current = true; }} />}
        {phase !== 'running' && <DefaultButton text={phase === 'edit' ? 'Abbrechen' : 'Schließen'} onClick={() => props.onClose(results.length > 0)} />}
      </DialogFooter>
    </Dialog>
  );
};
