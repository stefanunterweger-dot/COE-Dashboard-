import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { ISettingsValues, IWebPartSettings } from '../models/settings';
import { ACTION_LIST_TITLE } from '../models/actions';
import { parseAddresses } from '../models/notification';
import { isValidEnvironmentUrl } from '../utils/format';

/**
 * „Webpart-Einstellungen“: übersichtlicher Ersatz für den langen SharePoint-Eigenschaftenbereich.
 * Links Bereiche mit Status, rechts die Einstellungen des gewählten Bereichs. Änderungen gelten sofort
 * und werden mit der Seite gespeichert (im Bearbeitungsmodus „Erneut veröffentlichen“).
 */
type SectionKey = 'connection' | 'views' | 'default' | 'dev' | 'mail' | 'actions' | 'alm' | 'display' | 'tables';
type Tone = 'ok' | 'warn' | 'error' | 'off';

export interface ISettingsPanelProps {
  isOpen: boolean;
  settings: IWebPartSettings;
  themeVars: React.CSSProperties;
  onDismiss: () => void;
}

const https = (v: string): string => (!v || /^https:\/\/\S+$/i.test(v.trim()) ? '' : 'Bitte einen https-Link eintragen.');

/** Textfeld, das erst beim Verlassen (oder Enter) übernommen wird. */
const LazyText: React.FC<{
  label: string; value: string; placeholder?: string; description?: string; multiline?: boolean;
  validate?: (v: string) => string; onCommit: (v: string) => void;
}> = p => {
  const [draft, setDraft] = React.useState(p.value);
  React.useEffect(() => setDraft(p.value), [p.value]);
  const error = p.validate ? p.validate(draft) : '';
  const commit = (): void => {
    if (draft !== p.value && !error) {
      p.onCommit(draft.trim());
    }
  };
  return (
    <TextField
      label={p.label}
      value={draft}
      placeholder={p.placeholder}
      description={p.description}
      multiline={p.multiline}
      autoAdjustHeight={p.multiline}
      errorMessage={error || undefined}
      onChange={(_e, v) => setDraft(v || '')}
      onBlur={commit}
      onKeyDown={e => { if (e.key === 'Enter' && !p.multiline) { commit(); } }}
    />
  );
};

const Row: React.FC<{ label: string; description?: React.ReactNode }> = p => (
  <div className={styles.setRow}>
    <div className={styles.setRowText}>
      <div className={styles.setRowLabel}>{p.label}</div>
      {p.description && <div className={styles.setRowDesc}>{p.description}</div>}
    </div>
    <div className={styles.setRowControl}>{p.children}</div>
  </div>
);

const Check: React.FC<{ ok: boolean | undefined; warn?: boolean }> = p => (
  <span className={`${styles.setCheck} ${p.ok ? styles.setCheckOk : p.warn ? styles.setCheckWarn : styles.setCheckErr}`}>
    <Icon iconName={p.ok ? 'CompletedSolid' : p.warn ? 'WarningSolid' : 'StatusErrorFull'} />
  </span>
);

export const SettingsPanel: React.FC<ISettingsPanelProps> = props => {
  const { settings: s } = props;
  const v = s.values;
  const [section, setSection] = React.useState<SectionKey>('connection');
  const set = (patch: Partial<ISettingsValues>): void => s.onChange(patch);

  React.useEffect(() => {
    if (props.isOpen) {
      s.onLoad();
    }
  }, [props.isOpen]);

  const urlOk = isValidEnvironmentUrl(v.environmentUrl);
  const st = s.actionSetup;
  const actionsReady = !!st && !!st.group && st.listExists && !!st.listUnique;

  const sections: { key: SectionKey; label: string; icon: string; tone: Tone; hint: string }[] = [
    { key: 'connection', label: 'Verbindung', icon: 'PlugConnected', tone: urlOk ? 'ok' : 'error', hint: urlOk ? 'CoE-Umgebung eingetragen' : 'URL fehlt' },
    { key: 'views', label: 'Ansichten', icon: 'ViewDashboard', tone: 'ok', hint: 'Übersicht, Sicherheit, Copilot-Credits' },
    { key: 'default', label: 'Default Environment', icon: 'Home', tone: v.showDefaultEnvironment ? 'ok' : 'off', hint: v.showDefaultEnvironment ? 'eingeblendet' : 'ausgeblendet' },
    { key: 'dev', label: 'Betroffen ab 15.12.', icon: 'Warning', tone: v.showDevEnvironments ? 'ok' : 'off', hint: v.showDevEnvironments ? 'eingeblendet' : 'ausgeblendet' },
    { key: 'mail', label: 'E-Mail an Maker', icon: 'Mail', tone: v.showDevEnvironments ? 'ok' : 'off', hint: v.notifyAutoSend ? 'Outlook + automatisch' : 'nur Outlook' },
    {
      key: 'actions', label: 'Aktionen', icon: 'LightningBolt',
      tone: !v.showActions ? 'off' : actionsReady ? 'ok' : st ? 'warn' : 'off',
      hint: !v.showActions ? 'ausgeschaltet' : actionsReady ? 'eingerichtet' : st ? 'Einrichtung unvollständig' : 'wird geprüft …'
    },
    { key: 'alm', label: 'ALM-Pipelines', icon: 'Rocket', tone: v.pipelinesHostUrl ? 'ok' : 'off', hint: v.pipelinesHostUrl ? 'Host gewählt' : 'ausgeblendet' },
    { key: 'display', label: 'Darstellung', icon: 'Color', tone: 'ok', hint: 'Titel, Kennzahlen, Seitengröße' },
    { key: 'tables', label: 'Tabellen', icon: 'Table', tone: 'ok', hint: `${s.tableCount} Tabellen` }
  ];
  const current = sections.filter(x => x.key === section)[0];

  const intro: { [k in SectionKey]: string } = {
    connection: 'Aus dieser Dataverse-Umgebung mit dem CoE Starter Kit liest das Cockpit alle Daten – mit den Rechten der angemeldeten Person.',
    views: 'Welche Ansichten in der Navigation erscheinen.',
    default: 'Eigener Bereich für das Default Environment: Apps und Flows aus dem CoE, Solutions live aus dem Environment.',
    dev: 'Nicht verwaltete Developer-Environments, die Microsoft ab 15.12.2026 sperrt und ab 15.01.2027 löscht.',
    mail: 'Vorlage für die E-Mail an die Maker betroffener Developer-Environments.',
    actions: 'Aufträge wie Quarantäne, Löschen, Besitzer ändern oder Environment deaktivieren. Das Cockpit legt sie in einer SharePoint-Liste an, '
      + 'der Flow „CoE Cockpit | Aktion ausführen“ im CoE-Environment führt sie mit einem Admin-Dienstkonto aus.',
    alm: 'Bereitstellungen und Pipelines aus dem Dataverse des Pipelines-Host-Environments.',
    display: 'Überschrift, Kennzahlen und Tabellenansicht.',
    tables: 'Welche CoE-Tabellen mit welchen Spalten angezeigt werden.'
  };

  const renderSection = (): JSX.Element => {
    switch (section) {
      case 'connection':
        return (
          <div className={styles.setCard}>
            <LazyText
              label="URL der CoE-Umgebung"
              value={v.environmentUrl}
              placeholder="https://contoso-coe.crm4.dynamics.com"
              description="Dataverse-URL der Umgebung mit dem CoE Starter Kit"
              validate={x => (!x.trim() || isValidEnvironmentUrl(x) ? '' : 'Bitte eine Dataverse-URL eintragen, z. B. https://contoso-coe.crm4.dynamics.com')}
              onCommit={x => set({ environmentUrl: x })}
            />
            <div className={styles.setStatus}><Check ok={urlOk && !/Fehler|nicht/.test(s.envStatus)} warn={urlOk} /> {urlOk ? s.envStatus : 'Noch keine URL eingetragen.'}</div>
          </div>
        );
      case 'views':
        return (
          <div className={styles.setCard}>
            <Row label="Übersicht (Dashboard) als Startseite" description="Kennzahlen, Verlauf, Handlungsbedarf und Ranglisten">
              <Toggle checked={v.showOverview} onChange={(_e, c) => set({ showOverview: !!c })} />
            </Row>
            <Row label="Sicherheit & Freigaben" description="Oversharing, Gäste, Environments ohne DLP oder Sicherheitsgruppe …">
              <Toggle checked={v.showSecurity} onChange={(_e, c) => set({ showSecurity: !!c })} />
            </Row>
            <Row
              label="Copilot-Credits und Abrechnungspläne"
              description="Live aus der Power Platform API. Voraussetzung: Freigabe der Power Platform API (Skript im README) und die aktive Rolle „Power Platform Administrator“."
            >
              <Toggle checked={v.showCopilotCredits} onChange={(_e, c) => set({ showCopilotCredits: !!c })} />
            </Row>
            <Row label="Default Environment" description="Eigener Bereich – Details unter „Default Environment“">
              <Toggle checked={v.showDefaultEnvironment} onChange={(_e, c) => set({ showDefaultEnvironment: !!c })} />
            </Row>
            <Row label="Betroffen ab 15.12." description="Nicht verwaltete Developer-Environments – Details unter „Betroffen ab 15.12.“">
              <Toggle checked={v.showDevEnvironments} onChange={(_e, c) => set({ showDevEnvironments: !!c })} />
            </Row>
          </div>
        );
      case 'default':
        return (
          <div className={styles.setCard}>
            <Row label="Bereich anzeigen">
              <Toggle checked={v.showDefaultEnvironment} onChange={(_e, c) => set({ showDefaultEnvironment: !!c })} />
            </Row>
            <Dropdown
              label="Default Environment"
              options={s.envOptions.map(o => ({ key: o.key, text: o.text }))}
              selectedKey={v.defaultEnvironmentId}
              disabled={!v.showDefaultEnvironment}
              onChange={(_e, o) => set({ defaultEnvironmentId: o ? String(o.key) : '' })}
            />
            <div className={styles.setStatus}>{s.envStatus}</div>
          </div>
        );
      case 'dev':
        return (
          <div className={styles.setCard}>
            <Row label="Reiter „Betroffen ab 15.12.“ anzeigen">
              <Toggle checked={v.showDevEnvironments} onChange={(_e, c) => set({ showDevEnvironments: !!c })} />
            </Row>
            <Dropdown
              label="Spalte mit dem Managed-Status"
              options={s.managedOptions.map(o => ({ key: o.key, text: o.text }))}
              selectedKey={v.devEnvManagedField}
              disabled={!v.showDevEnvironments}
              onChange={(_e, o) => set({ devEnvManagedField: o ? String(o.key) : '' })}
            />
            <div className={styles.setStatus}>{s.managedInfo}</div>
          </div>
        );
      case 'mail':
        return (
          <>
            <div className={styles.setCard}>
              <Row
                label="Automatischen Versand anbieten"
                description="„Direkt senden“ und „Automatisch senden“ über Microsoft Graph aus dem eigenen Postfach. Voraussetzung: Berechtigung „Microsoft Graph – Mail.Send“."
              >
                <Toggle checked={v.notifyAutoSend} onChange={(_e, c) => set({ notifyAutoSend: !!c })} />
              </Row>
            </div>
            <div className={styles.setCard}>
              <div className={styles.setGrid2}>
                <Dropdown label="Anrede" options={[{ key: 'sie', text: 'Sie' }, { key: 'du', text: 'Du' }]} selectedKey={v.notifyTone}
                  onChange={(_e, o) => set({ notifyTone: (o && o.key === 'du') ? 'du' : 'sie' })} />
                <Dropdown label="Frist für die Rückmeldung" options={[7, 10, 14, 21].map(n => ({ key: n, text: `${n} Tage` }))} selectedKey={v.notifyReplyDays}
                  onChange={(_e, o) => set({ notifyReplyDays: o ? Number(o.key) : 14 })} />
              </div>
              <LazyText label="Link zur Microsoft-Information" value={v.notifyInfoUrl} placeholder={s.defaultInfoUrl}
                description="Leer = offizielle Ankündigung von Microsoft" validate={https} onCommit={x => set({ notifyInfoUrl: x })} />
              <LazyText label="Formular für die Rückmeldung (optional)" value={v.notifyFeedbackUrl}
                description="z. B. Microsoft Forms; {environment} im Link wird durch den Namen ersetzt. Ohne Formular antworten die Maker per E-Mail."
                validate={https} onCommit={x => set({ notifyFeedbackUrl: x })} />
              <div className={styles.setGrid2}>
                <LazyText label="CC (optional)" value={v.notifyCc} description="z. B. Team-Postfach"
                  validate={x => (!x || parseAddresses(x).length === x.split(/[;,\s]+/).filter(y => !!y).length ? '' : 'Bitte gültige E-Mail-Adressen (getrennt durch ;)')}
                  onCommit={x => set({ notifyCc: x })} />
                <LazyText label="Grußzeile" value={v.notifySignature} placeholder="Ihr Power-Platform-Team" onCommit={x => set({ notifySignature: x })} />
              </div>
            </div>
          </>
        );
      case 'actions':
        return (
          <>
            <div className={styles.setCard}>
              <Row label="Aktionen anbieten" description="Quarantäne, Löschen, Besitzer, Freigaben, Flows ein/aus, Environments, Solutions">
                <Toggle checked={v.showActions} onChange={(_e, c) => set({ showActions: !!c })} />
              </Row>
              <Row label="Löschen erst nach Quarantäne bzw. Ausschalten" description="Apps müssen in Quarantäne, Flows ausgeschaltet sein, bevor sie gelöscht werden können.">
                <Toggle checked={v.deleteRequiresStop} disabled={!v.showActions} onChange={(_e, c) => set({ deleteRequiresStop: !!c })} />
              </Row>
            </div>
            <div className={styles.setCard}>
              <div className={styles.setCardTitle}>Einrichtung</div>
              <Dropdown
                label="SharePoint-Gruppe mit Berechtigung"
                options={s.groupOptions.map(o => ({ key: o.key, text: o.text }))}
                selectedKey={v.actionsGroup}
                disabled={!v.showActions}
                onChange={(_e, o) => set({ actionsGroup: o ? String(o.key) : v.actionsGroup })}
              />
              <ul className={styles.setChecklist}>
                <li><Check ok={!!st && !!st.group} />{st && st.group ? `Gruppe „${st.group.title}“ vorhanden` : `Gruppe „${v.actionsGroup}“ fehlt`}</li>
                <li>
                  <Check ok={!!st && st.listExists && !!st.listUnique} warn={!!st && st.listExists} />
                  {!st || !st.listExists ? `Liste „${ACTION_LIST_TITLE}“ fehlt` : st.listUnique ? `Liste „${ACTION_LIST_TITLE}“ eingerichtet (eigene Berechtigungen)` : 'Liste vorhanden, erbt aber die Berechtigungen der Website'}
                </li>
                <li><Check ok={!!st && st.isMember} />{st && st.isMember ? 'Du bist Mitglied der Gruppe' : 'Du bist nicht Mitglied der Gruppe – Aktionen sind für dich ausgeblendet'}</li>
              </ul>
              <div className={styles.setButtons}>
                <DefaultButton
                  text={`Gruppe „${v.actionsGroup}“ anlegen`}
                  iconProps={{ iconName: 'AddGroup' }}
                  disabled={!v.showActions || !!s.actionBusy || !!(st && st.group)}
                  onClick={() => s.onActionSetup('group')}
                />
                <DefaultButton
                  text="Aktionsliste einrichten"
                  iconProps={{ iconName: 'Lock' }}
                  disabled={!v.showActions || !!s.actionBusy || !(st && st.group)}
                  onClick={() => s.onActionSetup('list')}
                />
              </div>
              {(s.actionBusy || s.actionMessage) && <div className={styles.setStatus}>{s.actionBusy || s.actionMessage}</div>}
            </div>
            <div className={styles.setNote}>
              <Icon iconName="Info" />
              <span>
                Ausgeführt werden die Aufträge vom Flow „CoE Cockpit | Aktion ausführen“ (Lösung im Ordner <code>flow/</code> des Projekts).
                Dort muss derselbe Gruppenname eingetragen sein; das Dienstkonto des Flows braucht „Bearbeiten“ auf der Liste.
              </span>
            </div>
          </>
        );
      case 'alm':
        return (
          <div className={styles.setCard}>
            <Dropdown
              label="Pipelines-Host-Environment"
              options={s.hostOptions.map(o => ({ key: o.key, text: o.text }))}
              selectedKey={v.pipelinesHostUrl}
              onChange={(_e, o) => set({ pipelinesHostUrl: o ? String(o.key) : '' })}
            />
            <div className={styles.setStatus}>Das Environment mit der App „Deployment Pipeline Configuration“. „Keine“ blendet die ALM-Pipelines aus.</div>
          </div>
        );
      case 'display':
        return (
          <>
            <div className={styles.setCard}>
              <div className={styles.setGrid2}>
                <LazyText label="Titel" value={v.title} onCommit={x => set({ title: x })} />
                <LazyText label="Untertitel" value={v.subtitle} onCommit={x => set({ subtitle: x })} />
              </div>
            </div>
            <div className={styles.setCard}>
              <Row label="Kennzahlen anzeigen" description="Kacheln über der Tabelle">
                <Toggle checked={v.showKpis} onChange={(_e, c) => set({ showKpis: !!c })} />
              </Row>
              <div className={styles.setGrid3}>
                <Dropdown label="Zeilenhöhe" options={[{ key: 'comfortable', text: 'Normal' }, { key: 'compact', text: 'Kompakt' }]} selectedKey={v.density}
                  onChange={(_e, o) => set({ density: o && o.key === 'compact' ? 'compact' : 'comfortable' })} />
                <Dropdown label="Einträge pro Seite" options={[10, 25, 50, 100].map(n => ({ key: n, text: String(n) }))} selectedKey={v.pageSize}
                  onChange={(_e, o) => set({ pageSize: o ? Number(o.key) : 25 })} />
                <Dropdown
                  label="Max. geladene Einträge"
                  options={[5000, 10000, 20000, 30000, 50000].concat([5000, 10000, 20000, 30000, 50000].indexOf(v.maxRecords) < 0 ? [v.maxRecords] : [])
                    .sort((a, b) => a - b).map(n => ({ key: n, text: n.toLocaleString('de-AT') }))}
                  selectedKey={v.maxRecords}
                  onChange={(_e, o) => set({ maxRecords: o ? Number(o.key) : 20000 })}
                />
              </div>
            </div>
          </>
        );
      default:
        return (
          <div className={styles.setCard}>
            <Row label={`${s.tableCount} Tabellen konfiguriert`} description="Tabellen hinzufügen, Spalten wählen, Reihenfolge ändern, JSON importieren/exportieren">
              <PrimaryButton text="Tabellen verwalten" iconProps={{ iconName: 'Table' }} onClick={s.onOpenTableEditor} />
            </Row>
          </div>
        );
    }
  };

  return (
    <Panel
      isOpen={props.isOpen}
      onDismiss={props.onDismiss}
      type={PanelType.custom}
      customWidth="920px"
      headerText="Webpart-Einstellungen"
      closeButtonAriaLabel="Schließen"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div className={styles.setFooter} style={props.themeVars}>
          <span className={styles.setFooterNote}><Icon iconName="Save" /> Änderungen gelten sofort und werden mit der Seite gespeichert („Erneut veröffentlichen“).</span>
          <span className={styles.setFooterButtons}>
            <DefaultButton text="Klassischer Eigenschaftenbereich" iconProps={{ iconName: 'SidePanel' }} onClick={s.onOpenPropertyPane} />
            <PrimaryButton text="Fertig" onClick={props.onDismiss} />
          </span>
        </div>
      )}
    >
      <div style={props.themeVars} className={`${styles.scope} ${styles.setShell}`}>
        <nav className={styles.setNav} aria-label="Bereiche">
          {sections.map(x => (
            <button
              key={x.key}
              type="button"
              className={`${styles.setNavItem} ${x.key === section ? styles.setNavItemActive : ''}`}
              aria-current={x.key === section ? 'page' : undefined}
              onClick={() => setSection(x.key)}
            >
              <Icon iconName={x.icon} className={styles.setNavIcon} />
              <span className={styles.setNavText}>
                <span className={styles.setNavLabel}>{x.label}</span>
                <span className={styles.setNavHint}>{x.hint}</span>
              </span>
              <span className={`${styles.setDot} ${x.tone === 'ok' ? styles.setDotOk : x.tone === 'warn' ? styles.setDotWarn : x.tone === 'error' ? styles.setDotErr : styles.setDotOff}`} aria-hidden="true" />
            </button>
          ))}
        </nav>
        <section className={styles.setContent} aria-label={current.label}>
          <h2 className={styles.setTitle}><Icon iconName={current.icon} /> {current.label}</h2>
          <p className={styles.setIntro}>{intro[section]}</p>
          {renderSection()}
        </section>
      </div>
    </Panel>
  );
};
