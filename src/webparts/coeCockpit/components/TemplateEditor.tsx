import * as React from 'react';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import {
  IMailTemplate, IMailTemplateOverrides, isCustomized, resolveTemplate, TEMPLATE_FIELDS, TEMPLATE_INFO, TemplateKey, usageKey
} from '../models/mailTemplates';
import { buildDevEnvMail, INotifySettings } from '../models/notification';
import { buildUsageMail, USAGE_TYPES } from '../models/usageMail';
import { DEV_FIELDS, RECOMMENDATION } from '../models/devEnvironments';

/**
 * Vorlagen-Editor (Webpart-Einstellungen → „E-Mail & Vorlagen“): Texte der Mails mit Platzhaltern und Live-Vorschau.
 * Gespeichert werden nur Abweichungen vom geerbten Text; „Auf Standard zurücksetzen“ entfernt sie.
 */
export interface ITemplateEditorProps {
  templates: IMailTemplateOverrides;
  /** Allgemeine Mail-Einstellungen (Anrede, Frist, CC, Grußzeile …) für die Vorschau. */
  notify: INotifySettings;
  accent: string;
  onChange: (templates: IMailTemplateOverrides) => void;
}

type FieldKey = keyof IMailTemplate;

/** Mehrzeiliges Feld, das beim Verlassen übernommen wird. */
const TemplateField: React.FC<{ label: string; value: string; multiline?: boolean; changed: boolean; onCommit: (v: string) => void }> = p => {
  const [draft, setDraft] = React.useState(p.value);
  React.useEffect(() => setDraft(p.value), [p.value]);
  return (
    <TextField
      label={p.label}
      ariaLabel={p.label}
      value={draft}
      multiline={p.multiline}
      autoAdjustHeight={p.multiline}
      description={p.changed ? 'Angepasst' : undefined}
      onChange={(_e, v) => setDraft(v || '')}
      onBlur={() => { if (draft !== p.value) { p.onCommit(draft); } }}
      onKeyDown={e => { if (e.key === 'Enter' && !p.multiline) { p.onCommit(draft); } }}
    />
  );
};

const SAMPLE_DEV = {
  admin_environmentid: '00000000-0000-0000-0000-000000000001',
  admin_environmentname: '00000000-0000-0000-0000-000000000001',
  admin_displayname: 'Dev – Max Muster',
  [DEV_FIELDS.apps]: 3,
  [DEV_FIELDS.flows]: 5,
  [DEV_FIELDS.lastActivity]: new Date(Date.now() - 12 * 86400000).toISOString(),
  [DEV_FIELDS.recommendation]: RECOMMENDATION.inUse,
  [DEV_FIELDS.makerEmails]: 'max.muster@contoso.com'
};

export const TemplateEditor: React.FC<ITemplateEditorProps> = props => {
  const [kind, setKind] = React.useState<TemplateKey>('usage');
  /** Bei „Nachfrage Verwendung“: '' = gemeinsame Vorlage, sonst Tabelle (eigene Fassung je Typ) */
  const [scope, setScope] = React.useState<string>('');
  const feedback = !!(props.notify.feedbackUrl || '').trim();
  const storeKey = kind === 'usage' && scope ? usageKey(scope) : kind;
  const own = props.templates[storeKey] || {};
  // Geerbt: Standard (bzw. bei einer Typ-Fassung die gemeinsame Vorlage)
  const inherited = kind === 'usage' && scope
    ? resolveTemplate('usage', props.notify.formal, feedback, props.templates)
    : resolveTemplate(kind, props.notify.formal, feedback, {});
  const effective = kind === 'usage' ? resolveTemplate('usage', props.notify.formal, feedback, props.templates, scope || undefined)
    : resolveTemplate(kind, props.notify.formal, feedback, props.templates);

  const commit = (field: FieldKey, raw: string): void => {
    const next: Partial<IMailTemplate> = { ...own };
    if (field === 'options') {
      const list = raw.split(/\r?\n/).map(x => x.trim()).filter(x => !!x);
      if (!list.length || list.join('\n') === inherited.options.join('\n')) {
        delete next.options;
      } else {
        next.options = list;
      }
    } else {
      const v = raw.trim();
      if (!v || v === inherited[field]) {
        delete next[field];
      } else {
        (next as { [f: string]: string })[field] = v;
      }
    }
    const all: IMailTemplateOverrides = { ...props.templates };
    if (Object.keys(next).length) {
      all[storeKey] = next;
    } else {
      delete all[storeKey];
    }
    props.onChange(all);
  };

  const reset = (): void => {
    const all: IMailTemplateOverrides = { ...props.templates };
    delete all[storeKey];
    props.onChange(all);
  };

  const preview = React.useMemo(() => {
    const s: INotifySettings = { ...props.notify, templates: props.templates };
    if (kind === 'devenv') {
      return buildDevEnvMail(SAMPLE_DEV, s, new Date(), props.accent);
    }
    const type = scope ? USAGE_TYPES[scope] : USAGE_TYPES.admin_app;
    return buildUsageMail({
      logicalName: scope || 'admin_app',
      typeLabel: type ? type.label : 'App',
      name: type && type.label === 'Environment' ? 'Contoso Vertrieb' : `Beispiel-${type ? type.label : 'App'}`,
      environment: 'Contoso Produktion',
      ownerName: 'Max Muster',
      ownerEmail: 'max.muster@contoso.com',
      lastLabel: 'Zuletzt verwendet',
      lastValue: new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(new Date(Date.now() - 95 * 86400000)),
      openUrl: 'https://make.powerapps.com/',
      openLabel: 'In Power Apps öffnen'
    }, s, new Date(), props.accent);
  }, [kind, scope, props.templates, props.notify, props.accent]);

  const info = TEMPLATE_INFO[kind];
  const typeOptions = [{ key: '', text: 'Alle Komponenten (gemeinsame Vorlage)' }].concat(
    Object.keys(USAGE_TYPES).map(k => ({ key: k, text: `nur ${USAGE_TYPES[k].label}${isCustomized(usageKey(k), props.templates) ? ' – angepasst' : ''}` })));

  return (
    <div className={styles.setCard}>
      <div className={styles.setCardTitle}>Vorlagen</div>
      <div className={styles.ccSegments} role="group" aria-label="Vorlage">
        {(['usage', 'devenv'] as TemplateKey[]).map(k => (
          <button
            key={k}
            type="button"
            className={`${styles.ccSegment} ${kind === k ? styles.ccSegmentActive : ''}`}
            aria-pressed={kind === k}
            onClick={() => { setKind(k); setScope(''); }}
          >
            {TEMPLATE_INFO[k].title}
            {(isCustomized(k, props.templates) || (k === 'usage' && Object.keys(props.templates).some(x => /^usage:/.test(x) && isCustomized(x, props.templates))))
              && <span className={styles.ccSegmentCount}>angepasst</span>}
          </button>
        ))}
      </div>
      <div className={styles.setRowDesc}>{info.description}</div>
      {kind === 'usage' && (
        <Dropdown
          label="Gilt für"
          options={typeOptions}
          selectedKey={scope}
          onChange={(_e, o) => setScope(o ? String(o.key) : '')}
        />
      )}
      {kind === 'usage' && scope && (
        <div className={styles.setNote}>
          <Icon iconName="Info" />
          <span>Eigene Fassung nur für {USAGE_TYPES[scope].label}. Felder, die du hier nicht änderst, kommen aus der gemeinsamen Vorlage.</span>
        </div>
      )}
      <div className={styles.tplPlaceholders}>
        <span className={styles.setRowDesc}>Platzhalter:</span>
        {info.placeholders.map(ph => <span key={ph.name} className={styles.chip} title={ph.text}>{ph.name}</span>)}
      </div>
      <div className={styles.tplLayout}>
        <div className={styles.tplFields}>
          {TEMPLATE_FIELDS.map(f => (
            <TemplateField
              key={`${storeKey}:${f.key}`}
              label={f.label}
              multiline={f.multiline}
              value={f.list ? effective.options.join('\n') : String(effective[f.key])}
              changed={own[f.key] !== undefined}
              onCommit={v => commit(f.key, v)}
            />
          ))}
          <div className={styles.setButtons}>
            <DefaultButton text={scope ? 'Typ-Fassung entfernen' : 'Auf Standard zurücksetzen'} iconProps={{ iconName: 'Undo' }} disabled={!isCustomized(storeKey, props.templates)} onClick={reset} />
          </div>
        </div>
        <div className={styles.tplPreview}>
          <div className={styles.notifyPreviewLabel}>Vorschau (Beispieldaten)</div>
          <div className={styles.tplSubject}><b>Betreff:</b> {preview.subject}</div>
          <iframe
            className={styles.notifyPreview}
            title="Vorschau der Vorlage"
            sandbox=""
            srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:8px;background:#ffffff;">${preview.html}</body></html>`}
          />
        </div>
      </div>
    </div>
  );
};
