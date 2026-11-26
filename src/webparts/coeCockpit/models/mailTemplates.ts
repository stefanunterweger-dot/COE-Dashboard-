/**
 * Pflegbare Texte der E-Mails (Webpart-Einstellungen → „E-Mail & Vorlagen“).
 * Gespeichert werden nur Abweichungen vom Standard; leere Felder bedeuten „Standardtext verwenden“.
 * Platzhalter in geschweiften Klammern, z. B. {Name}, werden beim Erstellen der Mail ersetzt.
 */
export type TemplateKey = 'devenv' | 'usage';
/** Gespeicherte Vorlagen: „devenv“, „usage“ (gemeinsam) und optional „usage:<Tabelle>“ (z. B. usage:admin_flow) je Komponententyp. */
export type StoredTemplateKey = TemplateKey | string;

export interface IMailTemplate {
  /** Betreff */
  subject: string;
  /** Überschrift im farbigen Kopfbalken */
  headline: string;
  /** Anrede, z. B. „Guten Tag,“ */
  greeting: string;
  /** Einleitung */
  intro: string;
  /** Bitte um Rückmeldung (über den Antwortmöglichkeiten) */
  question: string;
  /** Nummerierte Antwortmöglichkeiten */
  options: string[];
  /** Hinweis im grauen Kasten (z. B. „Ohne Rückmeldung …“) */
  note: string;
  /** Grußformel (die Grußzeile/Absender kommt aus den allgemeinen Einstellungen) */
  closing: string;
}

export type IMailTemplateOverrides = { [key: string]: Partial<IMailTemplate> | undefined };

/** Schlüssel der typspezifischen Fassung von „Nachfrage Verwendung“. */
export function usageKey(logicalName: string): string {
  return `usage:${logicalName}`;
}

export const TEMPLATE_FIELDS: { key: keyof IMailTemplate; label: string; multiline?: boolean; list?: boolean }[] = [
  { key: 'subject', label: 'Betreff' },
  { key: 'headline', label: 'Überschrift (Kopfbalken)' },
  { key: 'greeting', label: 'Anrede' },
  { key: 'intro', label: 'Einleitung', multiline: true },
  { key: 'question', label: 'Bitte um Rückmeldung', multiline: true },
  { key: 'options', label: 'Antwortmöglichkeiten (eine pro Zeile)', multiline: true, list: true },
  { key: 'note', label: 'Hinweis (grauer Kasten)', multiline: true },
  { key: 'closing', label: 'Grußformel' }
];

export const TEMPLATE_INFO: { [K in TemplateKey]: { title: string; description: string; placeholders: { name: string; text: string }[] } } = {
  devenv: {
    title: 'Betroffen ab 15.12. (Developer-Environments)',
    description: 'Mail an die Maker nicht verwalteter Developer-Environments. Fristen, Steckbrief des Environments, Sicherungs-Tipp und der Link zur Microsoft-Ankündigung werden automatisch ergänzt.',
    placeholders: [
      { name: '{Name}', text: 'Name des Environments' },
      { name: '{Frist}', text: 'Datum der Rückmeldefrist' },
      { name: '{Sperre}', text: 'Datum der Sperre (15.12.2026)' },
      { name: '{Löschung}', text: 'Datum der Löschung (15.01.2027)' },
      { name: '{Änderung}', text: 'Satz zur Änderung durch Microsoft (passt sich nach der Sperre an)' },
      { name: '{Zustand}', text: '„gesperrt“ bzw. nach der Sperre „gelöscht“' },
      { name: '{ZustandMitDatum}', text: '„am 15.12.2026 gesperrt“ bzw. nach der Sperre „gelöscht“' }
    ]
  },
  usage: {
    title: 'Nachfrage Verwendung',
    description: 'Mail an den Besitzer einer App, eines Flows, Bots, Environments … mit der Frage, ob die Komponente noch benötigt wird. Ein Steckbrief (Typ, Name, Environment, letzte Verwendung) und ein Link zum Öffnen werden automatisch ergänzt.',
    placeholders: [
      { name: '{Name}', text: 'Name der Komponente' },
      { name: '{Typ}', text: 'Art, z. B. App, Flow, Bot, Environment' },
      { name: '{Environment}', text: 'Environment der Komponente' },
      { name: '{Besitzer}', text: 'Name des Besitzers' },
      { name: '{Frist}', text: 'Datum der Rückmeldefrist' }
    ]
  }
};

/** Standardtexte (Sie bzw. Du). feedback = ein Rückmelde-Formular ist eingestellt. */
export function defaultTemplate(key: TemplateKey, formal: boolean, feedback: boolean): IMailTemplate {
  const S = formal;
  const ask = feedback
    ? (S ? 'Bitte nutzen Sie das Formular oder antworten Sie auf diese E-Mail mit der passenden Zahl:' : 'Bitte nutze das Formular oder antworte auf diese E-Mail mit der passenden Zahl:')
    : (S ? 'Bitte antworten Sie einfach auf diese E-Mail mit der passenden Zahl:' : 'Bitte antworte einfach auf diese E-Mail mit der passenden Zahl:');
  if (key === 'devenv') {
    return {
      subject: 'Rückmeldung bis {Frist}: Developer-Environment „{Name}“ wird {Zustand}',
      headline: 'Developer-Environment „{Name}“ wird {ZustandMitDatum}',
      greeting: S ? 'Guten Tag,' : 'Hallo,',
      intro: (S ? 'Sie sind' : 'du bist') + ' als Maker im Power-Platform-Developer-Environment „{Name}“ eingetragen. {Änderung}',
      question: ask,
      options: [
        'Wird weiter benötigt – das Environment wird auf Managed umgestellt. Wer dort Apps oder Flows nutzt, braucht dann eine Premium-Lizenz.',
        'Kann gelöscht werden.',
        S ? 'Inhalte kommen in ein anderes Environment – bitte nennen Sie das Ziel.' : 'Inhalte kommen in ein anderes Environment – bitte nenne das Ziel.'
      ],
      note: 'Ohne Rückmeldung unternehmen wir nichts – Microsoft sperrt und löscht das Environment dann automatisch.',
      closing: S ? 'Freundliche Grüße' : 'Viele Grüße'
    };
  }
  return {
    subject: 'Wird „{Name}“ noch benötigt? Bitte um Rückmeldung bis {Frist}',
    headline: '{Typ} „{Name}“: noch in Verwendung?',
    greeting: S ? 'Guten Tag,' : 'Hallo,',
    intro: (S ? 'Sie sind' : 'du bist') + ' im Power-Platform-Inventar als verantwortliche Person für die unten genannte Komponente eingetragen. '
      + 'Im Rahmen unserer regelmäßigen Bereinigung möchten wir wissen, ob sie noch benötigt wird.',
    question: ask,
    options: [
      'Wird weiter benötigt.',
      'Wird nicht mehr benötigt und kann gelöscht werden.',
      S ? 'Ich bin nicht (mehr) zuständig – bitte nennen Sie uns die richtige Person.' : 'Ich bin nicht (mehr) zuständig – bitte nenne uns die richtige Person.'
    ],
    note: 'Ohne Rückmeldung bis {Frist} gehen wir davon aus, dass die Komponente nicht mehr benötigt wird. Sie wird dann zunächst außer Betrieb genommen und später gelöscht.',
    closing: S ? 'Freundliche Grüße' : 'Viele Grüße'
  };
}

function applyOverrides(base: IMailTemplate, o: Partial<IMailTemplate> | undefined): IMailTemplate {
  const out: IMailTemplate = { ...base };
  if (!o) {
    return out;
  }
  (Object.keys(base) as (keyof IMailTemplate)[]).forEach(k => {
    const v = o[k];
    if (k === 'options') {
      const list = Array.isArray(v) ? (v as string[]).map(x => String(x).trim()).filter(x => !!x) : [];
      if (list.length) {
        out.options = list;
      }
    } else if (typeof v === 'string' && v.trim()) {
      (out as unknown as { [f: string]: string })[k] = v.trim();
    }
  });
  return out;
}

/**
 * Standard + gespeicherte Abweichungen (leere Felder → Standard).
 * Bei „usage“ mit Tabelle: Standard ← gemeinsame Vorlage ← Fassung für diesen Typ.
 */
export function resolveTemplate(key: TemplateKey, formal: boolean, feedback: boolean, overrides?: IMailTemplateOverrides, logicalName?: string): IMailTemplate {
  let t = applyOverrides(defaultTemplate(key, formal, feedback), overrides && overrides[key]);
  if (key === 'usage' && logicalName) {
    t = applyOverrides(t, overrides && overrides[usageKey(logicalName)]);
  }
  return t;
}

/** Ist ein Feld der Vorlage gegenüber dem Standard geändert? */
export function isCustomized(key: StoredTemplateKey, overrides?: IMailTemplateOverrides): boolean {
  const o = overrides && overrides[key];
  return !!o && Object.keys(o).some(k => {
    const v = (o as { [f: string]: unknown })[k];
    return Array.isArray(v) ? v.some(x => !!String(x).trim()) : typeof v === 'string' && !!v.trim();
  });
}

/** Platzhalter ersetzen: {Name} → Wert. Unbekannte Platzhalter bleiben stehen. */
export function fillText(text: string, vars: { [name: string]: string }): string {
  return text.replace(/\{([A-Za-zÄÖÜäöüß]+)\}/g, (m: string, name: string) => (Object.prototype.hasOwnProperty.call(vars, name) ? vars[name] : m));
}
