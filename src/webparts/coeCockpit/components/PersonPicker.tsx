import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { DataverseService } from '../services/DataverseService';
import { displayValue, isMailAddress, odataLikeLiteral } from '../utils/format';

/**
 * Person auswählen: Vorschläge aus dem CoE (Power Platform User), sonst E-Mail-Adresse
 * (das Webpart sucht sie dann über Microsoft Graph). Gruppen sind für Besitzer nicht erlaubt.
 */
export interface IPickedPerson { id: string; name: string; sub: string }

export interface IPersonPickerProps {
  label: string;
  dataverse: DataverseService | undefined;
  value: IPickedPerson | undefined;
  onChange: (p: IPickedPerson | undefined) => void;
}

export const PersonPicker: React.FC<IPersonPickerProps> = props => {
  const [query, setQuery] = React.useState<string>('');
  const [suggestions, setSuggestions] = React.useState<IPickedPerson[]>([]);
  const picked = props.value;

  React.useEffect(() => {
    const q = query.trim();
    if (!props.dataverse || q.length < 2 || (picked && picked.name === q)) {
      setSuggestions([]);
      return;
    }
    const lit = odataLikeLiteral(q);
    const timer = window.setTimeout(() => {
      props.dataverse!.getRecords({
        key: '_people', title: 'Power Platform User', logicalName: 'admin_powerplatformuser', icon: 'People', strictFilter: true,
        filter: `(contains(admin_displayname,'${lit}') or contains(admin_userprincipalname,'${lit}')) and admin_type eq 'User'`,
        defaultSort: { field: 'admin_displayname', desc: false },
        columns: [
          { field: 'admin_displayname', label: 'Name', type: 'text' },
          { field: 'admin_userprincipalname', label: 'UPN', type: 'text' }
        ]
      }, 8).then(r => setSuggestions(r.records.map(x => ({
        id: String(x[r.meta.primaryId] || ''),
        name: displayValue(x, 'admin_displayname'),
        sub: displayValue(x, 'admin_userprincipalname')
      }))), () => setSuggestions([]));
    }, 300);
    return () => window.clearTimeout(timer);
  }, [query]);

  const typed = query.trim();
  const mail = !picked && isMailAddress(typed) ? typed : '';

  return (
    <div>
      <TextField
        label={props.label}
        placeholder="Name oder E-Mail-Adresse"
        value={picked ? picked.name : query}
        onChange={(_e, v) => {
          const text = v || '';
          setQuery(text);
          const asMail = isMailAddress(text.trim()) ? { id: text.trim(), name: text.trim(), sub: 'per E-Mail-Adresse (Microsoft Graph)' } : undefined;
          props.onChange(asMail);
        }}
        description={mail ? 'Nicht im CoE – die Person wird über die E-Mail-Adresse gesucht (Microsoft Graph).' : 'Vorschläge aus dem CoE (Power Platform User). Personen außerhalb des CoE per E-Mail-Adresse.'}
      />
      {suggestions.length > 0 && (!picked || picked.name !== query) && (
        <div className={styles.actionSuggest} role="listbox" aria-label="Vorschläge">
          {suggestions.map(s => (
            <button key={s.id} type="button" role="option" aria-selected={false} className={styles.searchItem}
              onClick={() => { props.onChange(s); setQuery(s.name); setSuggestions([]); }}>
              <span className={styles.searchItemTitle}><Icon iconName="Contact" /> {s.name}</span>
              {s.sub && <span className={styles.searchItemSub}>{s.sub}</span>}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
