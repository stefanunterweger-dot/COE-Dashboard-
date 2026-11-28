import { IColumnConfig, IRecord } from '../models/ITableConfig';
import { displayValue, formatNumber, splitList, toDate } from './format';

export type DatePreset = 'today' | 'last7' | 'last30' | 'last90' | 'thisYear' | 'older90' | 'older365' | 'empty' | 'notEmpty';

/** Filter einer Spalte: entweder Werteliste oder Zeitraum (Datumsspalten). */
export interface IColumnFilter {
  /** Ausgewählte Anzeigewerte ('' = leer). */
  values?: string[];
  date?: DatePreset;
}

export interface IFilterOption {
  value: string;
  label: string;
  count: number;
}

export const DATE_PRESETS: { key: DatePreset; label: string }[] = [
  { key: 'today', label: 'Heute' },
  { key: 'last7', label: 'Letzte 7 Tage' },
  { key: 'last30', label: 'Letzte 30 Tage' },
  { key: 'last90', label: 'Letzte 90 Tage' },
  { key: 'thisYear', label: 'Dieses Jahr' },
  { key: 'older90', label: 'Älter als 90 Tage' },
  { key: 'older365', label: 'Älter als 1 Jahr' },
  { key: 'notEmpty', label: 'Mit Datum' },
  { key: 'empty', label: 'Ohne Datum' }
];

const DAY = 24 * 60 * 60 * 1000;

export function isDateColumn(column: IColumnConfig): boolean {
  return column.type === 'date' || column.type === 'datetime';
}

export function isFilterActive(filter: IColumnFilter | undefined): boolean {
  return !!filter && (!!filter.date || (!!filter.values && filter.values.length > 0));
}

function matchesDate(value: unknown, preset: DatePreset): boolean {
  const d = toDate(value);
  if (preset === 'empty') {
    return !d;
  }
  if (!d) {
    return false;
  }
  if (preset === 'notEmpty') {
    return true;
  }
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const t = d.getTime();
  switch (preset) {
    case 'today': return t >= startOfToday;
    case 'last7': return t >= startOfToday - 6 * DAY;
    case 'last30': return t >= startOfToday - 29 * DAY;
    case 'last90': return t >= startOfToday - 89 * DAY;
    case 'thisYear': return d.getFullYear() === now.getFullYear();
    case 'older90': return t < startOfToday - 89 * DAY;
    case 'older365': return t < startOfToday - 364 * DAY;
    default: return true;
  }
}

/** Einzelwerte einer Listen-Spalte („A; B; C“ bzw. „A, B, C“ → [A, B, C]). */
export function listItems(text: string): string[] {
  return splitList(text);
}

export function matchesFilter(record: IRecord, column: IColumnConfig, filter: IColumnFilter): boolean {
  if (filter.date) {
    return matchesDate(record[column.field], filter.date);
  }
  if (filter.values && filter.values.length > 0) {
    const text = displayValue(record, column.field);
    if (column.type === 'list') {
      // Treffer, wenn einer der gewählten Werte in der Liste vorkommt (bzw. leer gewählt und Liste leer)
      const items = listItems(text);
      return items.length === 0 ? filter.values.indexOf('') >= 0 : items.some(i => filter.values!.indexOf(i) >= 0);
    }
    return filter.values.indexOf(text) >= 0;
  }
  return true;
}

/** Eindeutige Werte einer Spalte mit Anzahl, sortiert. */
export function buildOptions(records: IRecord[], column: IColumnConfig): IFilterOption[] {
  const counts: { [v: string]: number } = {};
  records.forEach(r => {
    const v = displayValue(r, column.field);
    if (column.type === 'list') {
      // jede Person/jeder Wert einzeln (Anzahl = Zeilen, in denen er vorkommt)
      const items = listItems(v);
      (items.length ? items : ['']).forEach(i => { counts[i] = (counts[i] || 0) + 1; });
      return;
    }
    counts[v] = (counts[v] || 0) + 1;
  });
  return Object.keys(counts)
    .map(v => ({ value: v, label: v === '' ? '(leer)' : v, count: counts[v] }))
    .sort((a, b) => {
      if (a.value === '') { return 1; }
      if (b.value === '') { return -1; }
      if (column.type === 'number') {
        return Number(a.value) - Number(b.value);
      }
      return a.label.localeCompare(b.label, 'de');
    });
}

/** Anzahl Datensätze je Zeitraum. */
export function buildDateCounts(records: IRecord[], column: IColumnConfig): { [preset: string]: number } {
  const out: { [preset: string]: number } = {};
  DATE_PRESETS.forEach(p => {
    out[p.key] = records.filter(r => matchesDate(r[column.field], p.key)).length;
  });
  return out;
}

/** Kurzbeschreibung für den Filter-Chip in der Toolbar. */
export function describeFilter(column: IColumnConfig, filter: IColumnFilter): string {
  if (filter.date) {
    const p = DATE_PRESETS.filter(x => x.key === filter.date)[0];
    return `${column.label}: ${p ? p.label : filter.date}`;
  }
  const values = filter.values || [];
  if (values.length === 0) {
    return column.label;
  }
  const first = values[0] === '' ? '(leer)' : values[0];
  return values.length === 1 ? `${column.label}: ${first}` : `${column.label}: ${first} +${formatNumber(values.length - 1)}`;
}
