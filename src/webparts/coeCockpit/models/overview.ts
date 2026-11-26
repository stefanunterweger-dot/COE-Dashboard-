import { IRecord, ITableConfig } from './ITableConfig';
import { DataverseService, environmentType, ICountResult } from '../services/DataverseService';
import { combineFilters, displayValue } from '../utils/format';

/**
 * Daten der Übersichtsseite – ausschließlich Zählungen, Gruppierungen und kurze Ranglisten
 * aus den CoE-Tabellen (nur lesend, auch bei sehr großen Tabellen schnell).
 */

export const OVERVIEW_KEY = 'overview';
export const MONTHS = 12;

const MONTH_SHORT = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];
const MONTH_LONG = ['Jänner', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];

export interface IMonth {
  start: Date;
  end: Date;
  short: string;
  long: string;
}

/** Die letzten 12 Monate inkl. des laufenden. */
export function lastMonths(count: number = MONTHS): IMonth[] {
  const now = new Date();
  const out: IMonth[] = [];
  for (let i = count - 1; i >= 0; i--) {
    const start = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const end = new Date(now.getFullYear(), now.getMonth() - i + 1, 1);
    out.push({
      start, end,
      short: start.getMonth() === 0 || i === count - 1 || i === 0 ? `${MONTH_SHORT[start.getMonth()]} ${String(start.getFullYear()).slice(2)}` : MONTH_SHORT[start.getMonth()],
      long: `${MONTH_LONG[start.getMonth()]} ${start.getFullYear()}`
    });
  }
  return out;
}

const iso = (d: Date): string => d.toISOString().replace(/\.\d{3}Z$/, 'Z');

/** Anzahl neu erstellter Datensätze je Monat (eine Zählung pro Monat, exakt bis 50.000). */
export function monthlyCounts(service: DataverseService, table: ITableConfig, dateField: string, months: IMonth[]): Promise<number[]> {
  return Promise.all(months.map(m => service
    .getCountWithBase(table, `${dateField} ge ${iso(m.start)} and ${dateField} lt ${iso(m.end)}`)
    .then(c => c.value)));
}

export function count(service: DataverseService, table: ITableConfig, extra?: string): Promise<ICountResult> {
  return service.getCountWithBase(table, extra);
}

/** Anzahl je Wert einer Spalte (z. B. Apps je Typ), absteigend sortiert. */
export async function distribution(
  service: DataverseService, table: ITableConfig, field: string, labels?: { [raw: string]: string }
): Promise<{ label: string; value: number }[]> {
  const counts = await service.getGroupedCounts(table.logicalName, field, table.filter);
  return Object.keys(counts)
    .map(k => ({ label: (labels && labels[k]) || k || '(leer)', value: counts[k] }))
    .sort((a, b) => b.value - a.value);
}

/** Kurze Rangliste: die ersten N Datensätze nach einer Zahlenspalte. */
export async function topRecords(
  service: DataverseService, table: ITableConfig, sortField: string, limit: number, extraFields: string[] = []
): Promise<{ records: IRecord[]; table: ITableConfig; primaryId: string }> {
  const known = table.columns.map(c => c.field);
  const query: ITableConfig = {
    ...table,
    key: `_top_${table.key}`,
    defaultSort: { field: sortField, desc: true },
    columns: table.columns.concat(extraFields.filter(f => known.indexOf(f) < 0).map(f => ({ field: f, label: f, type: 'number' as const, hidden: true })))
  };
  const res = await service.getRecords(query, limit);
  return { records: res.records.slice(0, limit), table, primaryId: res.meta.primaryId };
}

/** Environments (kleine Tabelle) komplett – für Typ-Verteilung und Rangliste. */
export async function environmentsOverview(service: DataverseService, table: ITableConfig): Promise<{
  records: IRecord[]; primaryId: string; byType: { label: string; value: number }[];
}> {
  const res = await service.getRecords(table, 5000);
  const counts: { [t: string]: number } = {};
  res.records.forEach(r => {
    const t = environmentType(r) || '(unbekannt)';
    counts[t] = (counts[t] || 0) + 1;
  });
  return {
    records: res.records,
    primaryId: res.meta.primaryId,
    byType: Object.keys(counts).map(k => ({ label: k, value: counts[k] })).sort((a, b) => b.value - a.value)
  };
}

/** Datum vor x Tagen als OData-Literal. */
export function daysAgoIso(days: number): string {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - days);
  return iso(d);
}

export function numberOf(r: IRecord, field: string): number {
  const v = Number(r[field]);
  return isNaN(v) ? 0 : v;
}

export function textOf(r: IRecord, field: string): string {
  return displayValue(r, field);
}

export { combineFilters };
