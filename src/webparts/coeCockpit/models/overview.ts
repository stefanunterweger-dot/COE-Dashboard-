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

/**
 * Wie monthlyCounts, aber mit mehreren möglichen Datumsfeldern (CoE-Versionen unterscheiden sich): das erste Feld,
 * das Dataverse kennt, wird verwendet. Liefert auch das verwendete Feld (z. B. „createdon“ = erstmals im CoE erfasst).
 */
export async function monthlyCountsAny(service: DataverseService, table: ITableConfig, fields: string[], months: IMonth[]): Promise<{ values: number[]; field: string }> {
  let last: unknown;
  for (const field of fields) {
    try {
      return { values: await monthlyCounts(service, table, field, months), field };
    } catch (e) {
      last = e;
      const msg = e instanceof Error ? e.message : String(e);
      if (!/property|attribute|could not find|not found/i.test(msg)) {
        throw e;
      }
    }
  }
  throw last || new Error('Kein Datumsfeld gefunden.');
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

/**
 * Meistgenutzte Apps. Quelle 1: CoE-Spalte „Launches in the past 30 days“.
 * Ist sie leer, werden die App-Starts im CoE-Audit-Log gezählt (30 Tage, bei sehr vielen Einträgen 7 Tage).
 */
export async function topAppsByLaunches(service: DataverseService, apps: ITableConfig, limit: number): Promise<{
  records: IRecord[]; primaryId: string; valueField: string; source: 'field' | 'audit'; days: number;
}> {
  const viaField = await topRecords(service, apps, 'admin_launchesinthepast30days', limit, ['admin_launchesinthepast30days']).catch(() => undefined);
  if (viaField && viaField.records.some(r => numberOf(r, 'admin_launchesinthepast30days') > 0)) {
    return { records: viaField.records, primaryId: viaField.primaryId, valueField: 'admin_launchesinthepast30days', source: 'field', days: 30 };
  }
  for (const days of [30, 7]) {
    let counts: { [appId: string]: number };
    try {
      counts = await service.getGroupedCounts('admin_auditlog', '_admin_applookup_value',
        `admin_operation eq 'LaunchPowerApp' and admin_creationtime ge ${daysAgoIso(days)}`, undefined, true);
    } catch {
      continue;
    }
    const ids = Object.keys(counts).filter(id => /^[0-9a-f-]{36}$/i.test(id)).sort((a, b) => counts[b] - counts[a]).slice(0, limit);
    if (!ids.length) {
      break;
    }
    const meta = await service.getEntityMeta(apps.logicalName);
    const res = await service.getRecords({
      ...apps, key: `_top_${apps.key}`, strictFilter: true,
      filter: combineFilters(apps.filter, ids.map(id => `${meta.primaryId} eq ${id}`).join(' or '))
    }, limit);
    const records = res.records
      .map(r => ({ ...r, coe_launches: counts[String(r[meta.primaryId]).toLowerCase()] || counts[String(r[meta.primaryId])] || 0 }))
      .sort((a, b) => numberOf(b, 'coe_launches') - numberOf(a, 'coe_launches'));
    return { records, primaryId: meta.primaryId, valueField: 'coe_launches', source: 'audit', days };
  }
  return { records: [], primaryId: 'admin_appid', valueField: 'admin_launchesinthepast30days', source: 'field', days: 30 };
}

/** Environments komplett (bis 50.000 – mehr gibt es in keinem Tenant) – für Typ-Verteilung und Rangliste. */
export async function environmentsOverview(service: DataverseService, table: ITableConfig): Promise<{
  records: IRecord[]; primaryId: string; byType: { label: string; value: number }[];
}> {
  const res = await service.getRecords(table, 50000);
  if (res.truncated) {
    throw new Error('Mehr als 50.000 Environments – Typverteilung wäre unvollständig.');
  }
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
