/* Simulierte Dataverse Web API (v9.2) – beantwortet genau die Aufrufe des DataverseService. */
import { DATABASES, IDatabase, ITableDef } from './data';

const FORMATTED = '@OData.Community.Display.V1.FormattedValue';
const LOOKUP_NAME = '@Microsoft.Dynamics.CRM.lookuplogicalname';

export interface IMockResponse {
  ok: boolean;
  status: number;
  json(): Promise<unknown>;
  text(): Promise<string>;
}

class ODataError extends Error {
  constructor(public status: number, message: string) { super(message); }
}

function respond(status: number, body: unknown): IMockResponse {
  const text = JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    json: () => Promise.resolve(JSON.parse(text)),
    text: () => Promise.resolve(text)
  };
}

function findDb(host: string): IDatabase | undefined {
  return DATABASES.filter(d => d.host.toLowerCase() === host.toLowerCase())[0];
}

function tableByLogical(db: IDatabase, name: string): ITableDef {
  const t = db.tables[name.toLowerCase()];
  if (!t) {
    throw new ODataError(404, `Could not find an entity with name '${name}'. Entity '${name}' does not exist.`);
  }
  return t;
}

function tableBySet(db: IDatabase, set: string): ITableDef {
  for (const k of Object.keys(db.tables)) {
    if (db.tables[k].entitySet === set) {
      return db.tables[k];
    }
  }
  throw new ODataError(404, `Resource not found for the segment '${set}'.`);
}

/** Feldname ($select/$filter) → Attribut. */
function attrFor(t: ITableDef, field: string): { name: string; type: string; target?: string; options?: { [v: number]: string } } {
  const lookup = /^_(.+)_value$/.exec(field);
  const a = t.attrs.filter(x => (lookup ? x.name === lookup[1] && x.type === 'Lookup' : x.name === field && x.type !== 'Lookup'))[0];
  if (!a) {
    throw new ODataError(400, `Could not find a property named '${field}' on type 'Microsoft.Dynamics.CRM.${t.logicalName}'.`);
  }
  return a;
}

const dateFmt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });

/** Datensatz mit Annotationen wie Dataverse (Prefer: odata.include-annotations="*"). */
function project(t: ITableDef, row: Record<string, unknown>, fields: string[] | undefined): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const list = fields || t.attrs.map(a => (a.type === 'Lookup' ? `_${a.name}_value` : a.name));
  list.forEach(f => {
    const a = attrFor(t, f);
    const v = row[f] === undefined ? null : row[f];
    out[f] = v;
    if (v === null) {
      return;
    }
    if (a.type === 'Lookup' && a.target) {
      const target = lookupTarget(a.target);
      const ref = target ? target.rows.filter(r => r[target.primaryId] === v)[0] : undefined;
      out[f + FORMATTED] = ref ? String(ref[target!.primaryName]) : String(v);
      out[f + LOOKUP_NAME] = a.target;
    } else if (a.type === 'Boolean') {
      out[f + FORMATTED] = v ? 'Ja' : 'Nein';
    } else if (a.type === 'DateTime') {
      out[f + FORMATTED] = dateFmt.format(new Date(String(v)));
    } else if (a.type === 'Picklist' && a.options) {
      out[f + FORMATTED] = a.options[Number(v)] || String(v);
    } else if (a.type === 'Integer' || a.type === 'Decimal') {
      out[f + FORMATTED] = new Intl.NumberFormat('de-AT').format(Number(v));
    }
  });
  return out;
}

function lookupTarget(logical: string): ITableDef | undefined {
  for (const d of DATABASES) {
    if (d.tables[logical]) {
      return d.tables[logical];
    }
  }
  return undefined;
}

// ---------- $filter ----------
type Pred = (r: Record<string, unknown>) => boolean;
interface ITok { k: 'id' | 'str' | 'num' | 'guid' | 'date' | 'lp' | 'rp' | 'comma'; v: string }

function tokenize(s: string): ITok[] {
  const out: ITok[] = [];
  let i = 0;
  while (i < s.length) {
    const c = s.charAt(i);
    if (/\s/.test(c)) { i++; continue; }
    if (c === '(') { out.push({ k: 'lp', v: c }); i++; continue; }
    if (c === ')') { out.push({ k: 'rp', v: c }); i++; continue; }
    if (c === ',') { out.push({ k: 'comma', v: c }); i++; continue; }
    if (c === "'") {
      let j = i + 1; let v = '';
      while (j < s.length) {
        if (s.charAt(j) === "'" && s.charAt(j + 1) === "'") { v += "'"; j += 2; continue; }
        if (s.charAt(j) === "'") { break; }
        v += s.charAt(j); j++;
      }
      out.push({ k: 'str', v }); i = j + 1; continue;
    }
    const rest = s.substring(i);
    const g = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i.exec(rest);
    if (g) { out.push({ k: 'guid', v: g[0].toLowerCase() }); i += g[0].length; continue; }
    const d = /^\d{4}-\d{2}-\d{2}(T[0-9:.]+(Z|[+-]\d{2}:\d{2})?)?/.exec(rest);
    if (d) { out.push({ k: 'date', v: d[0] }); i += d[0].length; continue; }
    const n = /^-?\d+(\.\d+)?/.exec(rest);
    if (n) { out.push({ k: 'num', v: n[0] }); i += n[0].length; continue; }
    const id = /^[A-Za-z_][A-Za-z0-9_]*/.exec(rest);
    if (id) { out.push({ k: 'id', v: id[0] }); i += id[0].length; continue; }
    throw new ODataError(400, `Syntax error at position ${i} in '${s}'.`);
  }
  return out;
}

function parseFilter(t: ITableDef, s: string): Pred {
  const toks = tokenize(s);
  let p = 0;
  const peek = (): ITok | undefined => toks[p];
  const isWord = (w: string): boolean => !!peek() && peek()!.k === 'id' && peek()!.v.toLowerCase() === w;
  const expect = (k: ITok['k']): ITok => {
    const tok = toks[p++];
    if (!tok || tok.k !== k) { throw new ODataError(400, `Syntax error in $filter '${s}'.`); }
    return tok;
  };
  const literal = (): unknown => {
    const tok = toks[p++];
    if (!tok) { throw new ODataError(400, `Syntax error in $filter '${s}'.`); }
    if (tok.k === 'str' || tok.k === 'guid' || tok.k === 'date') { return tok.v; }
    if (tok.k === 'num') { return Number(tok.v); }
    if (tok.k === 'id') {
      const w = tok.v.toLowerCase();
      if (w === 'true') { return true; }
      if (w === 'false') { return false; }
      if (w === 'null') { return null; }
    }
    throw new ODataError(400, `Syntax error in $filter '${s}'.`);
  };
  const norm = (v: unknown): unknown => (typeof v === 'string' ? v.toLowerCase() : v === undefined ? null : v);
  const compare = (a: unknown, b: unknown): number => {
    if (a === null || b === null) { return NaN; }
    const da = typeof a === 'string' && /^\d{4}-\d{2}-\d{2}/.test(a) ? Date.parse(a) : a;
    const db = typeof b === 'string' && /^\d{4}-\d{2}-\d{2}/.test(b) ? Date.parse(b) : b;
    if (typeof da === 'number' && typeof db === 'number') { return da - db; }
    return String(norm(a)).localeCompare(String(norm(b)));
  };

  const primary = (): Pred => {
    if (peek() && peek()!.k === 'lp') {
      p++;
      const e = orExpr();
      expect('rp');
      return e;
    }
    const id = expect('id');
    if (peek() && peek()!.k === 'lp') {
      const fn = id.v.toLowerCase();
      p++;
      const field = expect('id').v;
      attrFor(t, field);
      expect('comma');
      const val = String(literal()).toLowerCase();
      expect('rp');
      return r => {
        const v = String(r[field] === null || r[field] === undefined ? '' : r[field]).toLowerCase();
        if (fn === 'contains') { return v.indexOf(val) >= 0; }
        if (fn === 'startswith') { return v.indexOf(val) === 0; }
        if (fn === 'endswith') { return v.length >= val.length && v.substring(v.length - val.length) === val; }
        throw new ODataError(400, `Unknown function '${fn}'.`);
      };
    }
    const field = id.v;
    attrFor(t, field);
    const op = expect('id').v.toLowerCase();
    const val = literal();
    return r => {
      const v = r[field] === undefined ? null : r[field];
      switch (op) {
        case 'eq': return norm(v) === norm(val);
        case 'ne': return norm(v) !== norm(val);
        case 'gt': return compare(v, val) > 0;
        case 'ge': return compare(v, val) >= 0;
        case 'lt': return compare(v, val) < 0;
        case 'le': return compare(v, val) <= 0;
        default: throw new ODataError(400, `Unknown operator '${op}'.`);
      }
    };
  };
  const unary = (): Pred => {
    if (isWord('not')) { p++; const e = unary(); return r => !e(r); }
    return primary();
  };
  const andExpr = (): Pred => {
    let e = unary();
    while (isWord('and')) { p++; const l = e; const rr = unary(); e = r => l(r) && rr(r); }
    return e;
  };
  const orExpr = (): Pred => {
    let e = andExpr();
    while (isWord('or')) { p++; const l = e; const rr = andExpr(); e = r => l(r) || rr(r); }
    return e;
  };
  const result = orExpr();
  if (p < toks.length) { throw new ODataError(400, `Syntax error in $filter '${s}'.`); }
  return result;
}

// ---------- Anfrage ----------
function handle(url: string, headers: { [k: string]: string }): IMockResponse {
  const u = new URL(url);
  const db = findDb(u.hostname);
  if (!db) {
    throw new TypeError('Failed to fetch');
  }
  const path = decodeURIComponent(u.pathname.replace(/^\/api\/data\/v9\.2\//, ''));
  const q = u.searchParams;
  const prefer = headers.Prefer || headers.prefer || '';
  const maxPage = Number((/odata\.maxpagesize=(\d+)/.exec(prefer) || [])[1] || 5000);

  if (path === 'EntityDefinitions') {
    return respond(200, {
      value: Object.keys(db.tables).map(k => db.tables[k]).map(t => ({
        LogicalName: t.logicalName, EntitySetName: t.entitySet,
        DisplayName: { UserLocalizedLabel: { Label: t.displayName } },
        IsCustomEntity: /^admin_/.test(t.logicalName), IsValidForAdvancedFind: true, IsIntersect: false, IsPrivate: false
      }))
    });
  }
  let m = /^EntityDefinitions\(LogicalName='([^']+)'\)(\/Attributes)?$/.exec(path);
  if (m) {
    const t = tableByLogical(db, m[1]);
    if (m[2]) {
      return respond(200, {
        value: t.attrs.map(a => ({
          LogicalName: a.name,
          AttributeType: a.type,
          DisplayName: { UserLocalizedLabel: { Label: a.label } },
          IsValidForRead: true,
          AttributeOf: null
        }))
      });
    }
    return respond(200, {
      LogicalName: t.logicalName,
      EntitySetName: t.entitySet,
      PrimaryIdAttribute: t.primaryId,
      PrimaryNameAttribute: t.primaryName,
      DisplayName: { UserLocalizedLabel: { Label: t.displayName } }
    });
  }

  m = /^([a-z_]+)\(([0-9a-f-]{36})\)$/i.exec(path);
  if (m) {
    const t = tableBySet(db, m[1]);
    const row = t.rows.filter(r => String(r[t.primaryId]).toLowerCase() === m![2].toLowerCase())[0];
    if (!row) {
      throw new ODataError(404, `${t.logicalName} With Id = ${m[2]} Does Not Exist`);
    }
    return respond(200, project(t, row, undefined));
  }

  m = /^([a-z_]+)$/i.exec(path);
  if (m) {
    const t = tableBySet(db, m[1]);
    const apply = q.get('$apply');
    if (apply) {
      // filter(<ausdruck>)/aggregate($count as total) bzw. aggregate(<pk> with countdistinct as total)
      const gm = /^(?:filter\((.*)\)\/)?groupby\(\(([^)]+)\),aggregate\((.*)\)\)$/.exec(apply);
      if (gm && gm[2].indexOf(',') >= 0) {
        // groupby über mehrere Spalten (z. B. Environment + Besitzer)
        const fields = gm[2].split(',').map(f => f.trim());
        fields.forEach(f => attrFor(t, f));
        let rows = t.rows;
        if (gm[1]) { const pred = parseFilter(t, gm[1]); rows = rows.filter(pred); }
        const counts: { [k: string]: { vals: unknown[]; n: number } } = {};
        rows.forEach(r => {
          const vals = fields.map(f => (r[f] === undefined ? null : r[f]));
          const k = JSON.stringify(vals);
          counts[k] = counts[k] || { vals, n: 0 };
          counts[k].n++;
        });
        return respond(200, { value: Object.keys(counts).map(k => {
          const row: Record<string, unknown> = { n: counts[k].n };
          fields.forEach((f, i) => { row[f] = counts[k].vals[i]; });
          return row;
        }) });
      }
      if (gm) {
        const groupField = gm[2].trim();
        attrFor(t, groupField);
        let rows = t.rows;
        if (gm[1]) { const pred = parseFilter(t, gm[1]); rows = rows.filter(pred); }
        if (rows.length > 50000) { throw new ODataError(400, 'AggregateQueryRecordLimit exceeded. Cannot perform this operation.'); }
        const maxM = /^([a-z0-9_]+) with max as ([a-z]+)$/i.exec(gm[3].trim());
        if (maxM) {
          attrFor(t, maxM[1]);
          const best: { [k: string]: string } = {};
          rows.forEach(r => {
            const k = String(r[groupField] === undefined ? null : r[groupField]);
            const v = r[maxM[1]] as string | null;
            if (v && (!best[k] || v > best[k])) { best[k] = v; }
          });
          return respond(200, { value: Object.keys(best).map(k => ({ [groupField]: k === 'null' ? null : k, [maxM[2]]: best[k] })) });
        }
        const counts: { [k: string]: number } = {};
        rows.forEach(r => { const k = String(r[groupField] === undefined ? null : r[groupField]); counts[k] = (counts[k] || 0) + 1; });
        const ga = attrFor(t, groupField);
        return respond(200, {
          value: Object.keys(counts).map(k => {
            const row: Record<string, unknown> = { [groupField]: k === 'null' ? null : (ga.type === 'Picklist' ? Number(k) : k), n: counts[k] };
            if (ga.type === 'Picklist' && ga.options && k !== 'null') { row[groupField + FORMATTED] = ga.options[Number(k)] || k; }
            return row;
          })
        });
      }
      const am = /^(?:filter\((.*)\)\/)?aggregate\((.*)\)$/.exec(apply);
      if (!am) { throw new ODataError(400, `Unsupported $apply '${apply}'.`); }
      const agg = /^(\$count|([a-z_]+) with countdistinct) as ([a-z_]+)$/i.exec(am[2].trim());
      if (!agg) { throw new ODataError(400, `Unsupported aggregate '${am[2]}'.`); }
      let rows = t.rows;
      if (am[1]) { const pred = parseFilter(t, am[1]); rows = rows.filter(pred); }
      if (rows.length > 50000) {
        throw new ODataError(400, 'AggregateQueryRecordLimit exceeded. Cannot perform this operation.');
      }
      return respond(200, { value: [{ [agg[3]]: rows.length }] });
    }
    const select = q.get('$select') ? q.get('$select')!.split(',').map(s => s.trim()).filter(Boolean) : undefined;
    if (select) { select.forEach(f => attrFor(t, f)); }
    let rows = t.rows.slice();
    const filter = q.get('$filter');
    if (filter) {
      const pred = parseFilter(t, filter);
      rows = rows.filter(pred);
    }
    const orderby = q.get('$orderby');
    if (orderby) {
      const parts = orderby.split(',').map(s => s.trim().split(/\s+/));
      parts.forEach(([f]) => attrFor(t, f));
      rows.sort((a, b) => {
        for (const [f, dir] of parts) {
          const va = a[f]; const vb = b[f];
          if (va === vb) { continue; }
          if (va === null || va === undefined) { return dir === 'desc' ? 1 : -1; }
          if (vb === null || vb === undefined) { return dir === 'desc' ? -1 : 1; }
          const c = typeof va === 'number' && typeof vb === 'number' ? va - vb : String(va).localeCompare(String(vb), 'de');
          if (c !== 0) { return dir === 'desc' ? -c : c; }
        }
        return 0;
      });
    }
    const skip = Number(q.get('$skiptoken') || 0);
    const pageRows = rows.slice(skip, skip + maxPage);
    const body: Record<string, unknown> = { value: pageRows.map(r => project(t, r, select)) };
    if (q.get('$count') === 'true') {
      body['@odata.count'] = Math.min(rows.length, 5000);
    }
    if (skip + maxPage < rows.length) {
      const next = new URL(url);
      next.searchParams.set('$skiptoken', String(skip + maxPage));
      body['@odata.nextLink'] = next.toString();
    }
    return respond(200, body);
  }

  throw new ODataError(404, `Resource not found for the segment '${path}'.`);
}

/** Antwort mit kurzer Verzögerung wie ein echter Server. */
export function mockGet(url: string, headers: { [k: string]: string }): Promise<IMockResponse> {
  const delay = 220 + Math.random() * 420;
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      try {
        resolve(handle(url, headers));
      } catch (e) {
        if (e instanceof ODataError) {
          resolve(respond(e.status, { error: { code: '0x80060888', message: e.message } }));
        } else {
          reject(e);
        }
      }
    }, delay);
  });
}
