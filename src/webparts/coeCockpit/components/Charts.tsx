import * as React from 'react';
import styles from './CoeCockpit.module.scss';
import { formatNumber } from '../utils/format';

/*
 * Schlanke SVG-Diagramme ohne Fremdbibliothek.
 * Regeln: dünne Markierungen (2px Linien, Balken ≤ 24px mit 4px gerundetem Ende),
 * Text immer in Textfarben (nie in Serienfarbe), Legende ab zwei Serien,
 * Hover-Tooltip auf jeder Markierung, Farben über CSS-Variablen (Hell/Dunkel getrennt gewählt).
 */

/** Gemessene Breite eines Elements (für pixelgenaue SVGs ohne verzerrte Schrift). */
function useWidth(fallback: number): [React.RefObject<HTMLDivElement>, number] {
  const ref = React.useRef<HTMLDivElement>(null);
  const [width, setWidth] = React.useState<number>(fallback);
  React.useEffect(() => {
    const el = ref.current;
    if (!el) {
      return;
    }
    const update = (): void => setWidth(Math.max(120, Math.round(el.getBoundingClientRect().width)));
    update();
    const RO = (window as unknown as { ResizeObserver?: new (cb: () => void) => { observe: (e: Element) => void; disconnect: () => void } }).ResizeObserver;
    if (RO) {
      const ro = new RO(update);
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);
  return [ref, width];
}

/** „Schöne“ Achsenwerte: 0 … max in 3–5 runden Schritten. */
function niceTicks(max: number): number[] {
  if (max <= 0) {
    return [0, 1];
  }
  const raw = max / 4;
  const pow = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map(m => m * pow).filter(s => s >= raw)[0] || pow * 10;
  const ticks: number[] = [];
  for (let v = 0; v <= max + step * 0.001; v += step) {
    ticks.push(Math.round(v));
  }
  if (ticks[ticks.length - 1] < max) {
    ticks.push(Math.round(ticks[ticks.length - 1] + step));
  }
  return ticks;
}

export function compactNumber(n: number): string {
  if (n >= 1000000) {
    return `${formatNumber(Math.round(n / 100000) / 10)} Mio.`;
  }
  if (n >= 10000) {
    return `${formatNumber(Math.round(n / 100) / 10)}K`;
  }
  return formatNumber(n);
}

// ---------------------------------------------------------------------------
// Sparkline (12 Punkte, Verlauf in Zurückhaltungsfarbe, aktueller Wert in Akzent)
// ---------------------------------------------------------------------------
export const Sparkline: React.FC<{ values: number[]; label: string }> = ({ values, label }) => {
  const [ref, measured] = useWidth(120);
  const w = measured;
  const h = 30;
  if (values.length < 2) {
    return <div ref={ref} />;
  }
  const max = Math.max(1, ...values);
  const x = (i: number): number => 2 + (i * (w - 8)) / (values.length - 1);
  const y = (v: number): number => h - 4 - (v / max) * (h - 8);
  const d = values.map((v, i) => `${i === 0 ? 'M' : 'L'}${x(i).toFixed(1)},${y(v).toFixed(1)}`).join(' ');
  const last = values.length - 1;
  return (
    <div ref={ref} className={styles.sparkBox}>
      <svg className={styles.sparkline} width={w} height={h} role="img" aria-label={`${label}: Verlauf der letzten ${values.length} Monate`}>
        <path d={`${d} L${x(last)},${h - 2} L${x(0)},${h - 2} Z`} className={styles.sparkArea} />
        <path d={d} className={styles.sparkLine} />
        <circle cx={x(last)} cy={y(values[last])} r={3.5} className={styles.sparkDot} />
      </svg>
    </div>
  );
};

// ---------------------------------------------------------------------------
// Verlauf: Linien je Serie, Fadenkreuz + Tooltip beim Überfahren
// ---------------------------------------------------------------------------
export interface ITrendSeries {
  key: string;
  label: string;
  /** CSS-Farbe, z. B. var(--coe-series-1) */
  color: string;
  values: number[];
}

export const TrendChart: React.FC<{
  series: ITrendSeries[]; labels: string[]; longLabels: string[]; height?: number;
  /** Letzter Punkt ist ein laufender (unvollständiger) Zeitraum – wird gestrichelt gezeichnet. */
  partialLast?: boolean;
}> = ({ series, labels, longLabels, height = 240, partialLast }) => {
  const [ref, width] = useWidth(640);
  const [hover, setHover] = React.useState<number | undefined>(undefined);
  const pad = { top: 16, right: 56, bottom: 28, left: 44 };
  const innerW = Math.max(10, width - pad.left - pad.right);
  const innerH = height - pad.top - pad.bottom;
  const n = labels.length;
  const max = Math.max(1, ...series.map(s => Math.max(...s.values)));
  const ticks = niceTicks(max);
  const top = ticks[ticks.length - 1];
  const x = (i: number): number => pad.left + (n <= 1 ? 0 : (i * innerW) / (n - 1));
  const y = (v: number): number => pad.top + innerH - (v / top) * innerH;
  const path = (vals: number[], from: number = 0): string => vals.map((v, i) => `${i === from ? 'M' : 'L'}${x(i).toFixed(1)},${y(v).toFixed(1)}`).filter((_p, i) => i >= from).join(' ');
  const solidEnd = partialLast && n > 2 ? n - 2 : n - 1;

  const onMove = (e: React.MouseEvent<SVGRectElement>): void => {
    const rect = (e.currentTarget as SVGRectElement).getBoundingClientRect();
    const px = e.clientX - rect.left;
    const i = Math.round((px / rect.width) * (n - 1));
    setHover(Math.max(0, Math.min(n - 1, i)));
  };
  const tipLeft = hover !== undefined ? Math.min(Math.max(x(hover) + 12, 8), width - 180) : 0;

  return (
    <div className={styles.chartBox} ref={ref}>
      <svg width={width} height={height} role="img" aria-label={`Verlauf: ${series.map(s => s.label).join(', ')}`}>
        {ticks.map(t => (
          <g key={t}>
            <line x1={pad.left} x2={pad.left + innerW} y1={y(t)} y2={y(t)} className={t === 0 ? styles.chartBaseline : styles.chartGrid} />
            <text x={pad.left - 8} y={y(t) + 4} textAnchor="end" className={styles.chartTick}>{compactNumber(t)}</text>
          </g>
        ))}
        {labels.map((l, i) => {
          // Beschriftung ausdünnen, damit nichts überlappt (ca. 56px pro Label); der letzte Monat bleibt immer sichtbar
          const every = Math.max(1, Math.ceil((n * 56) / innerW));
          const show = i === n - 1 || (i % every === 0 && n - 1 - i >= every);
          return show && <text key={i} x={x(i)} y={height - 8} textAnchor="middle" className={styles.chartTick}>{l}</text>;
        })}
        {series.map(s => (
          <g key={s.key}>
            <path d={`${path(s.values.slice(0, solidEnd + 1))} L${x(solidEnd)},${y(0)} L${x(0)},${y(0)} Z`} fill={s.color} className={styles.chartArea} />
            <path d={path(s.values.slice(0, solidEnd + 1))} stroke={s.color} className={styles.chartLine} />
            {solidEnd < n - 1 && <path d={path(s.values, solidEnd)} stroke={s.color} className={`${styles.chartLine} ${styles.chartLinePartial}`} />}
            <circle cx={x(n - 1)} cy={y(s.values[n - 1])} r={4.5} fill={s.color} className={styles.chartDot} />
            <text x={x(n - 1) + 10} y={y(s.values[n - 1]) + 4} className={styles.chartEndLabel}>{compactNumber(s.values[n - 1])}</text>
          </g>
        ))}
        {hover !== undefined && (
          <g pointerEvents="none">
            <line x1={x(hover)} x2={x(hover)} y1={pad.top} y2={pad.top + innerH} className={styles.chartCrosshair} />
            {series.map(s => <circle key={s.key} cx={x(hover)} cy={y(s.values[hover])} r={4.5} fill={s.color} className={styles.chartDot} />)}
          </g>
        )}
        <rect
          x={pad.left - 10} y={pad.top} width={innerW + 20} height={innerH} fill="transparent"
          onMouseMove={onMove} onMouseLeave={() => setHover(undefined)}
        />
      </svg>
      {hover !== undefined && (
        <div className={styles.chartTip} style={{ left: tipLeft, top: 8 }} role="status">
          <div className={styles.chartTipTitle}>{longLabels[hover]}</div>
          {series.map(s => (
            <div key={s.key} className={styles.chartTipRow}>
              <span className={styles.legendKey} style={{ background: s.color }} />
              <span>{s.label}</span>
              <strong>{formatNumber(s.values[hover])}</strong>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

// ---------------------------------------------------------------------------
// Segmentierter Balken (Teil vom Ganzen, ≤ 6 Teile) mit Legende
// ---------------------------------------------------------------------------
export interface IPart {
  key: string;
  label: string;
  value: number;
  color: string;
  /** Dunkle statt weißer Beschriftung im Block (helle Füllungen) */
  darkText?: boolean;
  onClick?: () => void;
}

export const SegmentedBar: React.FC<{ parts: IPart[]; label: string }> = ({ parts, label }) => {
  const [ref, width] = useWidth(320);
  const [hover, setHover] = React.useState<string | undefined>(undefined);
  const total = parts.reduce((s, p) => s + p.value, 0);
  const visible = parts.filter(p => p.value > 0);
  const gap = 2;
  const avail = Math.max(0, width - gap * Math.max(0, visible.length - 1));
  return (
    <div className={styles.segWrap} ref={ref}>
      <div className={styles.segBar} role="img" aria-label={`${label}: ${parts.map(p => `${p.label} ${formatNumber(p.value)}`).join(', ')}`}>
        {visible.map(p => {
          const w = total ? (p.value / total) * avail : 0;
          const share = total ? Math.round((p.value / total) * 100) : 0;
          return (
            <button
              key={p.key}
              type="button"
              className={`${styles.segPart} ${hover && hover !== p.key ? styles.segDim : ''}`}
              style={{ width: w, padding: w >= 46 ? '0 8px' : 0, background: p.color, color: p.darkText ? '#0b0b0b' : '#ffffff' }}
              onMouseEnter={() => setHover(p.key)}
              onMouseLeave={() => setHover(undefined)}
              onFocus={() => setHover(p.key)}
              onBlur={() => setHover(undefined)}
              onClick={p.onClick}
              title={`${p.label}: ${formatNumber(p.value)} (${share} %)`}
              aria-label={`${p.label}: ${formatNumber(p.value)} (${share} %)`}
            >
              {w >= 46 ? `${share} %` : ''}
            </button>
          );
        })}
      </div>
      <ul className={styles.segLegend}>
        {parts.map(p => (
          <li key={p.key}>
            <button
              type="button"
              className={`${styles.segLegendItem} ${hover === p.key ? styles.segLegendHover : ''}`}
              onMouseEnter={() => setHover(p.key)}
              onMouseLeave={() => setHover(undefined)}
              onClick={p.onClick}
              disabled={!p.onClick}
            >
              <span className={styles.legendKey} style={{ background: p.color }} />
              <span className={styles.segLegendLabel}>{p.label}</span>
              <span className={styles.segLegendValue}>{formatNumber(p.value)}</span>
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
};

// ---------------------------------------------------------------------------
// Rangliste: waagrechte Balken, Wert am Balkenende, Zeile anklickbar
// ---------------------------------------------------------------------------
export interface IBarItem {
  key: string;
  label: string;
  sub?: string;
  value: number;
  /** Zusatz im Tooltip, z. B. „183 Flows“ */
  detail?: string;
  onClick?: () => void;
}

export const BarList: React.FC<{ items: IBarItem[]; unit: string }> = ({ items, unit }) => {
  const max = Math.max(1, ...items.map(i => i.value));
  return (
    <ol className={styles.barList}>
      {items.map((it, idx) => (
        <li key={it.key}>
          <button
            type="button"
            className={styles.barRow}
            onClick={it.onClick}
            disabled={!it.onClick}
            title={`${it.label}: ${formatNumber(it.value)} ${unit}${it.detail ? ` · ${it.detail}` : ''}`}
          >
            <span className={styles.barRank}>{idx + 1}</span>
            <span className={styles.barText}>
              <span className={styles.barLabel}>{it.label}</span>
              {it.sub && <span className={styles.barSub}>{it.sub}</span>}
            </span>
            <span className={styles.barTrack}>
              <span className={styles.barFill} style={{ width: `${Math.max(2, (it.value / max) * 100)}%` }} />
            </span>
            <span className={styles.barValue}>{formatNumber(it.value)}</span>
          </button>
        </li>
      ))}
    </ol>
  );
};
