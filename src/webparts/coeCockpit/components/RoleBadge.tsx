import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRoleStatus } from '../services/RoleCheckService';
import { PIM_URL } from '../services/PowerPlatformApiService';
import { formatTime } from '../utils/format';

/**
 * Anzeige rechts oben: Ist die Rolle „Power Platform Administrator“ (bzw. „Globaler Administrator“) gerade aktiv?
 * Grün = aktiv, Gelb = nicht aktiv (PIM aktivieren), Grau = nicht prüfbar. Klick öffnet Details und „Erneut prüfen“.
 * Die Prüfung selbst macht das Cockpit (dasselbe Ergebnis blendet die Aktions-Schaltflächen ein bzw. aus).
 */
export interface IRoleBadgeProps {
  status: IRoleStatus | undefined;
  checking: boolean;
  /** Frisches Token holen und neu prüfen */
  onRecheck: () => void;
  /** Zeigt im Detail an, dass ohne aktive Rolle die Aktions-Schaltflächen ausgeblendet sind */
  actionsHint?: boolean;
  themeVars: React.CSSProperties;
}

export const RoleBadge: React.FC<IRoleBadgeProps> = props => {
  const { status, checking } = props;
  const [open, setOpen] = React.useState<boolean>(false);
  const ref = React.useRef<HTMLButtonElement>(null);

  const state = status ? status.state : 'unknown';
  const tone = checking && !status ? styles.roleBadgeChecking : state === 'active' ? styles.roleBadgeActive : state === 'inactive' ? styles.roleBadgeInactive : styles.roleBadgeUnknown;
  const label = checking && !status ? 'Admin-Rolle wird geprüft …'
    : state === 'active' ? `${status && status.roles.indexOf('Power Platform Administrator') >= 0 ? 'Power Platform Admin' : 'Globaler Admin'} aktiv`
      : state === 'inactive' ? 'PIM-Rolle nicht aktiv' : 'Admin-Rolle unbekannt';

  return (
    <>
      <button
        ref={ref}
        type="button"
        className={`${styles.roleBadge} ${tone}`}
        onClick={() => setOpen(!open)}
        aria-expanded={open}
        title="Status der Rolle „Power Platform Administrator“ (PIM)"
      >
        <Icon iconName={state === 'active' ? 'ShieldSolid' : state === 'inactive' ? 'ShieldAlert' : 'Shield'} />
        <span>{label}</span>
      </button>
      {open && ref.current && (
        <Callout target={ref.current} onDismiss={() => setOpen(false)} directionalHint={DirectionalHint.bottomRightEdge} setInitialFocus>
          <div className={styles.roleCallout} style={props.themeVars}>
            <div className={styles.roleCalloutTitle}>
              {state === 'active' ? 'Admin-Rolle aktiv' : state === 'inactive' ? 'Admin-Rolle nicht aktiv' : 'Admin-Rolle nicht prüfbar'}
            </div>
            {state === 'active' && (
              <p>Aktiv: <b>{status!.roles.join(', ')}</b>. Das Cockpit kann alle Bereiche lesen (Default Environment, ALM-Pipelines, Copilot-Credits …).</p>
            )}
            {state === 'inactive' && (
              <p>
                Für die volle Nutzung braucht das Cockpit die aktive Rolle <b>„Power Platform Administrator“</b> (oder „Globaler Administrator“).
                Aktiviere sie in Privileged Identity Management und klicke danach auf „Erneut prüfen“.
                {status && status.roles.length > 0 && <> Derzeit aktiv: {status.roles.join(', ')}.</>}
              </p>
            )}
            {state !== 'active' && props.actionsHint && (
              <p>Solange die Rolle nicht aktiv ist, sind die Aktions-Schaltflächen (Quarantäne, Löschen, Besitzer, Freigaben, Environments …) ausgeblendet.</p>
            )}
            {state === 'unknown' && (
              <p>
                Die aktiven Rollen konnten aus der Anmeldung nicht gelesen werden{status && status.problem ? ` (${status.problem})` : ''}.
                Ob die Rolle aktiv ist, zeigt sich dann an den einzelnen Bereichen; Aktions-Schaltflächen bleiben vorsichtshalber ausgeblendet.
              </p>
            )}
            <p className={styles.roleCalloutMeta}>
              {status ? `Geprüft ${formatTime(status.checkedAt)} – laut Anmelde-Token.` : 'Wird geprüft …'} Nach dem Aktivieren dauert es oft 1–2 Minuten, bis die Rolle im Token steht.
            </p>
            <div className={styles.roleCalloutButtons}>
              {state !== 'active' && (
                <PrimaryButton text="In PIM aktivieren" iconProps={{ iconName: 'OpenInNewWindow' }} href={PIM_URL} target="_blank" rel="noopener noreferrer" />
              )}
              <DefaultButton text={checking ? 'Wird geprüft …' : 'Erneut prüfen'} iconProps={{ iconName: 'Refresh' }} disabled={checking} onClick={props.onRecheck} />
            </div>
          </div>
        </Callout>
      )}
    </>
  );
};
