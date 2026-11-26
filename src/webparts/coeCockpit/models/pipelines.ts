import { ITableConfig } from './ITableConfig';

/**
 * ALM-Pipelines (Power Platform Pipelines). Die Daten liegen nicht im CoE, sondern im Dataverse des
 * Pipelines-Host-Environments – gelesen mit den Rechten des angemeldeten Benutzers.
 * Spalten, die es in der eigenen Version nicht gibt, werden beim Laden automatisch übersprungen.
 */
export const GROUP_ALM = 'ALM-Pipelines';
export const PIPELINE_KEY_PREFIX = 'alm:';

export function buildPipelineTables(hostUrl: string): ITableConfig[] {
  return [
    {
      key: `${PIPELINE_KEY_PREFIX}runs`,
      title: 'Bereitstellungen',
      logicalName: 'deploymentstagerun',
      icon: 'Rocket',
      group: GROUP_ALM,
      environmentUrl: hostUrl,
      showInKpis: false,
      defaultSort: { field: 'starttime', desc: true },
      columns: [
        { field: 'artifactname', label: 'Solution', type: 'title' },
        { field: 'artifactversion', label: 'Version', type: 'text' },
        { field: '_deploymentpipelineid_value', label: 'Pipeline', type: 'choice' },
        { field: '_deploymentstageid_value', label: 'Stufe', type: 'choice' },
        {
          field: 'stagerunstatus', label: 'Status', type: 'status',
          statusMap: { 'Succeeded': 'success', 'Erfolgreich': 'success', 'Failed': 'danger', 'Fehlgeschlagen': 'danger', 'In Progress': 'info', 'Wird ausgeführt': 'info', 'Canceled': 'neutral', 'Abgebrochen': 'neutral', 'Pending Approval': 'warning' }
        },
        { field: '_createdby_value', label: 'Ausgelöst von', type: 'person' },
        { field: 'starttime', label: 'Start', type: 'datetime' },
        { field: 'endtime', label: 'Ende', type: 'datetime', hidden: true }
      ],
      links: []
    },
    {
      key: `${PIPELINE_KEY_PREFIX}pipelines`,
      title: 'Pipelines',
      logicalName: 'deploymentpipeline',
      icon: 'Flow',
      group: GROUP_ALM,
      environmentUrl: hostUrl,
      showInKpis: false,
      defaultSort: { field: 'name', desc: false },
      columns: [
        { field: 'name', label: 'Pipeline', type: 'title' },
        { field: 'description', label: 'Beschreibung', type: 'text' },
        { field: '_ownerid_value', label: 'Besitzer', type: 'person' },
        { field: 'modifiedon', label: 'Geändert', type: 'date' }
      ],
      links: [],
      related: [
        {
          key: 'stages', title: 'Stufen', logicalName: 'deploymentstage', lookupField: '_deploymentpipelineid_value',
          columns: [
            { field: 'name', label: 'Stufe', type: 'title' },
            { field: '_targetdeploymentenvironmentid_value', label: 'Ziel-Environment', type: 'text' },
            { field: '_previousdeploymentstageid_value', label: 'Vorherige Stufe', type: 'text' }
          ]
        },
        {
          key: 'runs', title: 'Bereitstellungen', logicalName: 'deploymentstagerun', lookupField: '_deploymentpipelineid_value',
          defaultSort: { field: 'starttime', desc: true },
          columns: [
            { field: 'artifactname', label: 'Solution', type: 'title' },
            { field: 'artifactversion', label: 'Version', type: 'text' },
            { field: 'stagerunstatus', label: 'Status', type: 'choice' },
            { field: 'starttime', label: 'Start', type: 'datetime' }
          ]
        }
      ]
    },
    {
      key: `${PIPELINE_KEY_PREFIX}environments`,
      title: 'Pipeline-Environments',
      logicalName: 'deploymentenvironment',
      icon: 'Globe',
      group: GROUP_ALM,
      environmentUrl: hostUrl,
      showInKpis: false,
      defaultSort: { field: 'name', desc: false },
      columns: [
        { field: 'name', label: 'Environment', type: 'title' },
        { field: 'environmenttype', label: 'Rolle', type: 'choice' },
        { field: 'validationstatus', label: 'Prüfstatus', type: 'choice' },
        { field: 'environmentid', label: 'Environment-ID', type: 'text', hidden: true }
      ],
      links: []
    }
  ];
}
