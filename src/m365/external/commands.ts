const prefix: string = 'external';
const searchPrefix: string = 'search';

export default {
  CONNECTION_ADD: `${prefix} connection add`,
  CONNECTION_GET: `${prefix} connection get`,
  CONNECTION_LIST: `${prefix} connection list`,
  CONNECTION_REMOVE: `${prefix} connection remove`,
  CONNECTION_SCHEMA_ADD: `${prefix} connection schema add`,
  EXTERNALCONNECTION_ADD: `${searchPrefix} externalconnection add`,
  EXTERNALCONNECTION_GET: `${searchPrefix} externalconnection get`,
  EXTERNALCONNECTION_LIST: `${searchPrefix} externalconnection list`,
  EXTERNALCONNECTION_REMOVE: `${searchPrefix} externalconnection remove`,
  EXTERNALCONNECTION_SCHEMA_ADD: `${searchPrefix} externalconnection schema add`,
  ITEM_ADD: `${prefix} item add`
};