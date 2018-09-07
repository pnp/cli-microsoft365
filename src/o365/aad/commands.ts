const prefix: string = 'aad';

export default {
  CONNECT: `${prefix} connect`,
  DISCONNECT: `${prefix} disconnect`,
  LOGIN: `${prefix} login`,
  LOGOUT: `${prefix} logout`,
  OAUTH2GRANT_ADD: `${prefix} oauth2grant add`,
  OAUTH2GRANT_LIST: `${prefix} oauth2grant list`,
  OAUTH2GRANT_REMOVE: `${prefix} oauth2grant remove`,
  OAUTH2GRANT_SET: `${prefix} oauth2grant set`,
  SP_GET: `${prefix} sp get`,
  STATUS: `${prefix} status`
};