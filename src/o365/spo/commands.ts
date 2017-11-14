const prefix: string = 'spo';

export default {
  CONNECT: `${prefix} connect`,
  DISCONNECT: `${prefix} disconnect`,
  STATUS: `${prefix} status`,
  STORAGEENTITY_LIST: `${prefix} storageentity list`,
  STORAGEENTITY_GET: `${prefix} storageentity get`,
  STORAGEENTITY_SET: `${prefix} storageentity set`,
  STORAGEENTITY_REMOVE: `${prefix} storageentity remove`,
  TENANT_CDN_GET: `${prefix} tenant cdn get`,
  TENANT_CDN_SET: `${prefix} tenant cdn set`,
  TENANT_CDN_ORIGIN_LIST: `${prefix} tenant cdn origin list`,
  TENANT_CDN_ORIGIN_SET: `${prefix} tenant cdn origin set`,
  TENANT_CDN_ORIGIN_REMOVE: `${prefix} tenant cdn origin remove`,
  TENANT_CDN_POLICY_LIST: `${prefix} tenant cdn policy list`,
  TENANT_CDN_POLICY_SET: `${prefix} tenant cdn policy set`
};