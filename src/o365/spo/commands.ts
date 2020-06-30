const prefix: string = 'spo';

export default {
  APP_ADD: `${prefix} app add`,
  APP_DEPLOY: `${prefix} app deploy`,
  APP_GET: `${prefix} app get`,
  APP_INSTALL: `${prefix} app install`,
  APP_LIST: `${prefix} app list`,
  APP_REMOVE: `${prefix} app remove`,
  APP_RETRACT: `${prefix} app retract`,
  APP_UNINSTALL: `${prefix} app uninstall`,
  APP_UPGRADE: `${prefix} app upgrade`,
  APPPAGE_ADD:`${prefix} apppage add`,
  APPPAGE_SET:`${prefix} apppage set`,
  CDN_GET: `${prefix} cdn get`,
  CDN_ORIGIN_ADD: `${prefix} cdn origin add`,
  CDN_ORIGIN_LIST: `${prefix} cdn origin list`,
  CDN_ORIGIN_REMOVE: `${prefix} cdn origin remove`,
  CDN_POLICY_LIST: `${prefix} cdn policy list`,
  CDN_POLICY_SET: `${prefix} cdn policy set`,
  CDN_SET: `${prefix} cdn set`,
  CONTENTTYPE_ADD: `${prefix} contenttype add`,
  CONTENTTYPE_GET: `${prefix} contenttype get`,
  CONTENTTYPE_FIELD_REMOVE: `${prefix} contenttype field remove`,
  CONTENTTYPE_FIELD_SET: `${prefix} contenttype field set`,
  CONTENTTYPE_REMOVE: `${prefix} contenttype remove`,
  CONTENTTYPEHUB_GET: `${prefix} contenttypehub get`,
  CUSTOMACTION_ADD: `${prefix} customaction add`,
  CUSTOMACTION_CLEAR: `${prefix} customaction clear`,
  CUSTOMACTION_GET: `${prefix} customaction get`,
  CUSTOMACTION_SET: `${prefix} customaction set`,
  CUSTOMACTION_LIST: `${prefix} customaction list`,
  CUSTOMACTION_REMOVE: `${prefix} customaction remove`,
  EXTERNALUSER_LIST: `${prefix} externaluser list`,
  FEATURE_DISABLE: `${prefix} feature disable`,
  FEATURE_ENABLE: `${prefix} feature enable`,
  FEATURE_LIST: `${prefix} feature list`,
  FIELD_ADD: `${prefix} field add`,
  FIELD_GET: `${prefix} field get`,
  FIELD_REMOVE: `${prefix} field remove`,
  FIELD_SET: `${prefix} field set`,
  FILE_ADD: `${prefix} file add`,
  FILE_CHECKIN: `${prefix} file checkin`,
  FILE_CHECKOUT: `${prefix} file checkout`,
  FILE_COPY: `${prefix} file copy`,
  FILE_GET: `${prefix} file get`,
  FILE_LIST: `${prefix} file list`,
  FILE_MOVE: `${prefix} file move`,
  FILE_REMOVE: `${prefix} file remove`,
  FOLDER_ADD: `${prefix} folder add`,
  FOLDER_COPY: `${prefix} folder copy`,
  FOLDER_GET: `${prefix} folder get`,
  FOLDER_LIST: `${prefix} folder list`,
  FOLDER_MOVE: `${prefix} folder move`,
  FOLDER_REMOVE: `${prefix} folder remove`,
  FOLDER_RENAME: `${prefix} folder rename`,
  GET: `${prefix} get`,
  HIDEDEFAULTTHEMES_GET: `${prefix} hidedefaultthemes get`,
  HIDEDEFAULTTHEMES_SET: `${prefix} hidedefaultthemes set`,
  HOMESITE_GET: `${prefix} homesite get`,
  HOMESITE_REMOVE: `${prefix} homesite remove`,
  HOMESITE_SET: `${prefix} homesite set`,
  HUBSITE_CONNECT: `${prefix} hubsite connect`,
  HUBSITE_DATA_GET: `${prefix} hubsite data get`,
  HUBSITE_DISCONNECT: `${prefix} hubsite disconnect`,
  HUBSITE_GET: `${prefix} hubsite get`,
  HUBSITE_LIST: `${prefix} hubsite list`,
  HUBSITE_REGISTER: `${prefix} hubsite register`,
  HUBSITE_RIGHTS_GRANT: `${prefix} hubsite rights grant`,
  HUBSITE_RIGHTS_REVOKE: `${prefix} hubsite rights revoke`,
  HUBSITE_SET: `${prefix} hubsite set`,
  HUBSITE_THEME_SYNC: `${prefix} hubsite theme sync`,
  HUBSITE_UNREGISTER: `${prefix} hubsite unregister`,
  LIST_ADD: `${prefix} list add`,
  LIST_CONTENTTYPE_ADD: `${prefix} list contenttype add`,
  LIST_CONTENTTYPE_LIST: `${prefix} list contenttype list`,
  LIST_CONTENTTYPE_REMOVE: `${prefix} list contenttype remove`,
  LIST_GET: `${prefix} list get`,
  LIST_LABEL_GET: `${prefix} list label get`,
  LIST_LABEL_SET: `${prefix} list label set`,
  LIST_LIST: `${prefix} list list`,
  LIST_REMOVE: `${prefix} list remove`,
  LIST_SET: `${prefix} list set`,
  LIST_SITESCRIPT_GET: `${prefix} list sitescript get`,
  LIST_VIEW_GET: `${prefix} list view get`,
  LIST_VIEW_LIST: `${prefix} list view list`,
  LIST_VIEW_REMOVE: `${prefix} list view remove`,
  LIST_VIEW_SET: `${prefix} list view set`,
  LIST_VIEW_FIELD_ADD: `${prefix} list view field add`,
  LIST_VIEW_FIELD_REMOVE: `${prefix} list view field remove`,
  LIST_WEBHOOK_ADD: `${prefix} list webhook add`,
  LIST_WEBHOOK_GET: `${prefix} list webhook get`,
  LIST_WEBHOOK_LIST: `${prefix} list webhook list`,
  LIST_WEBHOOK_REMOVE: `${prefix} list webhook remove`,
  LIST_WEBHOOK_SET: `${prefix} list webhook set`,
  LISTITEM_ADD: `${prefix} listitem add`,
  LISTITEM_GET: `${prefix} listitem get`,
  LISTITEM_ISRECORD: `${prefix} listitem isrecord`,
  LISTITEM_LIST: `${prefix} listitem list`,
  LISTITEM_RECORD_DECLARE: `${prefix} listitem record declare`,
  LISTITEM_RECORD_UNDECLARE: `${prefix} listitem record undeclare`,
  LISTITEM_REMOVE: `${prefix} listitem remove`,
  LISTITEM_SET: `${prefix} listitem set`,
  MAIL_SEND: `${prefix} mail send`,
  NAVIGATION_NODE_ADD: `${prefix} navigation node add`,
  NAVIGATION_NODE_LIST: `${prefix} navigation node list`,
  NAVIGATION_NODE_REMOVE: `${prefix} navigation node remove`,
  ORGASSETSLIBRARY_ADD: `${prefix} orgassetslibrary add`,
  ORGASSETSLIBRARY_LIST: `${prefix} orgassetslibrary list`,
  ORGASSETSLIBRARY_REMOVE: `${prefix} orgassetslibrary remove`,
  ORGNEWSSITE_LIST: `${prefix} orgnewssite list`,
  ORGNEWSSITE_REMOVE: `${prefix} orgnewssite remove`,
  ORGNEWSSITE_SET: `${prefix} orgnewssite set`,
  PAGE_ADD: `${prefix} page add`,
  PAGE_GET: `${prefix} page get`,
  PAGE_LIST: `${prefix} page list`,
  PAGE_REMOVE: `${prefix} page remove`,
  PAGE_SET: `${prefix} page set`,
  PAGE_CLIENTSIDEWEBPART_ADD: `${prefix} page clientsidewebpart add`,
  PAGE_COLUMN_GET: `${prefix} page column get`,
  PAGE_COLUMN_LIST: `${prefix} page column list`,
  PAGE_CONTROL_GET: `${prefix} page control get`,
  PAGE_CONTROL_LIST: `${prefix} page control list`,
  PAGE_HEADER_SET: `${prefix} page header set`,
  PAGE_SECTION_ADD: `${prefix} page section add`,
  PAGE_SECTION_GET: `${prefix} page section get`,
  PAGE_SECTION_LIST: `${prefix} page section list`,
  PAGE_TEXT_ADD: `${prefix} page text add`,
  PROPERTYBAG_GET: `${prefix} propertybag get`,
  PROPERTYBAG_LIST: `${prefix} propertybag list`,
  PROPERTYBAG_REMOVE: `${prefix} propertybag remove`,
  PROPERTYBAG_SET: `${prefix} propertybag set`,
  REPORT_ACTIVITYFILECOUNTS: `${prefix} report activityfilecounts`,
  REPORT_ACTIVITYPAGES: `${prefix} report activitypages`,
  REPORT_ACTIVITYUSERCOUNTS: `${prefix} report activityusercounts`,
  REPORT_ACTIVITYUSERDETAIL: `${prefix} report activityuserdetail`,
  REPORT_SITEUSAGEDETAIL: `${prefix} report siteusagedetail`,
  REPORT_SITEUSAGEFILECOUNTS: `${prefix} report siteusagefilecounts`,
  REPORT_SITEUSAGEPAGES: `${prefix} report siteusagepages`,
  REPORT_SITEUSAGESITECOUNTS: `${prefix} report siteusagesitecounts`,
  REPORT_SITEUSAGESTORAGE: `${prefix} report siteusagestorage`,
  SEARCH: `${prefix} search`,
  SERVICEPRINCIPAL_GRANT_ADD: `${prefix} serviceprincipal grant add`,
  SERVICEPRINCIPAL_GRANT_LIST: `${prefix} serviceprincipal grant list`,
  SERVICEPRINCIPAL_GRANT_REVOKE: `${prefix} serviceprincipal grant revoke`,
  SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE: `${prefix} serviceprincipal permissionrequest approve`,
  SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY: `${prefix} serviceprincipal permissionrequest deny`,
  SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST: `${prefix} serviceprincipal permissionrequest list`,
  SERVICEPRINCIPAL_SET: `${prefix} serviceprincipal set`,
  SET: `${prefix} set`,
  SITE_ADD: `${prefix} site add`,
  SITE_APPCATALOG_ADD: `${prefix} site appcatalog add`,
  SITE_APPCATALOG_REMOVE: `${prefix} site appcatalog remove`,
  SITE_CLASSIC_ADD: `${prefix} site classic add`,
  SITE_CLASSIC_LIST: `${prefix} site classic list`,
  SITE_CLASSIC_REMOVE: `${prefix} site classic remove`,
  SITE_CLASSIC_SET: `${prefix} site classic set`,
  SITE_COMMSITE_ENABLE: `${prefix} site commsite enable`,
  SITE_GET: `${prefix} site get`,
  SITE_GROUPIFY: `${prefix} site groupify`,
  SITE_LIST: `${prefix} site list`,
  SITE_INPLACERECORDSMANAGEMENT_SET: `${prefix} site inplacerecordsmanagement set`,
  SITE_REMOVE: `${prefix} site remove`,
  SITE_RENAME: `${prefix} site rename`,
  SITE_SET: `${prefix} site set`,
  SITEDESIGN_ADD: `${prefix} sitedesign add`,
  SITEDESIGN_APPLY: `${prefix} sitedesign apply`,
  SITEDESIGN_GET: `${prefix} sitedesign get`,
  SITEDESIGN_LIST: `${prefix} sitedesign list`,
  SITEDESIGN_REMOVE: `${prefix} sitedesign remove`,
  SITEDESIGN_SET: `${prefix} sitedesign set`,
  SITEDESIGN_RIGHTS_GRANT: `${prefix} sitedesign rights grant`,
  SITEDESIGN_RIGHTS_LIST: `${prefix} sitedesign rights list`,
  SITEDESIGN_RIGHTS_REVOKE: `${prefix} sitedesign rights revoke`,
  SITEDESIGN_RUN_LIST: `${prefix} sitedesign run list`,
  SITEDESIGN_RUN_STATUS_GET: `${prefix} sitedesign run status get`,
  SITEDESIGN_TASK_GET: `${prefix} sitedesign task get`,
  SITEDESIGN_TASK_LIST: `${prefix} sitedesign task list`,
  SITEDESIGN_TASK_REMOVE: `${prefix} sitedesign task remove`,
  SITESCRIPT_ADD: `${prefix} sitescript add`,
  SITESCRIPT_GET: `${prefix} sitescript get`,
  SITESCRIPT_LIST: `${prefix} sitescript list`,
  SITESCRIPT_REMOVE: `${prefix} sitescript remove`,
  SITESCRIPT_SET: `${prefix} sitescript set`,
  SP_GRANT_ADD: `${prefix} sp grant add`,
  SP_GRANT_LIST: `${prefix} sp grant list`,
  SP_GRANT_REVOKE: `${prefix} sp grant revoke`,
  SP_PERMISSIONREQUEST_APPROVE: `${prefix} sp permissionrequest approve`,
  SP_PERMISSIONREQUEST_DENY: `${prefix} sp permissionrequest deny`,
  SP_PERMISSIONREQUEST_LIST: `${prefix} sp permissionrequest list`,
  SP_SET: `${prefix} sp set`,
  STORAGEENTITY_LIST: `${prefix} storageentity list`,
  STORAGEENTITY_GET: `${prefix} storageentity get`,
  STORAGEENTITY_SET: `${prefix} storageentity set`,
  STORAGEENTITY_REMOVE: `${prefix} storageentity remove`,
  TENANT_APPCATALOGURL_GET: `${prefix} tenant appcatalogurl get`,
  TENANT_RECYCLEBINITEM_LIST:  `${prefix} tenant recyclebinitem list`,
  TENANT_SETTINGS_LIST: `${prefix} tenant settings list`,
  TENANT_SETTINGS_SET: `${prefix} tenant settings set`,
  TERM_ADD: `${prefix} term add`,
  TERM_GET: `${prefix} term get`,
  TERM_LIST: `${prefix} term list`,
  TERM_GROUP_ADD: `${prefix} term group add`,
  TERM_GROUP_GET: `${prefix} term group get`,
  TERM_GROUP_LIST: `${prefix} term group list`,
  TERM_SET_ADD: `${prefix} term set add`,
  TERM_SET_GET: `${prefix} term set get`,
  TERM_SET_LIST: `${prefix} term set list`,
  THEME_APPLY: `${prefix} theme apply`,
  THEME_GET: `${prefix} theme get`,
  THEME_LIST: `${prefix} theme list`,
  THEME_REMOVE: `${prefix} theme remove`,
  THEME_SET: `${prefix} theme set`,
  WEB_ADD: `${prefix} web add`,
  WEB_CLIENTSIDEWEBPART_LIST: `${prefix} web clientsidewebpart list`,
  WEB_GET: `${prefix} web get`,
  WEB_LIST: `${prefix} web list`,
  WEB_REINDEX: `${prefix} web reindex`,
  WEB_REMOVE: `${prefix} web remove`,
  WEB_SET: `${prefix} web set`,
  USER_REMOVE:`${prefix} user remove`
};