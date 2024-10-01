import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './contentcenter-list.js';

describe(commands.CONTENTCENTER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTCENTER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'Url']);
  });

  it('retrieves list of content centers', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url = `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`)) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">CONTENTCTR#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 101", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 1010", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 101", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }, {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 1010", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }
    ]));
  });

  it('retrieves list of all content centers when results returned in multiple pages', async () => {
    const postStub = sinon.stub(request, 'post');
    postStub.onFirstCall().callsFake(async (opts) => {
      if ((opts.url = `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`)) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">CONTENTCTR#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": "SPSiteQuery,841cb9d7-61a2-4029-b405-8cef77f591e2,924a239d-6416-49ff-86e2-0283b03bc4aa,0f820ed9-1927-4d48-8f88-94f863949574", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 101", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 1010", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    postStub.onSecondCall().callsFake(async (opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">SPSiteQuery,841cb9d7-61a2-4029-b405-8cef77f591e2,924a239d-6416-49ff-86e2-0283b03bc4aa,0f820ed9-1927-4d48-8f88-94f863949574</Property><Property Name="Template" Type="String">CONTENTCTR#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
          }, 2, {
            "IsNull": false
          }, 4, {
            "IsNull": false
          }, 5, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 101", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 1010", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
              }
            ]
          }
        ]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWith([
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 101", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }, {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 1010", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      },
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 101", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }, {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "CONTENTCTR#0", "TimeZoneId": 13, "Title": "Content Center 1010", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }
    ]));
  });

  it('correctly handles error when retrieving sites', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url = `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`)) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">CONTENTCTR#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {

          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": {
                "ErrorMessage": "Syntax error in the filter expression 'Url like 'test''.", "ErrorValue": null, "TraceCorrelationId": "3984379e-3011-4000-8240-a1114b993cad", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "3984379e-3011-4000-8240-a1114b993cad"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError("Syntax error in the filter expression 'Url like 'test''."));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('An error has occurred'));
  });
});
