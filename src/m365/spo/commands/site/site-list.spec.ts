import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-list');

describe(commands.SITE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
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
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'Url']);
  });

  it('passes validation if type TeamSite specified', () => {
    const actual = command.validate({ options: { type: 'TeamSite' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if type CommunicationSite specified', () => {
    const actual = command.validate({ options: { type: 'CommunicationSite' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if type All specified', () => {
    const actual = command.validate({ options: { type: 'CommunicationSite' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if non existing type specified', () => {
    const actual = command.validate({ options: { type: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if type and webTemplate are both specified', () => {
    const actual = command.validate({ options: { type: 'TeamSite', webTemplate: 'STS#3' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if includeOneDriveSites is not specified together with type All', () => {
    const actual = command.validate({ options: { type: 'TeamSite', includeOneDriveSites: '1' } });
    assert.notStrictEqual(actual, true);
  });

  it('retrieves list of sites when no type and no filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when no type and no filter specified (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when type All is specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String"></Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, type: 'All' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when type CommunicationSite is specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">SITEPAGEPUBLISHING#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, type: 'CommunicationSite' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of deleted sites', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties><Property Name="NextStartIndexFromSharePoint" ScalarProperty="true" /></Properties></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="String">0</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, deleted: '1' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of all sites when results returned in multiple pages', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": "SPSiteQuery,841cb9d7-61a2-4029-b405-8cef77f591e2,924a239d-6416-49ff-86e2-0283b03bc4aa,0f820ed9-1927-4d48-8f88-94f863949574", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">SPSiteQuery,841cb9d7-61a2-4029-b405-8cef77f591e2,924a239d-6416-49ff-86e2-0283b03bc4aa,0f820ed9-1927-4d48-8f88-94f863949574</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          },
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('includes all properties for json output', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">STS#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, output: 'json', webTemplate: 'STS#0' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when STS#0 type and no filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">STS#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webTemplate: 'STS#0' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when no type and filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like 'ctest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, filter: "Url -like 'ctest'" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when type All, no filter and includeOneDriveSites specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">1</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String"></Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230-my.sharepoint.com%2fpersonal%2ffoo_bar_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Foo Bar", "Url": "https:\u002f\u002fm365x324230-my.sharepoint.com\u002fpersonal\u002ffoo_bar_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, type: 'All', includeOneDriveSites: '1' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230-my.sharepoint.com%2fpersonal%2ffoo_bar_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Foo Bar", "Url": "https:\u002f\u002fm365x324230-my.sharepoint.com\u002fpersonal\u002ffoo_bar_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of sites when STS#0 type and filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like 'ctest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">STS#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webTemplate: 'STS#0', filter: "Url -like 'ctest'" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Classic site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in the filter', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like '&lt;ctest&gt;'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "487c379e-80f8-4000-80be-1d37a4995717"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }, {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, filter: "Url -like '<ctest>'" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fctest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "STS#0", "TimeZoneId": 13, "Title": "Team site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fctest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving sites', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url like 'ctest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {

          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": {
                "ErrorMessage": "Syntax error in the filter expression 'Url like 'test''.", "ErrorValue": null, "TraceCorrelationId": "3984379e-3011-4000-8240-a1114b993cad", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "3984379e-3011-4000-8240-a1114b993cad"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, filter: "Url like 'ctest'" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Syntax error in the filter expression 'Url like 'test''.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: true, filter: "Url like 'ctest'" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });
});