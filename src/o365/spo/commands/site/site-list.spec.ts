import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import config from '../../../../config';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./site-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SITE_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth,
      request.get,
      (command as any).getRequestDigest
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SITE_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.SITE_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint tenant admin site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(`https://contoso.sharepoint.com is not a tenant admin site. Connect to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of modern teams sites when no type and no filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of modern teams sites when no type and no filter specified (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of modern teams sites when TeamSite type and no filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, type: 'TeamSite' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of communication sites when CommunicationSite type and no filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">SITEPAGEPUBLISHING#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "117d379e-10a9-4000-c5b4-0db32364b3ec"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "117d379e-10a9-4000-c5b4-0db32364b3ec|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fcomtest_001", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,6,59,233)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "admin@m365x324230.onmicrosoft.com", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 0, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "SITEPAGEPUBLISHING#0", "TimeZoneId": 13, "Title": "ComTest 001", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fcomtest_001", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, type: 'CommunicationSite' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'ComTest 001',
            Url: 'https://m365x324230.sharepoint.com/sites/comtest_001'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of modern teams sites when no type and filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like 'mtest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, filter: "Url -like 'mtest'" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of modern teams sites when TeamSite type and filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like 'mtest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, type: 'TeamSite', filter: "Url -like 'mtest'" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of communication sites when CommunicationSite type and filter specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like 'mtest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">SITEPAGEPUBLISHING#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "117d379e-10a9-4000-c5b4-0db32364b3ec"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "117d379e-10a9-4000-c5b4-0db32364b3ec|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fcomtest_001", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,6,59,233)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "admin@m365x324230.onmicrosoft.com", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 0, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "SITEPAGEPUBLISHING#0", "TimeZoneId": 13, "Title": "ComTest 001", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fcomtest_001", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, type: 'CommunicationSite', filter: "Url -like 'mtest'" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'ComTest 001',
            Url: 'https://m365x324230.sharepoint.com/sites/comtest_001'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in the filter', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like '&lt;mtest&gt;'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, filter: "Url -like '<mtest>'" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving sites', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url like 'mtest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, filter: "Url like 'mtest'" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError("Syntax error in the filter expression 'Url like 'test''.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('orders retrieved sites ascending by the URL', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">Url -like 'mtest'</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">GROUP#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                },
                // this shouldn't happen but it's included to test sorting
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_1010", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,17,46,0,910)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 1022361, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 1010", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_1010", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                },
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fm365x324230.sharepoint.com%2fsites%2fmtest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "GROUP#0", "TimeZoneId": 13, "Title": "Modern site 101", "Url": "https:\u002f\u002fm365x324230.sharepoint.com\u002fsites\u002fmtest_101", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                }
              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, filter: "Url -like 'mtest'" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Title: 'Modern site 101',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_101'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          },
          {
            Title: 'Modern site 1010',
            Url: 'https://m365x324230.sharepoint.com/sites/mtest_1010'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('accepts TeamSite site type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'TeamSite' } });
    assert(actual);
  });

  it('accepts CommunicationSite site type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'CommunicationSite' } });
    assert(actual);
  });

  it('rejects invalid site type', () => {
    const type = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: { type: type } });
    assert.equal(actual, `${type} is not a valid modern site type. Allowed types are TeamSite and CommunicationSite`);
  });

  it('doesn\'t fail validation if the optional type option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert(actual);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.SITE_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});