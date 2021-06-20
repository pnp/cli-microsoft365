import sinon = require("sinon");
import commands from "../commands";
import auth from "../../../Auth";
import Command, { CommandError } from "../../../Command";
import Utils from "../../../Utils";
import request from "../../../request";
import appInsights from "../../../appInsights";
import assert = require("assert");
import { Logger } from "../../../cli";
import config from "../../../config";
const command: Command = require('./onedrive-list');

describe(commands.LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
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
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'Url']);
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of OneDrive sites', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">1</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">SPSPERS</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.21402.12001", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2flidiah_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,15,517)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "lidiah@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Lidia Holloway", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002flidiah_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                },
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2fdiegos_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,44,240)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "diegos@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Diego Siciliani", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002fdiegos_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
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
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2flidiah_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,15,517)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "lidiah@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Lidia Holloway", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002flidiah_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          },
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2fdiegos_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,44,240)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "diegos@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Diego Siciliani", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002fdiegos_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list of OneDrive sites (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">1</Property><Property Name="StartIndex" Type="String">0</Property><Property Name="Template" Type="String">SPSPERS</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.21402.12001", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable", "NextStartIndex": -1, "NextStartIndexFromSharePoint": null, "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2flidiah_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,15,517)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "lidiah@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Lidia Holloway", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002flidiah_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
                },
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2fdiegos_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,44,240)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "diegos@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Diego Siciliani", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002fdiegos_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
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
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2flidiah_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,15,517)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "lidiah@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Lidia Holloway", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002flidiah_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
          },
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "2d63d39f-3016-0000-a532-30514e76ae73|908bed80-a04a-4433-b4a0-883d9847d110:d23a1d52-e19a-4bc5-be17-463a24e17fa2\nSiteProperties\nhttps%3a%2f%2fdev365-my.sharepoint.com%2fpersonal%2fdiegos_dev365_onmicrosoft_com", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AnonymousLinkExpirationInDays": 0, "AuthContextStrength": null, "AverageResourceUsage": 0, "BlockDownloadLinksFileType": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DefaultLinkPermission": 0, "DefaultLinkToExistingAccess": false, "DefaultLinkToExistingAccessReset": false, "DefaultSharingLinkType": 0, "DenyAddAndCustomizePages": 2, "Description": null, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "ExternalUserExpirationInDays": 0, "GroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "GroupOwnerLoginName": null, "HasHolds": false, "HubSiteId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "IBMode": null, "IBSegments": [], "IBSegmentsToAdd": null, "IBSegmentsToRemove": null, "IsGroupOwnerSiteAdmin": false, "IsHubSite": false, "LastContentModifiedDate": "\/Date(2021,3,1,16,45,44,240)\/", "Lcid": 1033, "LimitedAccessFileType": 0, "LockIssue": null, "LockState": "Unlock", "OverrideBlockUserInfoVisibility": 0, "OverrideTenantAnonymousLinkExpirationPolicy": false, "OverrideTenantExternalUserExpirationPolicy": false, "Owner": "diegos@dev365.onmicrosoft.com", "OwnerEmail": null, "OwnerLoginName": null, "OwnerName": null, "PWAEnabled": 1, "RelatedGroupId": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SensitivityLabel": "\/Guid(00000000-0000-0000-0000-000000000000)\/", "SensitivityLabel2": null, "SetOwnerWithoutUpdatingSecondaryAdmin": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 2, "SocialBarOnSitePagesDisabled": false, "Status": "Active", "StorageMaximumLevel": 1048576, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 943718, "Template": "SPSPERS#10", "TimeZoneId": 13, "Title": "Diego Siciliani", "Url": "https:\u002f\u002fdev365-my.sharepoint.com\u002fpersonal\u002fdiegos_dev365_onmicrosoft_com", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
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
          opts.headers['X-RequestDigest'] === 'abc') {

          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.21402.12001", "ErrorInfo": {
                "ErrorMessage": "FillSiteCollectionDTOInfo: Could not obtain valid templateId (-1) from provided template filter 'SPS-PERSONAL'", "ErrorValue": null, "TraceCorrelationId": "6ec7d39f-e0a0-0000-8a64-c2725ef5f458", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoException"
              }, "TraceCorrelationId": "6ec7d39f-e0a0-0000-8a64-c2725ef5f458"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("FillSiteCollectionDTOInfo: Could not obtain valid templateId (-1) from provided template filter 'SPS-PERSONAL'")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});