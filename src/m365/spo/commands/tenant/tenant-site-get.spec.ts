import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-site-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';

describe(commands.TENANT_SITE_GET, () => {
  const spoUrl = 'https://contoso.sharepoint.com';
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const id = '3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69';
  const title = 'Marketing';
  const url = 'https://contoso.sharepoint.com/sites/Marketing';
  const tenantSiteResponse = {
    value:
      [
        {
          "AllowDownloadingNonWebViewableFiles": false,
          "AllowEditing": true,
          "AllowFileArchive": false,
          "AllowSelfServiceUpgrade": true,
          "AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled": false,
          "AnonymousLinkExpirationInDays": 0,
          "ApplyToExistingDocumentLibraries": false,
          "ApplyToNewDocumentLibraries": false,
          "ArchivedBy": "",
          "ArchivedTime": "0001-01-01T00:00:00",
          "ArchiveStatus": "NotArchived",
          "AuthContextStrength": null,
          "AuthenticationContextLimitedAccess": false,
          "AuthenticationContextName": null,
          "AverageResourceUsage": 0,
          "BlockDownloadLinksFileType": 1,
          "BlockDownloadMicrosoft365GroupIds": null,
          "BlockDownloadPolicy": false,
          "BlockDownloadPolicyFileTypeIds": null,
          "BlockGuestsAsSiteAdmin": 0,
          "BonusDiskQuota": "233320",
          "ClearGroupId": false,
          "ClearRestrictedAccessControl": false,
          "CommentsOnSitePagesDisabled": false,
          "CompatibilityLevel": 15,
          "ConditionalAccessPolicy": 0,
          "CreatedTime": "2022-06-14T09:25:17.817",
          "CurrentResourceUsage": 0,
          "DefaultLinkPermission": 0,
          "DefaultLinkToExistingAccess": false,
          "DefaultLinkToExistingAccessReset": false,
          "DefaultShareLinkRole": 0,
          "DefaultShareLinkScope": -1,
          "DefaultSharingLinkType": 0,
          "DenyAddAndCustomizePages": 2,
          "Description": "",
          "DisableAppViews": 2,
          "DisableCompanyWideSharingLinks": 2,
          "DisableFlows": 2,
          "EnableAutoExpirationVersionTrim": false,
          "ExcludeBlockDownloadPolicySiteOwners": false,
          "ExcludeBlockDownloadSharePointGroups": [],
          "ExcludedBlockDownloadGroupIds": [],
          "ExpireVersionsAfterDays": 0,
          "ExternalUserExpirationInDays": 0,
          "GroupId": "bed1885e-4958-41e2-a091-b3a0dd418b1e",
          "GroupOwnerLoginName": "c:0o.c|federateddirectoryclaimprovider|bed1885e-4958-41e2-a091-b3a0dd418b1e_o",
          "HasHolds": false,
          "HidePeoplePreviewingFiles": false,
          "HidePeopleWhoHaveListsOpen": false,
          "HubSiteId": "00000000-0000-0000-0000-000000000000",
          "IBMode": "Open",
          "IBSegments": [],
          "IBSegmentsToAdd": null,
          "IBSegmentsToRemove": null,
          "InheritVersionPolicyFromTenant": true,
          "IsGroupOwnerSiteAdmin": true,
          "IsHubSite": false,
          "IsTeamsChannelConnected": false,
          "IsTeamsConnected": false,
          "LastContentModifiedDate": "2025-01-25T16:54:51.843",
          "Lcid": "1033",
          "LimitedAccessFileType": 1,
          "ListsShowHeaderAndNavigation": false,
          "LockIssue": null,
          "LockReason": 0,
          "LockState": "Unlock",
          "LoopDefaultSharingLinkRole": 0,
          "LoopDefaultSharingLinkScope": -1,
          "MajorVersionLimit": 0,
          "MajorWithMinorVersionsLimit": 0,
          "MediaTranscription": 0,
          "OverrideBlockUserInfoVisibility": 0,
          "OverrideSharingCapability": false,
          "OverrideTenantAnonymousLinkExpirationPolicy": false,
          "OverrideTenantExternalUserExpirationPolicy": false,
          "Owner": "bed1885e-4958-41e2-a091-b3a0dd418b1e_o",
          "OwnerEmail": "Marketing@contoso.onmicrosoft.com",
          "OwnerLoginName": "c:0o.c|federateddirectoryclaimprovider|bed1885e-4958-41e2-a091-b3a0dd418b1e_o",
          "OwnerName": "Marketing Owners",
          "PWAEnabled": 0,
          "ReadOnlyAccessPolicy": false,
          "ReadOnlyForBlockDownloadPolicy": false,
          "ReadOnlyForUnmanagedDevices": false,
          "RelatedGroupId": "bed1885e-4958-41e2-a091-b3a0dd418b1e",
          "RequestFilesLinkEnabled": false,
          "RequestFilesLinkExpirationInDays": -1,
          "RestrictContentOrgWideSearch": false,
          "RestrictedAccessControl": false,
          "RestrictedAccessControlGroups": [],
          "RestrictedAccessControlGroupsToAdd": null,
          "RestrictedAccessControlGroupsToRemove": null,
          "RestrictedToRegion": 3,
          "SandboxedCodeActivationCapability": 2,
          "SensitivityLabel": "00000000-0000-0000-0000-000000000000",
          "SensitivityLabel2": "",
          "SetOwnerWithoutUpdatingSecondaryAdmin": false,
          "SharingAllowedDomainList": "",
          "SharingBlockedDomainList": "",
          "SharingCapability": 1,
          "SharingDomainRestrictionMode": 0,
          "SharingLockDownCanBeCleared": true,
          "SharingLockDownEnabled": false,
          "ShowPeoplePickerSuggestionsForGuestUsers": false,
          "SiteDefinedSharingCapability": 1,
          "SiteId": "00000000-0000-0000-0000-000000000000",
          "SocialBarOnSitePagesDisabled": false,
          "Status": null,
          "StorageMaximumLevel": "26214400",
          "StorageQuotaType": null,
          "StorageUsage": "1",
          "StorageWarningLevel": "25574400",
          "TeamsChannelType": 0,
          "Template": "GROUP#0",
          "TimeZoneId": 23,
          "Title": "Marketing",
          "TitleTranslations": null,
          "Url": "https://contoso.sharepoint.com/sites/Marketing",
          "UserCodeMaximumLevel": 300,
          "UserCodeWarningLevel": 200,
          "VersionCount": "0",
          "VersionSize": "0",
          "WebsCount": 1
        }
      ]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
    sinon.stub(spo, 'getRequestDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    auth.connection.active = true;
    auth.connection.spoUrl = spoUrl;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => settingName === settingsNames.prompt ? false : defaultValue);
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
      request.get,
      spo.getSpoAdminUrl,
      spo.getAllSites,
      spo.getSiteAdminPropertiesByUrl,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SITE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when all options are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: { id: id, title: title, url: url }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when no options are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: {}
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and title options are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: { id: id, title: title }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and url options are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: { id: id, url: url }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and url options are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: { title: title, url: url }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: { url: 'invalid' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with only url', () => {
    const actual = commandOptionsSchema.safeParse({
      url: url
    });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves tenant site information by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/sites('3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69')`) {
        return tenantSiteResponse.value[0];
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { id: id, verbose: true }
    });

    assert(loggerLogSpy.calledWith(tenantSiteResponse.value[0]));
  });

  it('retrieves tenant site information by url', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves(tenantSiteResponse.value[0] as any);

    await command.action(logger, {
      options: { url: url, verbose: true }
    });

    assert(loggerLogSpy.calledWith(tenantSiteResponse.value[0]));
  });

  it('retrieves tenant site information by title', async () => {
    sinon.stub(spo, 'getAllSites').resolves([
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Marketing", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fMarketing", "SiteId": "\/Guid(3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69)\/", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }] as any
    );

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/sites('3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69')`) {
        return tenantSiteResponse.value[0];
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { title: title, verbose: true }
    });

    assert(loggerLogSpy.calledWith(tenantSiteResponse.value[0]));
  });

  it('handles selecting single result when multiple tenant sites with same title found and cli is set to prompt', async () => {
    sinon.stub(spo, 'getAllSites').resolves([
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Marketing", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fcMarketing", "SiteId": "\/Guid(3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69)\/", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      },
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "487c379e-80f8-4000-80be-1d37a4995717|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fctest_101", "AllowDownloadingNonWebViewableFiles": false, "AllowEditing": false, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 2, "DisableAppViews": 0, "DisableCompanyWideSharingLinks": 0, "DisableFlows": 0, "HasHolds": false, "LastContentModifiedDate": "\/Date(2017,11,17,4,12,28,997)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "", "OwnerEmail": null, "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 0, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 1, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 1, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 13, "Title": "Marketing", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fcMarketing2", "SiteId": "\/Guid(3ae83bc5-1f27-45c1-9eee-1bd1e2ddce68)\/", "UserCodeMaximumLevel": 300, "UserCodeWarningLevel": 200, "WebsCount": 0
      }
    ] as any
    );

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/sites('3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69')`) {
        return tenantSiteResponse.value[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ SiteId: '3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69' });

    await command.action(logger, {
      options: { title: 'Marketing', verbose: true }
    });

    assert(loggerLogSpy.calledWith(tenantSiteResponse.value[0]));
  });

  it('fails when the specified tenant site does not exist', async () => {
    sinon.stub(spo, 'getAllSites').resolves([] as any);

    await assert.rejects(command.action(logger, { options: { debug: true, title: title } } as any), new CommandError(`No site found with title '${title}'`));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: true, id: id } } as any), new CommandError('An error has occurred'));
  });
});