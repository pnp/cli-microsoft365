import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-site-get.js';
import { spo, TenantSiteProperties } from '../../../../utils/spo.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TENANT_SITE_GET, () => {
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const siteId = '3ae83bc5-1f27-45c1-9eee-1bd1e2ddce69';
  const siteUrl = 'https://contoso.sharepoint.com/sites/Marketing';

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const siteResponse: TenantSiteProperties = {
    AllowDownloadingNonWebViewableFiles: false,
    AllowEditing: true,
    AllowFileArchive: false,
    AllowSelfServiceUpgrade: true,
    AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled: false,
    AnonymousLinkExpirationInDays: 0,
    ApplyToExistingDocumentLibraries: false,
    ApplyToNewDocumentLibraries: false,
    ArchivedBy: "",
    ArchivedFileDiskUsed: "0",
    ArchivedTime: "0001-01-01T00:00:00",
    ArchiveStatus: "NotArchived",
    AuthContextStrength: null,
    AuthenticationContextLimitedAccess: false,
    AuthenticationContextName: null,
    AverageResourceUsage: 0,
    BlockDownloadLinksFileType: 1,
    BlockDownloadMicrosoft365GroupIds: null,
    BlockDownloadPolicy: false,
    BlockDownloadPolicyFileTypeIds: null,
    BlockGuestsAsSiteAdmin: 0,
    BonusDiskQuota: "920",
    ClearGroupId: false,
    ClearRestrictedAccessControl: false,
    CommentsOnSitePagesDisabled: false,
    CompatibilityLevel: 15,
    ConditionalAccessPolicy: 0,
    CreatedTime: "2021-10-12T09:54:16.52",
    CurrentResourceUsage: 0,
    DefaultLinkPermission: 0,
    DefaultLinkToExistingAccess: false,
    DefaultLinkToExistingAccessReset: false,
    DefaultShareLinkRole: 0,
    DefaultShareLinkScope: -1,
    DefaultSharingLinkType: 0,
    DenyAddAndCustomizePages: 2,
    Description: "",
    DisableAppViews: 2,
    DisableCompanyWideSharingLinks: 2,
    DisableFlows: 2,
    DisableSiteBranding: false,
    EnableAutoExpirationVersionTrim: false,
    ExcludeBlockDownloadPolicySiteOwners: false,
    ExcludeBlockDownloadSharePointGroups: [],
    ExcludedBlockDownloadGroupIds: [],
    ExpireVersionsAfterDays: 0,
    ExternalUserExpirationInDays: 0,
    FileTypesForVersionExpiration: null,
    GroupId: "00000000-0000-0000-0000-000000000000",
    GroupOwnerLoginName: "c:0o.c|federateddirectoryclaimprovider|00000000-0000-0000-0000-000000000000_o",
    HasHolds: false,
    HidePeoplePreviewingFiles: false,
    HidePeopleWhoHaveListsOpen: false,
    HubSiteId: "af80c11f-0138-4d72-bb37-514542c3aabb",
    IBMode: "",
    IBSegments: [],
    IBSegmentsToAdd: null,
    IBSegmentsToRemove: null,
    InheritVersionPolicyFromTenant: true,
    IsAuthoritative: false,
    IsGroupOwnerSiteAdmin: false,
    IsHubSite: false,
    IsTeamsChannelConnected: false,
    IsTeamsConnected: false,
    LastContentModifiedDate: "2025-10-03T00:20:28.62",
    Lcid: "1033",
    LimitedAccessFileType: 1,
    ListsShowHeaderAndNavigation: false,
    LockIssue: null,
    LockReason: 0,
    LockState: "Unlock",
    LoopDefaultSharingLinkRole: 0,
    LoopDefaultSharingLinkScope: -1,
    MajorVersionLimit: 0,
    MajorWithMinorVersionsLimit: 0,
    MediaTranscription: 0,
    OverrideBlockUserInfoVisibility: 0,
    OverrideSharingCapability: false,
    OverrideTenantAnonymousLinkExpirationPolicy: false,
    OverrideTenantExternalUserExpirationPolicy: false,
    Owner: "john@contoso.onmicrosoft.com",
    OwnerEmail: "john@contoso.onmicrosoft.com",
    OwnerLoginName: "i:0#.f|membership|john@contoso.onmicrosoft.com",
    OwnerName: "john",
    PWAEnabled: 1,
    ReadOnlyAccessPolicy: false,
    ReadOnlyForBlockDownloadPolicy: false,
    ReadOnlyForUnmanagedDevices: false,
    RelatedGroupId: "00000000-0000-0000-0000-000000000000",
    RemoveVersionExpirationFileTypeOverride: null,
    RequestFilesLinkEnabled: false,
    RequestFilesLinkExpirationInDays: -1,
    RestrictContentOrgWideSearch: false,
    RestrictedAccessControl: false,
    RestrictedAccessControlGroups: [],
    RestrictedAccessControlGroupsToAdd: null,
    RestrictedAccessControlGroupsToRemove: null,
    RestrictedContentDiscoveryforCopilotAndAgents: false,
    RestrictedToRegion: 3,
    SandboxedCodeActivationCapability: 2,
    SensitivityLabel: "00000000-0000-0000-0000-000000000000",
    SensitivityLabel2: null,
    SetOwnerWithoutUpdatingSecondaryAdmin: false,
    SharingAllowedDomainList: "",
    SharingBlockedDomainList: "",
    SharingCapability: 0,
    SharingDomainRestrictionMode: 0,
    SharingLockDownCanBeCleared: true,
    SharingLockDownEnabled: false,
    ShowPeoplePickerSuggestionsForGuestUsers: false,
    SiteDefinedSharingCapability: 0,
    SiteId: "8f6fdeda-f6ff-4d39-8a8c-fe86565afefd",
    SocialBarOnSitePagesDisabled: false,
    Status: "Active",
    StorageMaximumLevel: "26214400",
    StorageQuotaType: null,
    StorageUsage: "3",
    StorageWarningLevel: "25574400",
    TeamsChannelType: 0,
    Template: "SITEPAGEPUBLISHING#0",
    TimeZoneId: 2,
    Title: "Marketing and Communications",
    TitleTranslations: null,
    Url: "https://contoso.sharepoint.com/sites/marketing",
    UserCodeMaximumLevel: 300,
    UserCodeWarningLevel: 200,
    VersionCount: "7",
    VersionPolicyFileTypeOverride: [],
    VersionSize: "0",
    WebsCount: 1
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });
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
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.handleMultipleResultsFound,
      spo.getSpoAdminUrl,
      spo.getSiteAdminPropertiesByUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SITE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves site by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=Url`) {
        return { Url: siteUrl };
      }
      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves(siteResponse);

    await command.action(logger, { options: { id: siteId, verbose: true } });
    assert(loggerLogSpy.calledWith(siteResponse));
  });

  it('retrieves site by url', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves(siteResponse);

    await command.action(logger, { options: { url: siteUrl, verbose: true } });
    assert(loggerLogSpy.calledWith(siteResponse));
  });

  it('retrieves site by title (single result)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) {
        return { Row: [{ Title: 'Marketing', SiteUrl: siteUrl, SiteId: `/Guid(${siteId})/` }] };
      }
      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves(siteResponse);

    await command.action(logger, { options: { title: 'Marketing', verbose: true } });
    assert(loggerLogSpy.calledWith(siteResponse));
  });

  it('retrieves site by title (multiple results prompts)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) {
        return {
          Row: [
            { Title: 'Marketing', SiteUrl: siteUrl, SiteId: `/Guid(${siteId})/` },
            { Title: 'Marketing', SiteUrl: 'https://contoso.sharepoint.com/sites/Marketing2', SiteId: '/Guid(53dec431-9d4f-415b-b12b-010259d5b4e1)/' }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ url: siteUrl } as any);

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves(siteResponse);

    await command.action(logger, { options: { title: 'Marketing', verbose: true } });
    assert(loggerLogSpy.calledWith(siteResponse));
  });

  it('handles error when specified site by title not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) {
        return { Row: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: 'Marketing', verbose: true } } as any), new CommandError("The specified site 'Marketing' does not exist."));
  });

  it('fails validation when specifying none of id, title, url', async () => {
    const commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    const refined = commandInfo.command.getRefinedSchema!(commandOptionsSchema as any)!;
    const actual = refined.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when specifying multiple of id, title, url', async () => {
    const commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    const refined = commandInfo.command.getRefinedSchema!(commandOptionsSchema as any)!;
    const actual = refined.safeParse({ id: siteId, url: siteUrl });
    assert.strictEqual(actual.success, false);
  });

  it('handles OData error when site not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=Url`) {
        const error = {
          error: {
            'odata.error': {
              code: '-2147024891, System.UnauthorizedAccessException',
              message: {
                lang: 'en-US',
                value: 'Attempted to perform an unauthorized operation.'
              }
            }
          }
        };
        throw error;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: siteId, verbose: true } } as any), new CommandError('Attempted to perform an unauthorized operation.'));
  });
});


