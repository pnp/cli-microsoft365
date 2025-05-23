import assert from 'assert';
import chalk from 'chalk';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
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
import command from './tenant-settings-set.js';

describe(commands.TENANT_SETTINGS_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerStderrLogSpy: sinon.SinonSpy;

  const defaultRequestsSuccessStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1218", "ErrorInfo": null, "TraceCorrelationId": "6148899e-a042-6000-ee90-5bfa05d08b79"
          }, 4, {
            "IsNull": false
          }]);
      }
      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso-admin.sharepoint.com';
    auth.connection.spoTenantId = '6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant';
    commandInfo = cli.getCommandInfo(command);
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
    loggerStderrLogSpy = sinon.spy(logger, 'logToStderr');
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
    auth.connection.spoTenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('handles client.svc promise error', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('sets the tenant settings successfully', async () => {
    defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        NotificationsInSharePointEnabled: true
      }
    });
  });

  it('sends xml as array of strings for option excludedFileExtensionsForSyncClient', async () => {
    const request = defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        ExcludedFileExtensionsForSyncClient: 'xml,xslt,xsd'
      }
    });

    assert.strictEqual(request.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="ExcludedFileExtensionsForSyncClient"><Parameter Type="Array"><Object Type="String">xml</Object><Object Type="String">xslt</Object><Object Type="String">xsd</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);

  });

  it('sends xml as array of guids for option allowedDomainListForSyncClient', async () => {
    const request = defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        AllowedDomainListForSyncClient: '6648899e-a042-6000-ee90-5bfa05d08b79,6648899e-a042-6000-ee90-5bfa05d08b77'
      }
    });

    assert.strictEqual(request.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="AllowedDomainListForSyncClient"><Parameter Type="Array"><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b79}</Object><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b77}</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);

  });

  it('sends xml as array of guids for option disabledWebPartIds', async () => {
    const request = defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        DisabledWebPartIds: '6648899e-a042-6000-ee90-5bfa05d08b79,6648899e-a042-6000-ee90-5bfa05d08b77'
      }
    });

    assert.strictEqual(request.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="DisabledWebPartIds"><Parameter Type="Array"><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b79}</Object><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b77}</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
  });

  it('sends xml for multiple options specified', async () => {
    const request = defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        DisabledWebPartIds: '6648899e-a042-6000-ee90-5bfa05d08b79,6648899e-a042-6000-ee90-5bfa05d08b77',
        ExcludedFileExtensionsForSyncClient: 'xsl,doc,ttf',
        OfficeClientADALDisabled: true,
        OneDriveStorageQuota: 256,
        OrgNewsSiteUrl: 'https://contoso-admin.sharepoint.com'
      }
    });

    assert.strictEqual(request.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="DisabledWebPartIds"><Parameter Type="Array"><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b79}</Object><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b77}</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /><SetProperty Id="44" ObjectPathId="7" Name="ExcludedFileExtensionsForSyncClient"><Parameter Type="Array"><Object Type="String">xsl</Object><Object Type="String">doc</Object><Object Type="String">ttf</Object></Parameter></SetProperty><Method Name="Update" Id="45" ObjectPathId="7" /><SetProperty Id="46" ObjectPathId="7" Name="OfficeClientADALDisabled"><Parameter Type="String">true</Parameter></SetProperty><SetProperty Id="47" ObjectPathId="7" Name="OneDriveStorageQuota"><Parameter Type="String">256</Parameter></SetProperty><SetProperty Id="48" ObjectPathId="7" Name="OrgNewsSiteUrl"><Parameter Type="String">https://contoso-admin.sharepoint.com</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
  });

  it('handles tenant settings SelectAllProperties (first \'POST\') request error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
          }
        ]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('Timed out'));
  });

  it('handles tenant settings set (second \'POST\') request error', async () => {

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {

        if (opts.data.indexOf('SelectAllProperties') > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1218", "ErrorInfo": null, "TraceCorrelationId": "6148899e-a042-6000-ee90-5bfa05d08b79"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 1, "DefaultSharingLinkType": 1, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": true, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": true, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 66700, "ResourceQuotaAllocated": 13668, "RootSiteUrl": "https:\u002f\u002fprufinancial.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": "microsoft.com pramerica.ie pramericacdsdev.com prudential.com prufinancial.onmicrosoft.com", "SharingBlockedDomainList": "deloitte.com", "SharingCapability": 1, "SharingDomainRestrictionMode": 1, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": false, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 4448256, "StorageQuotaAllocated": 676508312, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": false, "HideDefaultThemes": true, "DisableCustomAppAuthentication": true, "CommentsOnListItemsDisabled": false
            }
          ]);
        }
        else {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
            }, 58, {
              "IsNull": false
            }, 59, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
            }
          ]);
        }
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Timed out'));
  });

  it('should turn enums to int in the request successfully', async () => {
    const stubRequest: sinon.SinonStub = defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        SharingCapability: 'ExternalUserSharingOnly',
        SharingDomainRestrictionMode: 'AllowList',
        DefaultSharingLinkType: 'Direct',
        ODBMembersCanShare: 'On',
        ODBAccessRequests: 'Off',
        FileAnonymousLinkType: 'View',
        FolderAnonymousLinkType: 'Edit',
        DefaultLinkPermission: 'View',
        ConditionalAccessPolicy: 'AllowLimitedAccess',
        LimitedAccessFileType: 'WebPreviewableFiles',
        SpecialCharactersStateInFileFolderNames: 'Allowed'
      }
    });
    assert.strictEqual(stubRequest.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="SharingCapability"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="43" ObjectPathId="7" Name="SharingDomainRestrictionMode"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="44" ObjectPathId="7" Name="DefaultSharingLinkType"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="45" ObjectPathId="7" Name="ODBMembersCanShare"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="46" ObjectPathId="7" Name="ODBAccessRequests"><Parameter Type="String">2</Parameter></SetProperty><SetProperty Id="47" ObjectPathId="7" Name="FileAnonymousLinkType"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="48" ObjectPathId="7" Name="FolderAnonymousLinkType"><Parameter Type="String">2</Parameter></SetProperty><SetProperty Id="49" ObjectPathId="7" Name="DefaultLinkPermission"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="50" ObjectPathId="7" Name="ConditionalAccessPolicy"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="51" ObjectPathId="7" Name="LimitedAccessFileType"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="52" ObjectPathId="7" Name="SpecialCharactersStateInFileFolderNames"><Parameter Type="String">1</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
  });

  it('validation fails if wrong enum value', async () => {
    const options: any = {
      SharingCapability: 'abc'
    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, 'SharingCapability option has invalid value of abc. Allowed values are ["Disabled","ExternalUserSharingOnly","ExternalUserAndGuestSharing","ExistingExternalUserSharingOnly"]');
  });

  it('validation passes if right enum value', async () => {
    const options: any = {
      debug: true,
      SharingCapability: 'ExternalUserSharingOnly',
      SharingDomainRestrictionMode: 'AllowList',
      DefaultSharingLinkType: 'Direct',
      ODBMembersCanShare: 'On',
      ODBAccessRequests: 'Off',
      FileAnonymousLinkType: 'View',
      FolderAnonymousLinkType: 'Edit',
      DefaultLinkPermission: 'View',
      ConditionalAccessPolicy: 'AllowLimitedAccess',
      LimitedAccessFileType: 'WebPreviewableFiles',
      SpecialCharactersStateInFileFolderNames: 'Allowed'
    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validation fails if wrong enum key', async () => {

    const actual = (command as any).mapEnumToInt('abc', 'abc');
    assert.strictEqual(actual, -1);
  });

  it('validation passes if right prop value', async () => {
    const options: any = {
      OrgNewsSiteUrl: 'abc'
    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validation fails if no options specified', async () => {
    const options: any = {
      debug: true,
      verbose: true
    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, `You must specify at least one option`);
  });

  it('validation passes autocomplete check if has the right value specified', async () => {
    const options: any = {
      ShowAllUsersClaim: true
    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows warning when option EnableAzureADB2BIntegration is used with value true', async () => {
    defaultRequestsSuccessStub();

    await command.action(logger, {
      options: {
        EnableAzureADB2BIntegration: true
      }
    });
    assert.strictEqual(loggerStderrLogSpy.calledWith(chalk.yellow("WARNING: Make sure to also enable the Microsoft Entra one-time passcode authentication preview. If it is not enabled then SharePoint will not use Microsoft Entra B2B even if EnableAzureADB2BIntegration is set to true. Learn more at http://aka.ms/spo-b2b-integration.")), true);
  });
});
