import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./tenant-settings-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import config from '../../../../config';
import auth from '../../../../Auth';

describe(commands.TENANT_SETTINGS_SET, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let defaultRequestsSuccessStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1218", "ErrorInfo": null, "TraceCorrelationId": "6148899e-a042-6000-ee90-5bfa05d08b79"
          }, 4, {
            "IsNull": false
          }]));
      }
      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso-admin.sharepoint.com';
    auth.service.tenantId = '6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant';
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    auth.service.tenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_SETTINGS_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('handles client.svc promise error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the tenant settings (debug) successfully', (done) => {
    defaultRequestsSuccessStub();

    cmdInstance.action({
      options: {
        verbose: true,
        NotificationsInSharePointEnabled: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('DONE'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the tenant settings successfully', (done) => {
    defaultRequestsSuccessStub();

    cmdInstance.action({
      options: {
        NotificationsInSharePointEnabled: true
      }
    }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends xml as array of strings for option excludedFileExtensionsForSyncClient', (done) => {
    let request = defaultRequestsSuccessStub();

    cmdInstance.action({
      options: {
        ExcludedFileExtensionsForSyncClient: 'xml,xslt,xsd'
      }
    }, () => {
      try {
        assert.strictEqual(request.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="ExcludedFileExtensionsForSyncClient"><Parameter Type="Array"><Object Type="String">xml</Object><Object Type="String">xslt</Object><Object Type="String">xsd</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends xml as array of guids for option allowedDomainListForSyncClient', (done) => {
    let request = defaultRequestsSuccessStub();

    cmdInstance.action({
      options: {
        AllowedDomainListForSyncClient: '6648899e-a042-6000-ee90-5bfa05d08b79,6648899e-a042-6000-ee90-5bfa05d08b77'
      }
    }, () => {
      try {
        assert.strictEqual(request.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="AllowedDomainListForSyncClient"><Parameter Type="Array"><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b79}</Object><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b77}</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends xml as array of guids for option disabledWebPartIds', (done) => {
    let request = defaultRequestsSuccessStub();

    cmdInstance.action({
      options: {
        DisabledWebPartIds: '6648899e-a042-6000-ee90-5bfa05d08b79,6648899e-a042-6000-ee90-5bfa05d08b77'
      }
    }, () => {
      try {
        assert.strictEqual(request.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="DisabledWebPartIds"><Parameter Type="Array"><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b79}</Object><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b77}</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends xml for multiple options specified', (done) => {
    let request = defaultRequestsSuccessStub();

    cmdInstance.action({
      options: {
        DisabledWebPartIds: '6648899e-a042-6000-ee90-5bfa05d08b79,6648899e-a042-6000-ee90-5bfa05d08b77',
        ExcludedFileExtensionsForSyncClient: 'xsl,doc,ttf',
        OfficeClientADALDisabled: true,
        OneDriveStorageQuota: 256,
        OrgNewsSiteUrl: 'https://contoso-admin.sharepoint.com'
      }
    }, () => {
      try {
        assert.strictEqual(request.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="DisabledWebPartIds"><Parameter Type="Array"><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b79}</Object><Object Type="Guid">{6648899e-a042-6000-ee90-5bfa05d08b77}</Object></Parameter></SetProperty><Method Name="Update" Id="43" ObjectPathId="7" /><SetProperty Id="44" ObjectPathId="7" Name="ExcludedFileExtensionsForSyncClient"><Parameter Type="Array"><Object Type="String">xsl</Object><Object Type="String">doc</Object><Object Type="String">ttf</Object></Parameter></SetProperty><Method Name="Update" Id="45" ObjectPathId="7" /><SetProperty Id="46" ObjectPathId="7" Name="OfficeClientADALDisabled"><Parameter Type="String">true</Parameter></SetProperty><SetProperty Id="47" ObjectPathId="7" Name="OneDriveStorageQuota"><Parameter Type="String">256</Parameter></SetProperty><SetProperty Id="48" ObjectPathId="7" Name="OrgNewsSiteUrl"><Parameter Type="String">https://contoso-admin.sharepoint.com</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles tenant settings SelectAllProperties (first \'POST\') request error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
          }
        ]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Timed out')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles tenant settings set (second \'POST\') request error', (done) => {

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {

        if (opts.body.indexOf('SelectAllProperties') > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1218", "ErrorInfo": null, "TraceCorrelationId": "6148899e-a042-6000-ee90-5bfa05d08b79"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 1, "DefaultSharingLinkType": 1, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": true, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": true, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 66700, "ResourceQuotaAllocated": 13668, "RootSiteUrl": "https:\u002f\u002fprufinancial.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": "microsoft.com pramerica.ie pramericacdsdev.com prudential.com prufinancial.onmicrosoft.com", "SharingBlockedDomainList": "deloitte.com", "SharingCapability": 1, "SharingDomainRestrictionMode": 1, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": false, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 4448256, "StorageQuotaAllocated": 676508312, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": false, "HideDefaultThemes": true
            }
          ]));
        } else {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
            }, 58, {
              "IsNull": false
            }, 59, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
            }
          ]));
        }
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Timed out')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should turn enums to int in the request successfully', (done) => {
    const stubRequest: sinon.SinonStub = defaultRequestsSuccessStub();

    cmdInstance.action({
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
    }, () => {
      try {
        assert.strictEqual(stubRequest.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="42" ObjectPathId="7" Name="SharingCapability"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="43" ObjectPathId="7" Name="SharingDomainRestrictionMode"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="44" ObjectPathId="7" Name="DefaultSharingLinkType"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="45" ObjectPathId="7" Name="ODBMembersCanShare"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="46" ObjectPathId="7" Name="ODBAccessRequests"><Parameter Type="String">2</Parameter></SetProperty><SetProperty Id="47" ObjectPathId="7" Name="FileAnonymousLinkType"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="48" ObjectPathId="7" Name="FolderAnonymousLinkType"><Parameter Type="String">2</Parameter></SetProperty><SetProperty Id="49" ObjectPathId="7" Name="DefaultLinkPermission"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="50" ObjectPathId="7" Name="ConditionalAccessPolicy"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="51" ObjectPathId="7" Name="LimitedAccessFileType"><Parameter Type="String">1</Parameter></SetProperty><SetProperty Id="52" ObjectPathId="7" Name="SpecialCharactersStateInFileFolderNames"><Parameter Type="String">1</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="7" Name="6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee&#xA;Tenant" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('validation fails if wrong enum value', () => {
    const options: any = {
      SharingCapability: 'abc'
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, 'SharingCapability option has invalid value of abc. Allowed values are ["Disabled","ExternalUserSharingOnly","ExternalUserAndGuestSharing","ExistingExternalUserSharingOnly"]');
  });

  it('validation passes if right enum value', () => {
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
      SpecialCharactersStateInFileFolderNames: 'Allowed',
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, true);
  });

  it('validation fails if wrong enum value', () => {
    const options: any = {
      SharingCapability: 'abc'
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, 'SharingCapability option has invalid value of abc. Allowed values are ["Disabled","ExternalUserSharingOnly","ExternalUserAndGuestSharing","ExistingExternalUserSharingOnly"]');
  });

  it('validation fails if wrong enum key', () => {

    const actual = (command as any).mapEnumToInt('abc', 'abc');
    assert.strictEqual(actual, -1);
  });

  it('validation passes if right prop value', () => {
    const options: any = {
      OrgNewsSiteUrl: 'abc'
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, true);
  });

  it('validation false if boolean option has non boolean value', () => {
    const options: any = {
      ShowAllUsersClaim: 'abc'
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, 'ShowAllUsersClaim option has invalid value of abc. Allowed values are ["true","false"]');
  });

  it('validation fails if no options specified', () => {
    const options: any = {
      debug: true,
      verbose: true
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, `You must specify at least one option`);
  });

  it('validation passes autocomplete check if has the right value specified', () => {
    const options: any = {
      ShowAllUsersClaim: true
    }
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, true);
  });
});