import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import * as request from 'request-promise-native';
import config from '../../../../config';
import commands from '../../commands';
import Utils from '../../../../Utils';
import {
  CommandOption,
  CommandError,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
const vorpal: Vorpal = require('../../../../vorpal-init');

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  minCompatibilityLevel: number;
  maxCompatibilityLevel: number;
  externalServicesEnabled: boolean;
  noAccessRedirectUrl: string;
  sharingCapability: string; // <SharingCapabilities>
  displayStartASiteOption: boolean;
  startASiteFormUrl: string;
  showEveryoneClaim: boolean;
  showAllUsersClaim: boolean;
  showEveryoneExceptExternalUsersClaim: boolean;
  searchResolveExactEmailOrUPN: boolean;
  officeClientADALDisabled: boolean;
  legacyAuthProtocolsEnabled: boolean;
  requireAcceptingAccountMatchInvitedAccount: boolean;
  provisionSharedWithEveryoneFolder: boolean;
  signInAccelerationDomain: string;
  enableGuestSignInAcceleration: boolean;
  usePersistentCookiesForExplorerView: boolean;
  bccExternalSharingInvitations: boolean;
  bccExternalSharingInvitationsList: string;
  userVoiceForFeedbackEnabled: boolean;
  publicCdnEnabled: boolean;
  publicCdnAllowedFileTypes: string;
  requireAnonymousLinksExpireInDays: number;
  sharingAllowedDomainList: string;
  sharingBlockedDomainList: string;
  sharingDomainRestrictionMode: string; // <SharingDomainRestrictionModes>
  oneDriveStorageQuota: number;
  oneDriveForGuestsEnabled: boolean;
  iPAddressEnforcement: boolean;
  iPAddressAllowList: string;
  iPAddressWACTokenLifetime: number;
  useFindPeopleInPeoplePicker: boolean;
  defaultSharingLinkType: string; // <SharingLinkType>
  oDBMembersCanShare: string; // <SharingState>
  oDBAccessRequests: string; // <SharingState>
  preventExternalUsersFromResharing: boolean;
  showPeoplePickerSuggestionsForGuestUsers: boolean;
  fileAnonymousLinkType: string; // <AnonymousLinkType>
  folderAnonymousLinkType: string; // <AnonymousLinkType>
  notifyOwnersWhenItemsReshared: boolean;
  notifyOwnersWhenInvitationsAccepted: boolean;
  notificationsInOneDriveForBusinessEnabled: boolean;
  notificationsInSharePointEnabled: boolean;
  ownerAnonymousNotification: boolean;
  commentsOnSitePagesDisabled: boolean;
  socialBarOnSitePagesDisabled: boolean;
  orphanedPersonalSitesRetentionPeriod: number;
  disallowInfectedFileDownload: boolean;
  defaultLinkPermission: string; // <SharingPermissionType>
  conditionalAccessPolicy: string; // <SPOConditionalAccessPolicyType>
  allowDownloadingNonWebViewableFiles: boolean;
  allowEditing: boolean;
  applyAppEnforcedRestrictionsToAdHocRecipients: boolean;
  filePickerExternalImageSearchEnabled: boolean;
  emailAttestationRequired: boolean;
  emailAttestationReAuthDays: number;
  hideDefaultThemes: boolean;
  // not included in the PnP PowerShell, most of them are new and maybe the cmdlet is not updated recently.
  blockAccessOnUnmanagedDevices: boolean;
  allowLimitedAccessOnUnmanagedDevices: boolean;
  blockDownloadOfAllFilesForGuests: boolean;
  blockDownloadOfAllFilesOnUnmanagedDevices: boolean;
  blockDownloadOfViewableFilesForGuests: boolean;
  blockDownloadOfViewableFilesOnUnmanagedDevices: boolean;
  blockMacSync: boolean;
  disableReportProblemDialog: boolean;
  displayNamesOfFileViewers: boolean;
  enableMinimumVersionRequirement: boolean;
  hideSyncButtonOnODB: boolean;
  isUnmanagedSyncClientForTenantRestricted: boolean;
  limitedAccessFileType: string; // <LimitedAccessFileType>
  optOutOfGrooveBlock: boolean;
  optOutOfGrooveSoftBlock: boolean;
  orgNewsSiteUrl: string;
  permissiveBrowserFileHandlingOverride: boolean;
  showNGSCDialogForSyncOnODB: boolean;
  specialCharactersStateInFileFolderNames: string; // <SpecialCharactersState>
  syncPrivacyProfileProperties: boolean;
  excludedFileExtensionsForSyncClient: string[];
  allowedDomainListForSyncClient: string[];
  disabledWebPartIds: string[];
}

class SpoTenantSettingsSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets tenant global setting';
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.minCompatibilityLevel = (!(!args.options.minCompatibilityLevel)).toString();
    telemetryProps.maxCompatibilityLevel = (!(!args.options.maxCompatibilityLevel)).toString();
    telemetryProps.externalServicesEnabled = (!(!args.options.externalServicesEnabled)).toString();
    telemetryProps.noAccessRedirectUrl = (!(!args.options.noAccessRedirectUrl)).toString();
    telemetryProps.sharingCapability = (!(!args.options.sharingCapability)).toString();
    telemetryProps.displayStartASiteOption = (!(!args.options.displayStartASiteOption)).toString();
    telemetryProps.startASiteFormUrl = (!(!args.options.startASiteFormUrl)).toString();
    telemetryProps.showEveryoneClaim = (!(!args.options.showEveryoneClaim)).toString();
    telemetryProps.showAllUsersClaim = (!(!args.options.showAllUsersClaim)).toString();
    telemetryProps.showEveryoneExceptExternalUsersClaim = (!(!args.options.showEveryoneExceptExternalUsersClaim)).toString();
    telemetryProps.searchResolveExactEmailOrUPN = (!(!args.options.searchResolveExactEmailOrUPN)).toString();
    telemetryProps.officeClientADALDisabled = (!(!args.options.officeClientADALDisabled)).toString();
    telemetryProps.legacyAuthProtocolsEnabled = (!(!args.options.legacyAuthProtocolsEnabled)).toString();
    telemetryProps.requireAcceptingAccountMatchInvitedAccount = (!(!args.options.requireAcceptingAccountMatchInvitedAccount)).toString();
    telemetryProps.provisionSharedWithEveryoneFolder = (!(!args.options.provisionSharedWithEveryoneFolder)).toString();
    telemetryProps.signInAccelerationDomain = (!(!args.options.signInAccelerationDomain)).toString();
    telemetryProps.enableGuestSignInAcceleration = (!(!args.options.enableGuestSignInAcceleration)).toString();
    telemetryProps.usePersistentCookiesForExplorerView = (!(!args.options.usePersistentCookiesForExplorerView)).toString();
    telemetryProps.bccExternalSharingInvitations = (!(!args.options.bccExternalSharingInvitations)).toString();
    telemetryProps.bccExternalSharingInvitationsList = (!(!args.options.bccExternalSharingInvitationsList)).toString();
    telemetryProps.userVoiceForFeedbackEnabled = (!(!args.options.userVoiceForFeedbackEnabled)).toString();
    telemetryProps.publicCdnEnabled = (!(!args.options.publicCdnEnabled)).toString();
    telemetryProps.publicCdnAllowedFileTypes = (!(!args.options.publicCdnAllowedFileTypes)).toString();
    telemetryProps.requireAnonymousLinksExpireInDays = (!(!args.options.requireAnonymousLinksExpireInDays)).toString();
    telemetryProps.sharingAllowedDomainList = (!(!args.options.sharingAllowedDomainList)).toString();
    telemetryProps.sharingBlockedDomainList = (!(!args.options.sharingBlockedDomainList)).toString();
    telemetryProps.sharingDomainRestrictionMode = (!(!args.options.sharingDomainRestrictionMode)).toString();
    telemetryProps.oneDriveStorageQuota = (!(!args.options.oneDriveStorageQuota)).toString();
    telemetryProps.oneDriveForGuestsEnabled = (!(!args.options.oneDriveForGuestsEnabled)).toString();
    telemetryProps.iPAddressEnforcement = (!(!args.options.iPAddressEnforcement)).toString();
    telemetryProps.iPAddressAllowList = (!(!args.options.iPAddressAllowList)).toString();
    telemetryProps.iPAddressWACTokenLifetime = (!(!args.options.iPAddressWACTokenLifetime)).toString();
    telemetryProps.useFindPeopleInPeoplePicker = (!(!args.options.useFindPeopleInPeoplePicker)).toString();
    telemetryProps.defaultSharingLinkType = (!(!args.options.defaultSharingLinkType)).toString();
    telemetryProps.oDBMembersCanShare = (!(!args.options.oDBMembersCanShare)).toString();
    telemetryProps.oDBAccessRequests = (!(!args.options.oDBAccessRequests)).toString();
    telemetryProps.preventExternalUsersFromResharing = (!(!args.options.preventExternalUsersFromResharing)).toString();
    telemetryProps.showPeoplePickerSuggestionsForGuestUsers = (!(!args.options.showPeoplePickerSuggestionsForGuestUsers)).toString();
    telemetryProps.fileAnonymousLinkType = (!(!args.options.fileAnonymousLinkType)).toString();
    telemetryProps.folderAnonymousLinkType = (!(!args.options.folderAnonymousLinkType)).toString();
    telemetryProps.notifyOwnersWhenItemsReshared = (!(!args.options.notifyOwnersWhenItemsReshared)).toString();
    telemetryProps.notifyOwnersWhenInvitationsAccepted = (!(!args.options.notifyOwnersWhenInvitationsAccepted)).toString();
    telemetryProps.notificationsInOneDriveForBusinessEnabled = (!(!args.options.notificationsInOneDriveForBusinessEnabled)).toString();
    telemetryProps.notificationsInSharePointEnabled = (!(!args.options.notificationsInSharePointEnabled)).toString();
    telemetryProps.ownerAnonymousNotification = (!(!args.options.ownerAnonymousNotification)).toString();
    telemetryProps.commentsOnSitePagesDisabled = (!(!args.options.commentsOnSitePagesDisabled)).toString();
    telemetryProps.socialBarOnSitePagesDisabled = (!(!args.options.socialBarOnSitePagesDisabled)).toString();
    telemetryProps.orphanedPersonalSitesRetentionPeriod = (!(!args.options.orphanedPersonalSitesRetentionPeriod)).toString();
    telemetryProps.disallowInfectedFileDownload = (!(!args.options.disallowInfectedFileDownload)).toString();
    telemetryProps.defaultLinkPermission = (!(!args.options.defaultLinkPermission)).toString();
    telemetryProps.conditionalAccessPolicy = (!(!args.options.conditionalAccessPolicy)).toString();
    telemetryProps.allowDownloadingNonWebViewableFiles = (!(!args.options.allowDownloadingNonWebViewableFiles)).toString();
    telemetryProps.allowEditing = (!(!args.options.allowEditing)).toString();
    telemetryProps.applyAppEnforcedRestrictionsToAdHocRecipients = (!(!args.options.applyAppEnforcedRestrictionsToAdHocRecipients)).toString();
    telemetryProps.filePickerExternalImageSearchEnabled = (!(!args.options.filePickerExternalImageSearchEnabled)).toString();
    telemetryProps.emailAttestationRequired = (!(!args.options.emailAttestationRequired)).toString();
    telemetryProps.emailAttestationReAuthDays = (!(!args.options.emailAttestationReAuthDays)).toString();
    telemetryProps.hideDefaultThemes = (!(!args.options.hideDefaultThemes)).toString();
    telemetryProps.blockAccessOnUnmanagedDevices = (!(!args.options.blockAccessOnUnmanagedDevices)).toString();
    telemetryProps.allowLimitedAccessOnUnmanagedDevices = (!(!args.options.allowLimitedAccessOnUnmanagedDevices)).toString();
    telemetryProps.blockDownloadOfAllFilesForGuests = (!(!args.options.blockDownloadOfAllFilesForGuests)).toString();
    telemetryProps.blockDownloadOfAllFilesOnUnmanagedDevices = (!(!args.options.blockDownloadOfAllFilesOnUnmanagedDevices)).toString();
    telemetryProps.blockDownloadOfViewableFilesForGuests = (!(!args.options.blockDownloadOfViewableFilesForGuests)).toString();
    telemetryProps.blockDownloadOfViewableFilesOnUnmanagedDevices = (!(!args.options.blockDownloadOfViewableFilesOnUnmanagedDevices)).toString();
    telemetryProps.blockMacSync = (!(!args.options.blockMacSync)).toString();
    telemetryProps.disableReportProblemDialog = (!(!args.options.disableReportProblemDialog)).toString();
    telemetryProps.displayNamesOfFileViewers = (!(!args.options.displayNamesOfFileViewers)).toString();
    telemetryProps.enableMinimumVersionRequirement = (!(!args.options.enableMinimumVersionRequirement)).toString();
    telemetryProps.hideSyncButtonOnODB = (!(!args.options.hideSyncButtonOnODB)).toString();
    telemetryProps.isUnmanagedSyncClientForTenantRestricted = (!(!args.options.isUnmanagedSyncClientForTenantRestricted)).toString();
    telemetryProps.limitedAccessFileType = (!(!args.options.limitedAccessFileType)).toString();
    telemetryProps.optOutOfGrooveBlock = (!(!args.options.optOutOfGrooveBlock)).toString();
    telemetryProps.optOutOfGrooveSoftBlock = (!(!args.options.optOutOfGrooveSoftBlock)).toString();
    telemetryProps.orgNewsSiteUrl = (!(!args.options.orgNewsSiteUrl)).toString();
    telemetryProps.permissiveBrowserFileHandlingOverride = (!(!args.options.permissiveBrowserFileHandlingOverride)).toString();
    telemetryProps.showNGSCDialogForSyncOnODB = (!(!args.options.showNGSCDialogForSyncOnODB)).toString();
    telemetryProps.specialCharactersStateInFileFolderNames = (!(!args.options.specialCharactersStateInFileFolderNames)).toString();
    telemetryProps.syncPrivacyProfileProperties = (!(!args.options.syncPrivacyProfileProperties)).toString();
    telemetryProps.excludedFileExtensionsForSyncClient = (!(!args.options.excludedFileExtensionsForSyncClient)).toString();
    telemetryProps.disabledWebPartIds = (!(!args.options.disabledWebPartIds)).toString();
    telemetryProps.allowedDomainListForSyncClient = (!(!args.options.allowedDomainListForSyncClient)).toString();
    return telemetryProps;
  }

  public getAllEnumOptions(): string[] {
    return ['sharingCapability', 'sharingDomainRestrictionMode', 'defaultSharingLinkType', 'oDBMembersCanShare', 'oDBAccessRequests', 'fileAnonymousLinkType', 'folderAnonymousLinkType', 'defaultLinkPermission', 'conditionalAccessPolicy', 'limitedAccessFileType', 'specialCharactersStateInFileFolderNames'];
  }

  // all enums as get methods
  public getSharingLinkType(): string[] { return ['None', 'Direct', 'Internal', 'AnonymousAccess']; }
  public getSharingCapabilities(): string[] { return ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']; }
  public getSharingDomainRestrictionModes(): string[] { return ['None', 'AllowList', 'BlockList'] };
  public getSharingState(): string[] { return ['Unspecified', 'On', 'Off']; }
  public getAnonymousLinkType(): string[] { return ['None', 'View', 'Edit']; }
  public getSharingPermissionType(): string[] { return ['None', 'View', 'Edit']; }
  public getSPOConditionalAccessPolicyType(): string[] { return ['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess']; }
  public getSpecialCharactersState(): string[] { return ['NoPreference', 'Allowed', 'Disallowed']; }
  public getSPOLimitedAccessFileType(): string[] { return ['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles']; }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    let accessToken = '';
    let formDigestValue = '';

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((resp: string): request.RequestPromise => {
        accessToken = resp;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((contextResponse: ContextInfo): request.RequestPromise => {

        if (this.debug) {
          cmd.log('Retrieved ContextInfo...');
          cmd.log(contextResponse);
          cmd.log('');
        }

        formDigestValue = contextResponse.FormDigestValue;

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': formDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Retrieved ContextInfo...');
          cmd.log(contextResponse);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<any> => {

        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          const error: string = response.ErrorInfo.ErrorMessage;
          return new Promise<any>((resolve, reject) => { reject(error) });
        }

        // map the args.options to XML Properties
        let propsXml: string = '';
        let id = 42; // geek's humor
        for (let optionKey of Object.keys(args.options)) {

          if (this.isExcludedOption(optionKey)) {
            continue;
          }

          let optionValue = args.options[optionKey];
          if (this.getAllEnumOptions().indexOf(optionKey) > -1) {
            // map enum values to int
            optionValue = this.mapEnumToInt(optionKey, args.options[optionKey]);
          }

          if (['allowedDomainListForSyncClient', 'disabledWebPartIds'].indexOf(optionKey) > -1) {

            // the XML has to be represented as array of guids
            let valuesXml = '';
            optionValue.split(',').forEach((value: string) => {
              valuesXml += `<Object Type="Guid">{${value}}</Object>`;
            })
            propsXml += `<SetProperty Id="${id}" ObjectPathId="7" Name="${optionKey[0].toUpperCase() + optionKey.substring(1)}"><Parameter Type="Array">${valuesXml}</Parameter></SetProperty><Method Name="Update" Id="${id + 1}" ObjectPathId="7" />`;

            id += 2;

          } else if (['excludedFileExtensionsForSyncClient'].indexOf(optionKey) > -1) {

            // the XML has to be represented as array of strings
            let valuesXml = '';
            optionValue.split(',').forEach((value: string) => {
              valuesXml += `<Object Type="String">${value}</Object>`;
            })
            propsXml += `<SetProperty Id="${id}" ObjectPathId="7" Name="${optionKey[0].toUpperCase() + optionKey.substring(1)}"><Parameter Type="Array">${valuesXml}</Parameter></SetProperty><Method Name="Update" Id="${id + 1}" ObjectPathId="7" />`;

            id += 2;

          } else {

            propsXml += `<SetProperty Id="${id}" ObjectPathId="7" Name="${optionKey[0].toUpperCase() + optionKey.substring(1)}"><Parameter Type="String">${optionValue}</Parameter></SetProperty>`;

            id++;
          }
        };

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': formDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${propsXml}</Actions><ObjectPaths><Identity Id="7" Name="${json[4]['_ObjectIdentity_'].replace('\n', '&#xA;')}" /></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }

        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--minCompatibilityLevel [minCompatibilityLevel]',
        description: 'Specifies the lower bound on the compatibility level for new sites'
      },
      {
        option: '--maxCompatibilityLevel [maxCompatibilityLevel]',
        description: 'Specifies the upper bound on the compatibility level for new sites'
      },
      {
        option: '--externalServicesEnabled [externalServicesEnabled]',
        description: 'Enables external services for a tenant. External services are defined as services that are not in the Office 365 datacenters. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--noAccessRedirectUrl [noAccessRedirectUrl]',
        description: 'Specifies the URL of the redirected site for those site collections which have the locked state "NoAccess"'
      },
      {
        option: '--sharingCapability [sharingCapability]',
        description: 'Determines what level of sharing is available for the site. The valid values are: ExternalUserAndGuestSharing (default) - External user sharing (share by email) and guest link sharing are both enabled. Disabled - External user sharing (share by email) and guest link sharing are both disabled. ExternalUserSharingOnly - External user sharing (share by email) is enabled, but guest link sharing is disabled. Allowed values Disabled|ExternalUserSharingOnly|ExternalUserAndGuestSharing|ExistingExternalUserSharingOnly',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--displayStartASiteOption [displayStartASiteOption]',
        description: 'Determines whether tenant users see the Start a Site menu option. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--startASiteFormUrl [startASiteFormUrl]',
        description: 'Specifies URL of the form to load in the Start a Site dialog. The valid values are: "" (default) - Blank by default, this will also remove or clear any value that has been set. Full URL - Example:"https://contoso.sharepoint.com/path/to/form"'
      },
      {
        option: '--showEveryoneClaim [showEveryoneClaim]',
        description: 'Enables the administrator to hide the Everyone claim in the People Picker. When users share an item with Everyone, it is accessible to all authenticated users in the tenant\'s Azure Active Directory, including any active external users who have previously accepted invitations. Note, that some SharePoint system resources such as templates and pages are required to be shared to Everyone and this type of sharing does not expose any user data or metadata. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--showAllUsersClaim [showAllUsersClaim]',
        description: 'Enables the administrator to hide the All Users claim groups in People Picker. When users share an item with "All Users (x)", it is accessible to all organization members in the tenant\'s Azure Active Directory who have authenticated with via this method. When users share an item with "All Users (x)" it is accessible to all organtization members in the tenant that used NTLM to authentication with SharePoint. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--showEveryoneExceptExternalUsersClaim [showEveryoneExceptExternalUsersClaim]',
        description: 'Enables the administrator to hide the "Everyone except external users" claim in the People Picker. When users share an item with "Everyone except external users", it is accessible to all organization members in the tenant\'s Azure Active Directory, but not to any users who have previously accepted invitations. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--searchResolveExactEmailOrUPN [searchResolveExactEmailOrUPN]',
        description: 'Removes the search capability from People Picker. Note, recently resolved names will still appear in the list until browser cache is cleared or expired. SharePoint Administrators will still be able to use starts with or partial name matching when enabled. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--officeClientADALDisabled [officeClientADALDisabled]',
        description: 'When set to true this will disable the ability to use Modern Authentication that leverages ADAL across the tenant. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--legacyAuthProtocolsEnabled [legacyAuthProtocolsEnabled]',
        description: 'By default this value is set to true. Setting this parameter prevents Office clients using non-modern authentication protocols from accessing SharePoint Online resources. A value of true - Enables Office clients using non-modern authentication protocols(such as, Forms-Based Authentication (FBA) or Identity Client Runtime Library (IDCRL)) to access SharePoint resources. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--requireAcceptingAccountMatchInvitedAccount [requireAcceptingAccountMatchInvitedAccount]',
        description: 'Ensures that an external user can only accept an external sharing invitation with an account matching the invited email address. Administrators who desire increased control over external collaborators should consider enabling this feature. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--provisionSharedWithEveryoneFolder [provisionSharedWithEveryoneFolder]',
        description: 'Creates a Shared with Everyone folder in every user\'s new OneDrive for Business document library. The valid values are: True (default) - The Shared with Everyone folder is created. False - No folder is created when the site and OneDrive for Business document library is created. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--signInAccelerationDomain [signInAccelerationDomain]',
        description: 'Specifies the home realm discovery value to be sent to Azure Active Directory (AAD) during the user sign-in process. When the organization uses a third-party identity provider, this prevents the user from seeing the Azure Active Directory Home Realm Discovery web page and ensures the user only sees their company\'s Identity Provider\'s portal. This value can also be used with Azure Active Directory Premium to customize the Azure Active Directory login page. Acceleration will not occur on site collections that are shared externally. This value should be configured with the login domain that is used by your company (that is, example@contoso.com). If your company has multiple third-party identity providers, configuring the sign-in acceleration value will break sign-in for your organization. The valid values are: "" (default) - Blank by default, this will also remove or clear any value that has been set. Login Domain - For example: "contoso.com". No value assigned by default'
      },
      {
        option: '--enableGuestSignInAcceleration [enableGuestSignInAcceleration]',
        description: 'Accelerates guest-enabled site collections as well as member-only site collections when the SignInAccelerationDomain parameter is set. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--usePersistentCookiesForExplorerView [usePersistentCookiesForExplorerView]',
        description: 'Lets SharePoint issue a special cookie that will allow this feature to work even when "Keep Me Signed In" is not selected. "Open with Explorer" requires persisted cookies to operate correctly. When the user does not select "Keep Me Signed in" at the time of sign -in, "Open with Explorer" will fail. This special cookie expires after 30 minutes and cannot be cleared by closing the browser or signing out of SharePoint Online.To clear this cookie, the user must log out of their Windows session. The valid values are: False(default) - No special cookie is generated and the normal Office 365 sign -in length / timing applies. True - Generates a special cookie that will allow "Open with Explorer" to function if the "Keep Me Signed In" box is not checked at sign -in. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--bccExternalSharingInvitations [bccExternalSharingInvitations]',
        description: 'When the feature is enabled, all external sharing invitations that are sent will blind copy the e-mail messages listed in the BccExternalSharingsInvitationList. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--bccExternalSharingInvitationsList [bccExternalSharingInvitationsList]',
        description: 'Specifies a list of e-mail addresses to be BCC\'d when the BCC for External Sharing feature is enabled. Multiple addresses can be specified by creating a comma separated list with no spaces'
      },
      {
        option: '--userVoiceForFeedbackEnabled [userVoiceForFeedbackEnabled]',
        description: 'Enables or disables the User Voice Feedback button. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--publicCdnEnabled [publicCdnEnabled]',
        description: 'Enables or disables the publish CDN. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--publicCdnAllowedFileTypes [publicCdnAllowedFileTypes]',
        description: 'Sets public CDN allowed file types'
      },
      {
        option: '--requireAnonymousLinksExpireInDays [requireAnonymousLinksExpireInDays]',
        description: 'Specifies all anonymous links that have been created (or will be created) will expire after the set number of days. To remove the expiration requirement, set the value to zero (0)'
      },
      {
        option: '--sharingAllowedDomainList [sharingAllowedDomainList]',
        description: 'Specifies a list of email domains that is allowed for sharing with the external collaborators. Use the space character as the delimiter for entering multiple values. For example, "contoso.com fabrikam.com"'
      },
      {
        option: '--sharingBlockedDomainList [sharingBlockedDomainList]',
        description: 'Specifies a list of email domains that is blocked or prohibited for sharing with the external collaborators. Use space character as the delimiter for entering multiple values. For example, "contoso.com fabrikam.com"'
      },
      {
        option: '--sharingDomainRestrictionMode [sharingDomainRestrictionMode]',
        description: 'Specifies the external sharing mode for domains. Allowed values None|AllowList|BlockList',
        autocomplete: this.getSharingDomainRestrictionModes()
      },
      {
        option: '--oneDriveStorageQuota [oneDriveStorageQuota]',
        description: 'Sets a default OneDrive for Business storage quota for the tenant. It will be used for new OneDrive for Business sites created. A typical use will be to reduce the amount of storage associated with OneDrive for Business to a level below what the License entitles the users. For example, it could be used to set the quota to 10 gigabytes (GB) by default'
      },
      {
        option: '--oneDriveForGuestsEnabled [oneDriveForGuestsEnabled]',
        description: 'Lets OneDrive for Business creation for administrator managed guest users. Administrator managed Guest users use credentials in the resource tenant to access the resources. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--iPAddressEnforcement [iPAddressEnforcement]',
        description: 'Allows access from network locations that are defined by an administrator. The values are true and false. The default value is false which means the setting is disabled. Before the iPAddressEnforcement parameter is set, make sure you add a valid IPv4 or IPv6 address to the iPAddressAllowList parameter. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--iPAddressAllowList [iPAddressAllowList]',
        description: 'Configures multiple IP addresses or IP address ranges (IPv4 or IPv6). Use commas to separate multiple IP addresses or IP address ranges. Verify there are no overlapping IP addresses and ensure IP ranges use Classless Inter-Domain Routing (CIDR) notation. For example, 172.16.0.0, 192.168.1.0/27. No value is assigned by default'
      },
      {
        option: '--iPAddressWACTokenLifetime [iPAddressWACTokenLifetime]',
        description: 'Sets IP Address WAC token lifetime'
      },
      {
        option: '--useFindPeopleInPeoplePicker [useFindPeopleInPeoplePicker]',
        description: 'Sets use find people in PeoplePicker to true or false. Note: When set to true, users aren\'t able to share with security groups or SharePoint groups. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--defaultSharingLinkType [defaultSharingLinkType]',
        description: 'Lets administrators choose what type of link appears is selected in the “Get a link” sharing dialog box in OneDrive for Business and SharePoint Online. Allowed values None|Direct|Internal|AnonymousAccess',
        autocomplete: this.getSharingLinkType()
      },
      {
        option: '--oDBMembersCanShare [oDBMembersCanShare]',
        description: 'Lets administrators set policy on re-sharing behavior in OneDrive for Business. Allowed values Unspecified|On|Off',
        autocomplete: this.getSharingState()
      },
      {
        option: '--oDBAccessRequests [oDBAccessRequests]',
        description: 'Lets administrators set policy on access requests and requests to share in OneDrive for Business. Allowed values Unspecified|On|Off',
        autocomplete: this.getSharingState()
      },
      {
        option: '--preventExternalUsersFromResharing [preventExternalUsersFromResharing]',
        description: 'Prevents external users from resharing. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--showPeoplePickerSuggestionsForGuestUsers [showPeoplePickerSuggestionsForGuestUsers]',
        description: 'Shows people picker suggestions for guest users. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--fileAnonymousLinkType [fileAnonymousLinkType]',
        description: 'Sets the file anonymous link type to None, View or Edit',
        autocomplete: this.getAnonymousLinkType()
      },
      {
        option: '--folderAnonymousLinkType [folderAnonymousLinkType]',
        description: 'Sets the folder anonymous link type to None, View or Edit',
        autocomplete: this.getAnonymousLinkType()
      },
      {
        option: '--notifyOwnersWhenItemsReshared [notifyOwnersWhenItemsReshared]',
        description: 'When this parameter is set to true and another user re-shares a document from a user\'s OneDrive for Business, the OneDrive for Business owner is notified by email. For additional information about how to configure notifications for external sharing, see Configure notifications for external sharing for OneDrive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--notifyOwnersWhenInvitationsAccepted [notifyOwnersWhenInvitationsAccepted]',
        description: 'When this parameter is set to true and when an external user accepts an invitation to a resource in a user\'s OneDrive for Business, the OneDrive for Business owner is notified by email. For additional information about how to configure notifications for external sharing, see Configure notifications for external sharing for OneDrive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--notificationsInOneDriveForBusinessEnabled [notificationsInOneDriveForBusinessEnabled]',
        description: 'Enables or disables notifications in OneDrive for business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--notificationsInSharePointEnabled [notificationsInSharePointEnabled]',
        description: 'Enables or disables notifications in SharePoint. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ownerAnonymousNotification [ownerAnonymousNotification]',
        description: 'Enables or disables owner anonymous notification. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--commentsOnSitePagesDisabled [commentsOnSitePagesDisabled]',
        description: 'Enables or disables comments on site pages. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--socialBarOnSitePagesDisabled [socialBarOnSitePagesDisabled]',
        description: 'Enables or disables social bar on site pages. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--orphanedPersonalSitesRetentionPeriod [orphanedPersonalSitesRetentionPeriod]',
        description: 'Specifies the number of days after a user\'s Active Directory account is deleted that their OneDrive for Business content will be deleted. The value range is in days, between 30 and 3650. The default value is 30'
      },
      {
        option: '--disallowInfectedFileDownload [disallowInfectedFileDownload]',
        description: 'Prevents the Download button from being displayed on the Virus Found warning page. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--defaultLinkPermission [defaultLinkPermission]',
        description: 'Choose the dafault permission that is selected when users share. This applies to anonymous access, internal and direct links. Allowed values None|View|Edit',
        autocomplete: this.getSharingPermissionType()
      },
      {
        option: '--conditionalAccessPolicy [conditionalAccessPolicy]',
        description: 'Configures conditional access policy. Allowed values AllowFullAccess|AllowLimitedAccess|BlockAccess',
        autocomplete: this.getSPOConditionalAccessPolicyType()
      },
      {
        option: '--allowDownloadingNonWebViewableFiles [allowDownloadingNonWebViewableFiles]',
        description: 'Allows downloading non web viewable files. The Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowEditing [allowEditing]',
        description: 'Allows editing. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--applyAppEnforcedRestrictionsToAdHocRecipients [applyAppEnforcedRestrictionsToAdHocRecipients]',
        description: 'Applies app enforced restrictions to AdHoc recipients. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--filePickerExternalImageSearchEnabled [filePickerExternalImageSearchEnabled]',
        description: 'Enables file picker external image search. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--emailAttestationRequired [emailAttestationRequired]',
        description: 'Sets email attestation to required. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--emailAttestationReAuthDays [emailAttestationReAuthDays]',
        description: 'Sets email attestation re-auth days'
      },
      {
        option: '--hideDefaultThemes [hideDefaultThemes]',
        description: 'Defines if the default themes are visible or hidden. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      // not included in the PnP PowerShell, most of them are new and maybe the cmdlet is not updated recently.
      {
        option: '--blockAccessOnUnmanagedDevices [blockAccessOnUnmanagedDevices]',
        description: 'Blocks access on unmanaged devices. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowLimitedAccessOnUnmanagedDevices [allowLimitedAccessOnUnmanagedDevices]',
        description: 'Allows limited access on unmanaged devices blocks. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--blockDownloadOfAllFilesForGuests [blockDownloadOfAllFilesForGuests]',
        description: 'Blocks download of all files for guests. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--blockDownloadOfAllFilesOnUnmanagedDevices [blockDownloadOfAllFilesOnUnmanagedDevices]',
        description: 'Blocks download of all files on unmanaged devices. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--blockDownloadOfViewableFilesForGuests [blockDownloadOfViewableFilesForGuests]',
        description: 'Blocks download of viewable files for guests. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--blockDownloadOfViewableFilesOnUnmanagedDevices [blockDownloadOfViewableFilesOnUnmanagedDevices]',
        description: 'Blocks download of viewable files on unmanaged devices. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--blockMacSync [blockMacSync]',
        description: 'Blocks Mac sync. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableReportProblemDialog [disableReportProblemDialog]',
        description: 'Disables report problem dialog. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--displayNamesOfFileViewers [displayNamesOfFileViewers]',
        description: 'Displayes names of file viewers. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableMinimumVersionRequirement [enableMinimumVersionRequirement]',
        description: 'Enables minimum version requirement. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hideSyncButtonOnODB [hideSyncButtonOnODB]',
        description: 'Hides the sync button on One Drive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--isUnmanagedSyncClientForTenantRestricted [isUnmanagedSyncClientForTenantRestricted]',
        description: 'Is unmanaged sync client for tenant restricted. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--limitedAccessFileType [limitedAccessFileType]',
        description: 'Allows users to preview only Office files in the browser. This option increases security but may be a barrier to user productivity. Allowed values OfficeOnlineFilesOnly|WebPreviewableFiles|OtherFiles',
        autocomplete: this.getSPOLimitedAccessFileType()
      },
      {
        option: '--optOutOfGrooveBlock [optOutOfGrooveBlock]',
        description: 'Opts out of the groove block. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--optOutOfGrooveSoftBlock [optOutOfGrooveSoftBlock]',
        description: 'Opts out of Groove soft block. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--orgNewsSiteUrl [orgNewsSiteUrl]',
        description: 'Organization news site url'
      },
      {
        option: '--permissiveBrowserFileHandlingOverride [permissiveBrowserFileHandlingOverride]',
        description: 'Permissive browser fileHandling override. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--showNGSCDialogForSyncOnODB [showNGSCDialogForSyncOnODB]',
        description: 'Show NGSC dialog for sync on OneDrive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--specialCharactersStateInFileFolderNames [specialCharactersStateInFileFolderNames]',
        description: 'Sets the special characters state in file and folder names in SharePoint and OneDrive for Business. Allowed values NoPreference|Allowed|Disallowed',
        autocomplete: this.getSpecialCharactersState()
      },
      {
        option: '--syncPrivacyProfileProperties [syncPrivacyProfileProperties]',
        description: 'Syncs privacy profile properties. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--excludedFileExtensionsForSyncClient [excludedFileExtensionsForSyncClient]',
        description: 'Excluded file extensions for sync client. Array of strings split by comma (\',\')'
      },
      {
        option: '--allowedDomainListForSyncClient [allowedDomainListForSyncClient]',
        description: 'Sets allowed domain list for sync client. Array of GUIDs split by comma (\',\'). Example:c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201'
      },
      {
        option: '--disabledWebPartIds [disabledWebPartIds]',
        description: 'Sets disabled web part Ids. Array of GUIDs split by comma (\',\'). Example:c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      const opts: any = args.options;
      let hasAtLeastOneOption: boolean = false;

      for (let propertyKey of Object.keys(opts)) {

        let propertyValue = opts[propertyKey];

        if (this.isExcludedOption(propertyKey)) {
          continue;
        }
        hasAtLeastOneOption = true;
        const commandOptions: CommandOption[] = this.options();

        for (let item of commandOptions) {
          if (item.option.indexOf(propertyKey) > -1) {

            if (item.autocomplete) {
              if (item.autocomplete.indexOf(propertyValue.toString()) === -1) {
                return `${propertyKey} option has invalid value of ${propertyValue}. Allowed values are ${JSON.stringify(item.autocomplete)}`;
              }
            }
          }
        }
      }

      if (!hasAtLeastOneOption) {
        return `You must specify at least one option`;
      }
      return true;
    };
  }

  public isExcludedOption(optionKey: string): boolean {

    // it is not possible to dynamically get the GlobalOptions
    // prop keys since they are nullable
    // so we have to maintain that array bellow once new global option
    // is added to the GlobalOptions interface
    return ['output', 'debug', 'verbose'].indexOf(optionKey) > -1;
  }

  public mapEnumToInt(key: string, value: string): number {

    switch (key) {
      case 'sharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'sharingDomainRestrictionMode':
        return this.getSharingDomainRestrictionModes().indexOf(value);
      case 'defaultSharingLinkType':
        return this.getSharingLinkType().indexOf(value);
      case 'oDBMembersCanShare':
        return this.getSharingState().indexOf(value);
      case 'oDBAccessRequests':
        return this.getSharingState().indexOf(value);
      case 'fileAnonymousLinkType':
        return this.getAnonymousLinkType().indexOf(value);
      case 'folderAnonymousLinkType':
        return this.getAnonymousLinkType().indexOf(value);
      case 'defaultLinkPermission':
        return this.getSharingPermissionType().indexOf(value);
      case 'conditionalAccessPolicy':
        return this.getSPOConditionalAccessPolicyType().indexOf(value);
      case 'limitedAccessFileType':
        return this.getSPOLimitedAccessFileType().indexOf(value);
      case 'specialCharactersStateInFileFolderNames':
        return this.getSpecialCharactersState().indexOf(value);
      default:
        return -1;
    }
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online
    tenant admin site, using the ${chalk.blue(commands.CONNECT)} command.

  Examples:
  
    Sets single tenant global setting
      ${chalk.grey(config.delimiter)} ${commands.TENANT_SETTINGS_SET} --userVoiceForFeedbackEnabled true

    Sets multiple tenant global settings at once
      ${chalk.grey(config.delimiter)} ${commands.TENANT_SETTINGS_SET} --userVoiceForFeedbackEnabled true --hideSyncButtonOnODB true --disabledWebPartIds c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201

  More information:

    PnP PowerShell Set-PnPTenant
      https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnptenant?view=sharepoint-ps

    SharePoint Online Set-SPOTenant:
      https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenant?view=sharepoint-ps

    SharePoint Online Set-SPOTenantCdnEnabled:
      https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenantcdnenabled?view=sharepoint-ps

    SharePoint Online Set-SPOTenantSyncClientRestriction
      https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenantsyncclientrestriction?view=sharepoint-ps
  ` );
  }
}

module.exports = new SpoTenantSettingsSetCommand();