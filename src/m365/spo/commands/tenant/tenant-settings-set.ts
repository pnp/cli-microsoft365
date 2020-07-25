import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import request from '../../../../request';
import config from '../../../../config';
import commands from '../../commands';
import Utils from '../../../../Utils';
import {
  CommandOption,
  CommandError,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandInstance } from '../../../../cli';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  MinCompatibilityLevel: number;
  MaxCompatibilityLevel: number;
  ExternalServicesEnabled: boolean;
  NoAccessRedirectUrl: string;
  SharingCapability: string; // <SharingCapabilities>
  DisplayStartASiteOption: boolean;
  StartASiteFormUrl: string;
  ShowEveryoneClaim: boolean;
  ShowAllUsersClaim: boolean;
  ShowEveryoneExceptExternalUsersClaim: boolean;
  SearchResolveExactEmailOrUPN: boolean;
  OfficeClientADALDisabled: boolean;
  LegacyAuthProtocolsEnabled: boolean;
  RequireAcceptingAccountMatchInvitedAccount: boolean;
  ProvisionSharedWithEveryoneFolder: boolean;
  SignInAccelerationDomain: string;
  EnableGuestSignInAcceleration: boolean;
  UsePersistentCookiesForExplorerView: boolean;
  BccExternalSharingInvitations: boolean;
  BccExternalSharingInvitationsList: string;
  UserVoiceForFeedbackEnabled: boolean;
  PublicCdnEnabled: boolean;
  PublicCdnAllowedFileTypes: string;
  RequireAnonymousLinksExpireInDays: number;
  SharingAllowedDomainList: string;
  SharingBlockedDomainList: string;
  SharingDomainRestrictionMode: string; // <SharingDomainRestrictionModes>
  OneDriveStorageQuota: number;
  OneDriveForGuestsEnabled: boolean;
  IPAddressEnforcement: boolean;
  IPAddressAllowList: string;
  IPAddressWACTokenLifetime: number;
  UseFindPeopleInPeoplePicker: boolean;
  DefaultSharingLinkType: string; // <SharingLinkType>
  ODBMembersCanShare: string; // <SharingState>
  ODBAccessRequests: string; // <SharingState>
  PreventExternalUsersFromResharing: boolean;
  ShowPeoplePickerSuggestionsForGuestUsers: boolean;
  FileAnonymousLinkType: string; // <AnonymousLinkType>
  FolderAnonymousLinkType: string; // <AnonymousLinkType>
  NotifyOwnersWhenItemsReshared: boolean;
  NotifyOwnersWhenInvitationsAccepted: boolean;
  NotificationsInOneDriveForBusinessEnabled: boolean;
  NotificationsInSharePointEnabled: boolean;
  OwnerAnonymousNotification: boolean;
  CommentsOnSitePagesDisabled: boolean;
  SocialBarOnSitePagesDisabled: boolean;
  OrphanedPersonalSitesRetentionPeriod: number;
  DisallowInfectedFileDownload: boolean;
  DefaultLinkPermission: string; // <SharingPermissionType>
  ConditionalAccessPolicy: string; // <SPOConditionalAccessPolicyType>
  AllowDownloadingNonWebViewableFiles: boolean;
  AllowEditing: boolean;
  ApplyAppEnforcedRestrictionsToAdHocRecipients: boolean;
  FilePickerExternalImageSearchEnabled: boolean;
  EmailAttestationRequired: boolean;
  EmailAttestationReAuthDays: number;
  HideDefaultThemes: boolean;
  // not included in the PnP PowerShell, most of them are new and maybe the cmdlet is not updated recently.
  BlockAccessOnUnmanagedDevices: boolean;
  AllowLimitedAccessOnUnmanagedDevices: boolean;
  BlockDownloadOfAllFilesForGuests: boolean;
  BlockDownloadOfAllFilesOnUnmanagedDevices: boolean;
  BlockDownloadOfViewableFilesForGuests: boolean;
  BlockDownloadOfViewableFilesOnUnmanagedDevices: boolean;
  BlockMacSync: boolean;
  DisableReportProblemDialog: boolean;
  DisplayNamesOfFileViewers: boolean;
  EnableMinimumVersionRequirement: boolean;
  HideSyncButtonOnODB: boolean;
  IsUnmanagedSyncClientForTenantRestricted: boolean;
  LimitedAccessFileType: string; // <LimitedAccessFileType>
  OptOutOfGrooveBlock: boolean;
  OptOutOfGrooveSoftBlock: boolean;
  OrgNewsSiteUrl: string;
  PermissiveBrowserFileHandlingOverride: boolean;
  ShowNGSCDialogForSyncOnODB: boolean;
  SpecialCharactersStateInFileFolderNames: string; // <SpecialCharactersState>
  SyncPrivacyProfileProperties: boolean;
  ExcludedFileExtensionsForSyncClient: string[];
  AllowedDomainListForSyncClient: string[];
  DisabledWebPartIds: string[];
}

class SpoTenantSettingsSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets tenant global settings';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.MinCompatibilityLevel = (!(!args.options.MinCompatibilityLevel)).toString();
    telemetryProps.MaxCompatibilityLevel = (!(!args.options.MaxCompatibilityLevel)).toString();
    telemetryProps.ExternalServicesEnabled = (!(!args.options.ExternalServicesEnabled)).toString();
    telemetryProps.NoAccessRedirectUrl = (!(!args.options.NoAccessRedirectUrl)).toString();
    telemetryProps.SharingCapability = (!(!args.options.SharingCapability)).toString();
    telemetryProps.DisplayStartASiteOption = (!(!args.options.DisplayStartASiteOption)).toString();
    telemetryProps.StartASiteFormUrl = (!(!args.options.StartASiteFormUrl)).toString();
    telemetryProps.ShowEveryoneClaim = (!(!args.options.ShowEveryoneClaim)).toString();
    telemetryProps.ShowAllUsersClaim = (!(!args.options.ShowAllUsersClaim)).toString();
    telemetryProps.ShowEveryoneExceptExternalUsersClaim = (!(!args.options.ShowEveryoneExceptExternalUsersClaim)).toString();
    telemetryProps.SearchResolveExactEmailOrUPN = (!(!args.options.SearchResolveExactEmailOrUPN)).toString();
    telemetryProps.OfficeClientADALDisabled = (!(!args.options.OfficeClientADALDisabled)).toString();
    telemetryProps.LegacyAuthProtocolsEnabled = (!(!args.options.LegacyAuthProtocolsEnabled)).toString();
    telemetryProps.RequireAcceptingAccountMatchInvitedAccount = (!(!args.options.RequireAcceptingAccountMatchInvitedAccount)).toString();
    telemetryProps.ProvisionSharedWithEveryoneFolder = (!(!args.options.ProvisionSharedWithEveryoneFolder)).toString();
    telemetryProps.SignInAccelerationDomain = (!(!args.options.SignInAccelerationDomain)).toString();
    telemetryProps.EnableGuestSignInAcceleration = (!(!args.options.EnableGuestSignInAcceleration)).toString();
    telemetryProps.UsePersistentCookiesForExplorerView = (!(!args.options.UsePersistentCookiesForExplorerView)).toString();
    telemetryProps.BccExternalSharingInvitations = (!(!args.options.BccExternalSharingInvitations)).toString();
    telemetryProps.BccExternalSharingInvitationsList = (!(!args.options.BccExternalSharingInvitationsList)).toString();
    telemetryProps.UserVoiceForFeedbackEnabled = (!(!args.options.UserVoiceForFeedbackEnabled)).toString();
    telemetryProps.PublicCdnEnabled = (!(!args.options.PublicCdnEnabled)).toString();
    telemetryProps.PublicCdnAllowedFileTypes = (!(!args.options.PublicCdnAllowedFileTypes)).toString();
    telemetryProps.RequireAnonymousLinksExpireInDays = (!(!args.options.RequireAnonymousLinksExpireInDays)).toString();
    telemetryProps.SharingAllowedDomainList = (!(!args.options.SharingAllowedDomainList)).toString();
    telemetryProps.SharingBlockedDomainList = (!(!args.options.SharingBlockedDomainList)).toString();
    telemetryProps.SharingDomainRestrictionMode = (!(!args.options.SharingDomainRestrictionMode)).toString();
    telemetryProps.OneDriveStorageQuota = (!(!args.options.OneDriveStorageQuota)).toString();
    telemetryProps.OneDriveForGuestsEnabled = (!(!args.options.OneDriveForGuestsEnabled)).toString();
    telemetryProps.IPAddressEnforcement = (!(!args.options.IPAddressEnforcement)).toString();
    telemetryProps.IPAddressAllowList = (!(!args.options.IPAddressAllowList)).toString();
    telemetryProps.IPAddressWACTokenLifetime = (!(!args.options.IPAddressWACTokenLifetime)).toString();
    telemetryProps.UseFindPeopleInPeoplePicker = (!(!args.options.UseFindPeopleInPeoplePicker)).toString();
    telemetryProps.DefaultSharingLinkType = (!(!args.options.DefaultSharingLinkType)).toString();
    telemetryProps.ODBMembersCanShare = (!(!args.options.ODBMembersCanShare)).toString();
    telemetryProps.ODBAccessRequests = (!(!args.options.ODBAccessRequests)).toString();
    telemetryProps.PreventExternalUsersFromResharing = (!(!args.options.PreventExternalUsersFromResharing)).toString();
    telemetryProps.ShowPeoplePickerSuggestionsForGuestUsers = (!(!args.options.ShowPeoplePickerSuggestionsForGuestUsers)).toString();
    telemetryProps.FileAnonymousLinkType = (!(!args.options.FileAnonymousLinkType)).toString();
    telemetryProps.FolderAnonymousLinkType = (!(!args.options.FolderAnonymousLinkType)).toString();
    telemetryProps.NotifyOwnersWhenItemsReshared = (!(!args.options.NotifyOwnersWhenItemsReshared)).toString();
    telemetryProps.NotifyOwnersWhenInvitationsAccepted = (!(!args.options.NotifyOwnersWhenInvitationsAccepted)).toString();
    telemetryProps.NotificationsInOneDriveForBusinessEnabled = (!(!args.options.NotificationsInOneDriveForBusinessEnabled)).toString();
    telemetryProps.NotificationsInSharePointEnabled = (!(!args.options.NotificationsInSharePointEnabled)).toString();
    telemetryProps.OwnerAnonymousNotification = (!(!args.options.OwnerAnonymousNotification)).toString();
    telemetryProps.CommentsOnSitePagesDisabled = (!(!args.options.CommentsOnSitePagesDisabled)).toString();
    telemetryProps.SocialBarOnSitePagesDisabled = (!(!args.options.SocialBarOnSitePagesDisabled)).toString();
    telemetryProps.OrphanedPersonalSitesRetentionPeriod = (!(!args.options.OrphanedPersonalSitesRetentionPeriod)).toString();
    telemetryProps.DisallowInfectedFileDownload = (!(!args.options.DisallowInfectedFileDownload)).toString();
    telemetryProps.DefaultLinkPermission = (!(!args.options.DefaultLinkPermission)).toString();
    telemetryProps.ConditionalAccessPolicy = (!(!args.options.ConditionalAccessPolicy)).toString();
    telemetryProps.AllowDownloadingNonWebViewableFiles = (!(!args.options.AllowDownloadingNonWebViewableFiles)).toString();
    telemetryProps.AllowEditing = (!(!args.options.AllowEditing)).toString();
    telemetryProps.ApplyAppEnforcedRestrictionsToAdHocRecipients = (!(!args.options.ApplyAppEnforcedRestrictionsToAdHocRecipients)).toString();
    telemetryProps.FilePickerExternalImageSearchEnabled = (!(!args.options.FilePickerExternalImageSearchEnabled)).toString();
    telemetryProps.EmailAttestationRequired = (!(!args.options.EmailAttestationRequired)).toString();
    telemetryProps.EmailAttestationReAuthDays = (!(!args.options.EmailAttestationReAuthDays)).toString();
    telemetryProps.HideDefaultThemes = (!(!args.options.HideDefaultThemes)).toString();
    telemetryProps.BlockAccessOnUnmanagedDevices = (!(!args.options.BlockAccessOnUnmanagedDevices)).toString();
    telemetryProps.AllowLimitedAccessOnUnmanagedDevices = (!(!args.options.AllowLimitedAccessOnUnmanagedDevices)).toString();
    telemetryProps.BlockDownloadOfAllFilesForGuests = (!(!args.options.BlockDownloadOfAllFilesForGuests)).toString();
    telemetryProps.BlockDownloadOfAllFilesOnUnmanagedDevices = (!(!args.options.BlockDownloadOfAllFilesOnUnmanagedDevices)).toString();
    telemetryProps.BlockDownloadOfViewableFilesForGuests = (!(!args.options.BlockDownloadOfViewableFilesForGuests)).toString();
    telemetryProps.BlockDownloadOfViewableFilesOnUnmanagedDevices = (!(!args.options.BlockDownloadOfViewableFilesOnUnmanagedDevices)).toString();
    telemetryProps.BlockMacSync = (!(!args.options.BlockMacSync)).toString();
    telemetryProps.DisableReportProblemDialog = (!(!args.options.DisableReportProblemDialog)).toString();
    telemetryProps.DisplayNamesOfFileViewers = (!(!args.options.DisplayNamesOfFileViewers)).toString();
    telemetryProps.EnableMinimumVersionRequirement = (!(!args.options.EnableMinimumVersionRequirement)).toString();
    telemetryProps.HideSyncButtonOnODB = (!(!args.options.HideSyncButtonOnODB)).toString();
    telemetryProps.IsUnmanagedSyncClientForTenantRestricted = (!(!args.options.IsUnmanagedSyncClientForTenantRestricted)).toString();
    telemetryProps.LimitedAccessFileType = (!(!args.options.LimitedAccessFileType)).toString();
    telemetryProps.OptOutOfGrooveBlock = (!(!args.options.OptOutOfGrooveBlock)).toString();
    telemetryProps.OptOutOfGrooveSoftBlock = (!(!args.options.OptOutOfGrooveSoftBlock)).toString();
    telemetryProps.OrgNewsSiteUrl = (!(!args.options.OrgNewsSiteUrl)).toString();
    telemetryProps.PermissiveBrowserFileHandlingOverride = (!(!args.options.PermissiveBrowserFileHandlingOverride)).toString();
    telemetryProps.ShowNGSCDialogForSyncOnODB = (!(!args.options.ShowNGSCDialogForSyncOnODB)).toString();
    telemetryProps.SpecialCharactersStateInFileFolderNames = (!(!args.options.SpecialCharactersStateInFileFolderNames)).toString();
    telemetryProps.SyncPrivacyProfileProperties = (!(!args.options.SyncPrivacyProfileProperties)).toString();
    telemetryProps.ExcludedFileExtensionsForSyncClient = (!(!args.options.ExcludedFileExtensionsForSyncClient)).toString();
    telemetryProps.DisabledWebPartIds = (!(!args.options.DisabledWebPartIds)).toString();
    telemetryProps.AllowedDomainListForSyncClient = (!(!args.options.AllowedDomainListForSyncClient)).toString();
    return telemetryProps;
  }

  public getAllEnumOptions(): string[] {
    return ['SharingCapability', 'SharingDomainRestrictionMode', 'DefaultSharingLinkType', 'ODBMembersCanShare', 'ODBAccessRequests', 'FileAnonymousLinkType', 'FolderAnonymousLinkType', 'DefaultLinkPermission', 'ConditionalAccessPolicy', 'LimitedAccessFileType', 'SpecialCharactersStateInFileFolderNames'];
  }

  // all enums as get methods
  private getSharingLinkType(): string[] { return ['None', 'Direct', 'Internal', 'AnonymousAccess']; }
  private getSharingCapabilities(): string[] { return ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']; }
  private getSharingDomainRestrictionModes(): string[] { return ['None', 'AllowList', 'BlockList'] };
  private getSharingState(): string[] { return ['Unspecified', 'On', 'Off']; }
  private getAnonymousLinkType(): string[] { return ['None', 'View', 'Edit']; }
  private getSharingPermissionType(): string[] { return ['None', 'View', 'Edit']; }
  private getSPOConditionalAccessPolicyType(): string[] { return ['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess']; }
  private getSpecialCharactersState(): string[] { return ['NoPreference', 'Allowed', 'Disallowed']; }
  private getSPOLimitedAccessFileType(): string[] { return ['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles']; }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    let formDigestValue = '';
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    this
      .getTenantId(cmd, this.debug)
      .then((_tenantId: string): Promise<string> => {
        tenantId = _tenantId;
        return this.getSpoAdminUrl(cmd, this.debug);
      })
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        formDigestValue = res.FormDigestValue;

        // map the args.options to XML Properties
        let propsXml: string = '';
        let id: number = 42; // geek's humor
        for (let optionKey of Object.keys(args.options)) {
          if (this.isExcludedOption(optionKey)) {
            continue;
          }

          let optionValue = args.options[optionKey];
          if (this.getAllEnumOptions().indexOf(optionKey) > -1) {
            // map enum values to int
            optionValue = this.mapEnumToInt(optionKey, args.options[optionKey]);
          }

          if (['AllowedDomainListForSyncClient', 'DisabledWebPartIds'].indexOf(optionKey) > -1) {
            // the XML has to be represented as array of guids
            let valuesXml: string = '';
            optionValue.split(',').forEach((value: string) => {
              valuesXml += `<Object Type="Guid">{${Utils.escapeXml(value)}}</Object>`;
            });
            propsXml += `<SetProperty Id="${id++}" ObjectPathId="7" Name="${optionKey}"><Parameter Type="Array">${valuesXml}</Parameter></SetProperty><Method Name="Update" Id="${id++}" ObjectPathId="7" />`;
          }
          else if (['ExcludedFileExtensionsForSyncClient'].indexOf(optionKey) > -1) {
            // the XML has to be represented as array of strings
            let valuesXml: string = '';
            optionValue.split(',').forEach((value: string) => {
              valuesXml += `<Object Type="String">${value}</Object>`;
            });
            propsXml += `<SetProperty Id="${id++}" ObjectPathId="7" Name="${optionKey}"><Parameter Type="Array">${valuesXml}</Parameter></SetProperty><Method Name="Update" Id="${id++}" ObjectPathId="7" />`;
          }
          else {
            propsXml += `<SetProperty Id="${id++}" ObjectPathId="7" Name="${optionKey}"><Parameter Type="String">${optionValue}</Parameter></SetProperty>`;
          }
        };

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${propsXml}</Actions><ObjectPaths><Identity Id="7" Name="${tenantId}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
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
        option: '--MinCompatibilityLevel [MinCompatibilityLevel]',
        description: 'Specifies the lower bound on the compatibility level for new sites'
      },
      {
        option: '--MaxCompatibilityLevel [MaxCompatibilityLevel]',
        description: 'Specifies the upper bound on the compatibility level for new sites'
      },
      {
        option: '--ExternalServicesEnabled [ExternalServicesEnabled]',
        description: 'Enables external services for a tenant. External services are defined as services that are not in the Microsoft 365 datacenters. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NoAccessRedirectUrl [NoAccessRedirectUrl]',
        description: 'Specifies the URL of the redirected site for those site collections which have the locked state "NoAccess"'
      },
      {
        option: '--SharingCapability [SharingCapability]',
        description: 'Determines what level of sharing is available for the site. The valid values are: ExternalUserAndGuestSharing (default) - External user sharing (share by email) and guest link sharing are both enabled. Disabled - External user sharing (share by email) and guest link sharing are both disabled. ExternalUserSharingOnly - External user sharing (share by email) is enabled, but guest link sharing is disabled. Allowed values Disabled|ExternalUserSharingOnly|ExternalUserAndGuestSharing|ExistingExternalUserSharingOnly',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--DisplayStartASiteOption [DisplayStartASiteOption]',
        description: 'Determines whether tenant users see the Start a Site menu option. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--StartASiteFormUrl [StartASiteFormUrl]',
        description: 'Specifies URL of the form to load in the Start a Site dialog. The valid values are: "" (default) - Blank by default, this will also remove or clear any value that has been set. Full URL - Example:"https://contoso.sharepoint.com/path/to/form"'
      },
      {
        option: '--ShowEveryoneClaim [ShowEveryoneClaim]',
        description: 'Enables the administrator to hide the Everyone claim in the People Picker. When users share an item with Everyone, it is accessible to all authenticated users in the tenant\'s Azure Active Directory, including any active external users who have previously accepted invitations. Note, that some SharePoint system resources such as templates and pages are required to be shared to Everyone and this type of sharing does not expose any user data or metadata. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowAllUsersClaim [ShowAllUsersClaim]',
        description: 'Enables the administrator to hide the All Users claim groups in People Picker. When users share an item with "All Users (x)", it is accessible to all organization members in the tenant\'s Azure Active Directory who have authenticated with via this method. When users share an item with "All Users (x)" it is accessible to all organtization members in the tenant that used NTLM to authentication with SharePoint. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowEveryoneExceptExternalUsersClaim [ShowEveryoneExceptExternalUsersClaim]',
        description: 'Enables the administrator to hide the "Everyone except external users" claim in the People Picker. When users share an item with "Everyone except external users", it is accessible to all organization members in the tenant\'s Azure Active Directory, but not to any users who have previously accepted invitations. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SearchResolveExactEmailOrUPN [SearchResolveExactEmailOrUPN]',
        description: 'Removes the search capability from People Picker. Note, recently resolved names will still appear in the list until browser cache is cleared or expired. SharePoint Administrators will still be able to use starts with or partial name matching when enabled. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OfficeClientADALDisabled [OfficeClientADALDisabled]',
        description: 'When set to true this will disable the ability to use Modern Authentication that leverages ADAL across the tenant. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--LegacyAuthProtocolsEnabled [LegacyAuthProtocolsEnabled]',
        description: 'By default this value is set to true. Setting this parameter prevents Office clients using non-modern authentication protocols from accessing SharePoint Online resources. A value of true - Enables Office clients using non-modern authentication protocols(such as, Forms-Based Authentication (FBA) or Identity Client Runtime Library (IDCRL)) to access SharePoint resources. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--RequireAcceptingAccountMatchInvitedAccount [RequireAcceptingAccountMatchInvitedAccount]',
        description: 'Ensures that an external user can only accept an external sharing invitation with an account matching the invited email address. Administrators who desire increased control over external collaborators should consider enabling this feature. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ProvisionSharedWithEveryoneFolder [ProvisionSharedWithEveryoneFolder]',
        description: 'Creates a Shared with Everyone folder in every user\'s new OneDrive for Business document library. The valid values are: True (default) - The Shared with Everyone folder is created. False - No folder is created when the site and OneDrive for Business document library is created. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SignInAccelerationDomain [SignInAccelerationDomain]',
        description: 'Specifies the home realm discovery value to be sent to Azure Active Directory (AAD) during the user sign-in process. When the organization uses a third-party identity provider, this prevents the user from seeing the Azure Active Directory Home Realm Discovery web page and ensures the user only sees their company\'s Identity Provider\'s portal. This value can also be used with Azure Active Directory Premium to customize the Azure Active Directory login page. Acceleration will not occur on site collections that are shared externally. This value should be configured with the login domain that is used by your company (that is, example@contoso.com). If your company has multiple third-party identity providers, configuring the sign-in acceleration value will break sign-in for your organization. The valid values are: "" (default) - Blank by default, this will also remove or clear any value that has been set. Login Domain - For example: "contoso.com". No value assigned by default'
      },
      {
        option: '--EnableGuestSignInAcceleration [EnableGuestSignInAcceleration]',
        description: 'Accelerates guest-enabled site collections as well as member-only site collections when the SignInAccelerationDomain parameter is set. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--UsePersistentCookiesForExplorerView [UsePersistentCookiesForExplorerView]',
        description: 'Lets SharePoint issue a special cookie that will allow this feature to work even when "Keep Me Signed In" is not selected. "Open with Explorer" requires persisted cookies to operate correctly. When the user does not select "Keep Me Signed in" at the time of sign -in, "Open with Explorer" will fail. This special cookie expires after 30 minutes and cannot be cleared by closing the browser or signing out of SharePoint Online.To clear this cookie, the user must log out of their Windows session. The valid values are: False(default) - No special cookie is generated and the normal Microsoft 365 sign -in length / timing applies. True - Generates a special cookie that will allow "Open with Explorer" to function if the "Keep Me Signed In" box is not checked at sign -in. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BccExternalSharingInvitations [BccExternalSharingInvitations]',
        description: 'When the feature is enabled, all external sharing invitations that are sent will blind copy the e-mail messages listed in the BccExternalSharingsInvitationList. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BccExternalSharingInvitationsList [BccExternalSharingInvitationsList]',
        description: 'Specifies a list of e-mail addresses to be BCC\'d when the BCC for External Sharing feature is enabled. Multiple addresses can be specified by creating a comma separated list with no spaces'
      },
      {
        option: '--UserVoiceForFeedbackEnabled [UserVoiceForFeedbackEnabled]',
        description: 'Enables or disables the User Voice Feedback button. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--PublicCdnEnabled [PublicCdnEnabled]',
        description: 'Enables or disables the publish CDN. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--PublicCdnAllowedFileTypes [PublicCdnAllowedFileTypes]',
        description: 'Sets public CDN allowed file types'
      },
      {
        option: '--RequireAnonymousLinksExpireInDays [RequireAnonymousLinksExpireInDays]',
        description: 'Specifies all anonymous links that have been created (or will be created) will expire after the set number of days. To remove the expiration requirement, set the value to zero (0)'
      },
      {
        option: '--SharingAllowedDomainList [SharingAllowedDomainList]',
        description: 'Specifies a list of email domains that is allowed for sharing with the external collaborators. Use the space character as the delimiter for entering multiple values. For example, "contoso.com fabrikam.com"'
      },
      {
        option: '--SharingBlockedDomainList [SharingBlockedDomainList]',
        description: 'Specifies a list of email domains that is blocked or prohibited for sharing with the external collaborators. Use space character as the delimiter for entering multiple values. For example, "contoso.com fabrikam.com"'
      },
      {
        option: '--SharingDomainRestrictionMode [SharingDomainRestrictionMode]',
        description: 'Specifies the external sharing mode for domains. Allowed values None|AllowList|BlockList',
        autocomplete: this.getSharingDomainRestrictionModes()
      },
      {
        option: '--OneDriveStorageQuota [OneDriveStorageQuota]',
        description: 'Sets a default OneDrive for Business storage quota for the tenant. It will be used for new OneDrive for Business sites created. A typical use will be to reduce the amount of storage associated with OneDrive for Business to a level below what the License entitles the users. For example, it could be used to set the quota to 10 gigabytes (GB) by default'
      },
      {
        option: '--OneDriveForGuestsEnabled [OneDriveForGuestsEnabled]',
        description: 'Lets OneDrive for Business creation for administrator managed guest users. Administrator managed Guest users use credentials in the resource tenant to access the resources. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IPAddressEnforcement [IPAddressEnforcement]',
        description: 'Allows access from network locations that are defined by an administrator. The values are true and false. The default value is false which means the setting is disabled. Before the IPAddressEnforcement parameter is set, make sure you add a valid IPv4 or IPv6 address to the IPAddressAllowList parameter. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IPAddressAllowList [IPAddressAllowList]',
        description: 'Configures multiple IP addresses or IP address ranges (IPv4 or IPv6). Use commas to separate multiple IP addresses or IP address ranges. Verify there are no overlapping IP addresses and ensure IP ranges use Classless Inter-Domain Routing (CIDR) notation. For example, 172.16.0.0, 192.168.1.0/27. No value is assigned by default'
      },
      {
        option: '--IPAddressWACTokenLifetime [IPAddressWACTokenLifetime]',
        description: 'Sets IP Address WAC token lifetime'
      },
      {
        option: '--UseFindPeopleInPeoplePicker [UseFindPeopleInPeoplePicker]',
        description: 'Sets use find people in PeoplePicker to true or false. Note: When set to true, users aren\'t able to share with security groups or SharePoint groups. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DefaultSharingLinkType [DefaultSharingLinkType]',
        description: 'Lets administrators choose what type of link appears is selected in the “Get a link” sharing dialog box in OneDrive for Business and SharePoint Online. Allowed values None|Direct|Internal|AnonymousAccess',
        autocomplete: this.getSharingLinkType()
      },
      {
        option: '--ODBMembersCanShare [ODBMembersCanShare]',
        description: 'Lets administrators set policy on re-sharing behavior in OneDrive for Business. Allowed values Unspecified|On|Off',
        autocomplete: this.getSharingState()
      },
      {
        option: '--ODBAccessRequests [ODBAccessRequests]',
        description: 'Lets administrators set policy on access requests and requests to share in OneDrive for Business. Allowed values Unspecified|On|Off',
        autocomplete: this.getSharingState()
      },
      {
        option: '--PreventExternalUsersFromResharing [PreventExternalUsersFromResharing]',
        description: 'Prevents external users from resharing. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowPeoplePickerSuggestionsForGuestUsers [ShowPeoplePickerSuggestionsForGuestUsers]',
        description: 'Shows people picker suggestions for guest users. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--FileAnonymousLinkType [FileAnonymousLinkType]',
        description: 'Sets the file anonymous link type to None, View or Edit',
        autocomplete: this.getAnonymousLinkType()
      },
      {
        option: '--FolderAnonymousLinkType [FolderAnonymousLinkType]',
        description: 'Sets the folder anonymous link type to None, View or Edit',
        autocomplete: this.getAnonymousLinkType()
      },
      {
        option: '--NotifyOwnersWhenItemsReshared [NotifyOwnersWhenItemsReshared]',
        description: 'When this parameter is set to true and another user re-shares a document from a user\'s OneDrive for Business, the OneDrive for Business owner is notified by email. For additional information about how to configure notifications for external sharing, see Configure notifications for external sharing for OneDrive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NotifyOwnersWhenInvitationsAccepted [NotifyOwnersWhenInvitationsAccepted]',
        description: 'When this parameter is set to true and when an external user accepts an invitation to a resource in a user\'s OneDrive for Business, the OneDrive for Business owner is notified by email. For additional information about how to configure notifications for external sharing, see Configure notifications for external sharing for OneDrive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NotificationsInOneDriveForBusinessEnabled [NotificationsInOneDriveForBusinessEnabled]',
        description: 'Enables or disables notifications in OneDrive for business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NotificationsInSharePointEnabled [NotificationsInSharePointEnabled]',
        description: 'Enables or disables notifications in SharePoint. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OwnerAnonymousNotification [OwnerAnonymousNotification]',
        description: 'Enables or disables owner anonymous notification. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--CommentsOnSitePagesDisabled [CommentsOnSitePagesDisabled]',
        description: 'Enables or disables comments on site pages. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SocialBarOnSitePagesDisabled [SocialBarOnSitePagesDisabled]',
        description: 'Enables or disables social bar on site pages. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OrphanedPersonalSitesRetentionPeriod [OrphanedPersonalSitesRetentionPeriod]',
        description: 'Specifies the number of days after a user\'s Active Directory account is deleted that their OneDrive for Business content will be deleted. The value range is in days, between 30 and 3650. The default value is 30'
      },
      {
        option: '--DisallowInfectedFileDownload [DisallowInfectedFileDownload]',
        description: 'Prevents the Download button from being displayed on the Virus Found warning page. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DefaultLinkPermission [DefaultLinkPermission]',
        description: 'Choose the dafault permission that is selected when users share. This applies to anonymous access, internal and direct links. Allowed values None|View|Edit',
        autocomplete: this.getSharingPermissionType()
      },
      {
        option: '--ConditionalAccessPolicy [ConditionalAccessPolicy]',
        description: 'Configures conditional access policy. Allowed values AllowFullAccess|AllowLimitedAccess|BlockAccess',
        autocomplete: this.getSPOConditionalAccessPolicyType()
      },
      {
        option: '--AllowDownloadingNonWebViewableFiles [AllowDownloadingNonWebViewableFiles]',
        description: 'Allows downloading non web viewable files. The Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowEditing [AllowEditing]',
        description: 'Allows editing. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ApplyAppEnforcedRestrictionsToAdHocRecipients [ApplyAppEnforcedRestrictionsToAdHocRecipients]',
        description: 'Applies app enforced restrictions to AdHoc recipients. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--FilePickerExternalImageSearchEnabled [FilePickerExternalImageSearchEnabled]',
        description: 'Enables file picker external image search. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EmailAttestationRequired [EmailAttestationRequired]',
        description: 'Sets email attestation to required. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EmailAttestationReAuthDays [EmailAttestationReAuthDays]',
        description: 'Sets email attestation re-auth days'
      },
      {
        option: '--HideDefaultThemes [HideDefaultThemes]',
        description: 'Defines if the default themes are visible or hidden. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      // not included in the PnP PowerShell, most of them are new and maybe the cmdlet is not updated recently.
      {
        option: '--BlockAccessOnUnmanagedDevices [BlockAccessOnUnmanagedDevices]',
        description: 'Blocks access on unmanaged devices. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowLimitedAccessOnUnmanagedDevices [AllowLimitedAccessOnUnmanagedDevices]',
        description: 'Allows limited access on unmanaged devices blocks. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfAllFilesForGuests [BlockDownloadOfAllFilesForGuests]',
        description: 'Blocks download of all files for guests. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfAllFilesOnUnmanagedDevices [BlockDownloadOfAllFilesOnUnmanagedDevices]',
        description: 'Blocks download of all files on unmanaged devices. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfViewableFilesForGuests [BlockDownloadOfViewableFilesForGuests]',
        description: 'Blocks download of viewable files for guests. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfViewableFilesOnUnmanagedDevices [BlockDownloadOfViewableFilesOnUnmanagedDevices]',
        description: 'Blocks download of viewable files on unmanaged devices. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockMacSync [BlockMacSync]',
        description: 'Blocks Mac sync. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisableReportProblemDialog [DisableReportProblemDialog]',
        description: 'Disables report problem dialog. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisplayNamesOfFileViewers [DisplayNamesOfFileViewers]',
        description: 'Displayes names of file viewers. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EnableMinimumVersionRequirement [EnableMinimumVersionRequirement]',
        description: 'Enables minimum version requirement. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--HideSyncButtonOnODB [HideSyncButtonOnODB]',
        description: 'Hides the sync button on One Drive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsUnmanagedSyncClientForTenantRestricted [IsUnmanagedSyncClientForTenantRestricted]',
        description: 'Is unmanaged sync client for tenant restricted. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--LimitedAccessFileType [LimitedAccessFileType]',
        description: 'Allows users to preview only Office files in the browser. This option increases security but may be a barrier to user productivity. Allowed values OfficeOnlineFilesOnly|WebPreviewableFiles|OtherFiles',
        autocomplete: this.getSPOLimitedAccessFileType()
      },
      {
        option: '--OptOutOfGrooveBlock [OptOutOfGrooveBlock]',
        description: 'Opts out of the groove block. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OptOutOfGrooveSoftBlock [OptOutOfGrooveSoftBlock]',
        description: 'Opts out of Groove soft block. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OrgNewsSiteUrl [OrgNewsSiteUrl]',
        description: 'Organization news site url'
      },
      {
        option: '--PermissiveBrowserFileHandlingOverride [PermissiveBrowserFileHandlingOverride]',
        description: 'Permissive browser fileHandling override. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowNGSCDialogForSyncOnODB [ShowNGSCDialogForSyncOnODB]',
        description: 'Show NGSC dialog for sync on OneDrive for Business. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SpecialCharactersStateInFileFolderNames [SpecialCharactersStateInFileFolderNames]',
        description: 'Sets the special characters state in file and folder names in SharePoint and OneDrive for Business. Allowed values NoPreference|Allowed|Disallowed',
        autocomplete: this.getSpecialCharactersState()
      },
      {
        option: '--SyncPrivacyProfileProperties [SyncPrivacyProfileProperties]',
        description: 'Syncs privacy profile properties. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ExcludedFileExtensionsForSyncClient [ExcludedFileExtensionsForSyncClient]',
        description: 'Excluded file extensions for sync client. Array of strings split by comma (\',\')'
      },
      {
        option: '--AllowedDomainListForSyncClient [AllowedDomainListForSyncClient]',
        description: 'Sets allowed domain list for sync client. Array of GUIDs split by comma (\',\'). Example:c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201'
      },
      {
        option: '--DisabledWebPartIds [DisabledWebPartIds]',
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
        if (this.isExcludedOption(propertyKey)) {
          continue;
        }

        hasAtLeastOneOption = true;
        const propertyValue = opts[propertyKey];
        const commandOptions: CommandOption[] = this.options();

        for (let item of commandOptions) {
          if (item.option.indexOf(propertyKey) > -1 &&
            item.autocomplete &&
            item.autocomplete.indexOf(propertyValue.toString()) === -1) {
            return `${propertyKey} option has invalid value of ${propertyValue}. Allowed values are ${JSON.stringify(item.autocomplete)}`;
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
      case 'SharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'SharingDomainRestrictionMode':
        return this.getSharingDomainRestrictionModes().indexOf(value);
      case 'DefaultSharingLinkType':
        return this.getSharingLinkType().indexOf(value);
      case 'ODBMembersCanShare':
        return this.getSharingState().indexOf(value);
      case 'ODBAccessRequests':
        return this.getSharingState().indexOf(value);
      case 'FileAnonymousLinkType':
        return this.getAnonymousLinkType().indexOf(value);
      case 'FolderAnonymousLinkType':
        return this.getAnonymousLinkType().indexOf(value);
      case 'DefaultLinkPermission':
        return this.getSharingPermissionType().indexOf(value);
      case 'ConditionalAccessPolicy':
        return this.getSPOConditionalAccessPolicyType().indexOf(value);
      case 'LimitedAccessFileType':
        return this.getSPOLimitedAccessFileType().indexOf(value);
      case 'SpecialCharactersStateInFileFolderNames':
        return this.getSpecialCharactersState().indexOf(value);
      default:
        return -1;
    }
  }
}

module.exports = new SpoTenantSettingsSetCommand();