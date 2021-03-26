import { Logger } from '../../../../cli';
import {
  CommandError, CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';

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
  DisableCustomAppAuthentication: boolean;
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
    telemetryProps.DisableCustomAppAuthentication = (!(!args.options.DisableCustomAppAuthentication)).toString();
    return telemetryProps;
  }

  public getAllEnumOptions(): string[] {
    return ['SharingCapability', 'SharingDomainRestrictionMode', 'DefaultSharingLinkType', 'ODBMembersCanShare', 'ODBAccessRequests', 'FileAnonymousLinkType', 'FolderAnonymousLinkType', 'DefaultLinkPermission', 'ConditionalAccessPolicy', 'LimitedAccessFileType', 'SpecialCharactersStateInFileFolderNames'];
  }

  // all enums as get methods
  private getSharingLinkType(): string[] { return ['None', 'Direct', 'Internal', 'AnonymousAccess']; }
  private getSharingCapabilities(): string[] { return ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']; }
  private getSharingDomainRestrictionModes(): string[] { return ['None', 'AllowList', 'BlockList']; }
  private getSharingState(): string[] { return ['Unspecified', 'On', 'Off']; }
  private getAnonymousLinkType(): string[] { return ['None', 'View', 'Edit']; }
  private getSharingPermissionType(): string[] { return ['None', 'View', 'Edit']; }
  private getSPOConditionalAccessPolicyType(): string[] { return ['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess']; }
  private getSpecialCharactersState(): string[] { return ['NoPreference', 'Allowed', 'Disallowed']; }
  private getSPOLimitedAccessFileType(): string[] { return ['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles']; }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    let formDigestValue = '';
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    this
      .getTenantId(logger, this.debug)
      .then((_tenantId: string): Promise<string> => {
        tenantId = _tenantId;
        return this.getSpoAdminUrl(logger, this.debug);
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
        for (const optionKey of Object.keys(args.options)) {
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
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${propsXml}</Actions><ObjectPaths><Identity Id="7" Name="${tenantId}" /></ObjectPaths></Request>`
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

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--MinCompatibilityLevel [MinCompatibilityLevel]'
      },
      {
        option: '--MaxCompatibilityLevel [MaxCompatibilityLevel]'
      },
      {
        option: '--ExternalServicesEnabled [ExternalServicesEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NoAccessRedirectUrl [NoAccessRedirectUrl]'
      },
      {
        option: '--SharingCapability [SharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--DisplayStartASiteOption [DisplayStartASiteOption]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--StartASiteFormUrl [StartASiteFormUrl]'
      },
      {
        option: '--ShowEveryoneClaim [ShowEveryoneClaim]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowAllUsersClaim [ShowAllUsersClaim]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowEveryoneExceptExternalUsersClaim [ShowEveryoneExceptExternalUsersClaim]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SearchResolveExactEmailOrUPN [SearchResolveExactEmailOrUPN]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OfficeClientADALDisabled [OfficeClientADALDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--LegacyAuthProtocolsEnabled [LegacyAuthProtocolsEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--RequireAcceptingAccountMatchInvitedAccount [RequireAcceptingAccountMatchInvitedAccount]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ProvisionSharedWithEveryoneFolder [ProvisionSharedWithEveryoneFolder]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SignInAccelerationDomain [SignInAccelerationDomain]'
      },
      {
        option: '--EnableGuestSignInAcceleration [EnableGuestSignInAcceleration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--UsePersistentCookiesForExplorerView [UsePersistentCookiesForExplorerView]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BccExternalSharingInvitations [BccExternalSharingInvitations]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BccExternalSharingInvitationsList [BccExternalSharingInvitationsList]'
      },
      {
        option: '--UserVoiceForFeedbackEnabled [UserVoiceForFeedbackEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--PublicCdnEnabled [PublicCdnEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--PublicCdnAllowedFileTypes [PublicCdnAllowedFileTypes]'
      },
      {
        option: '--RequireAnonymousLinksExpireInDays [RequireAnonymousLinksExpireInDays]'
      },
      {
        option: '--SharingAllowedDomainList [SharingAllowedDomainList]'
      },
      {
        option: '--SharingBlockedDomainList [SharingBlockedDomainList]'
      },
      {
        option: '--SharingDomainRestrictionMode [SharingDomainRestrictionMode]',
        autocomplete: this.getSharingDomainRestrictionModes()
      },
      {
        option: '--OneDriveStorageQuota [OneDriveStorageQuota]'
      },
      {
        option: '--OneDriveForGuestsEnabled [OneDriveForGuestsEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IPAddressEnforcement [IPAddressEnforcement]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IPAddressAllowList [IPAddressAllowList]'
      },
      {
        option: '--IPAddressWACTokenLifetime [IPAddressWACTokenLifetime]'
      },
      {
        option: '--UseFindPeopleInPeoplePicker [UseFindPeopleInPeoplePicker]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DefaultSharingLinkType [DefaultSharingLinkType]',
        autocomplete: this.getSharingLinkType()
      },
      {
        option: '--ODBMembersCanShare [ODBMembersCanShare]',
        autocomplete: this.getSharingState()
      },
      {
        option: '--ODBAccessRequests [ODBAccessRequests]',
        autocomplete: this.getSharingState()
      },
      {
        option: '--PreventExternalUsersFromResharing [PreventExternalUsersFromResharing]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowPeoplePickerSuggestionsForGuestUsers [ShowPeoplePickerSuggestionsForGuestUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--FileAnonymousLinkType [FileAnonymousLinkType]',
        autocomplete: this.getAnonymousLinkType()
      },
      {
        option: '--FolderAnonymousLinkType [FolderAnonymousLinkType]',
        autocomplete: this.getAnonymousLinkType()
      },
      {
        option: '--NotifyOwnersWhenItemsReshared [NotifyOwnersWhenItemsReshared]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NotifyOwnersWhenInvitationsAccepted [NotifyOwnersWhenInvitationsAccepted]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NotificationsInOneDriveForBusinessEnabled [NotificationsInOneDriveForBusinessEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--NotificationsInSharePointEnabled [NotificationsInSharePointEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OwnerAnonymousNotification [OwnerAnonymousNotification]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--CommentsOnSitePagesDisabled [CommentsOnSitePagesDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SocialBarOnSitePagesDisabled [SocialBarOnSitePagesDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OrphanedPersonalSitesRetentionPeriod [OrphanedPersonalSitesRetentionPeriod]'
      },
      {
        option: '--DisallowInfectedFileDownload [DisallowInfectedFileDownload]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DefaultLinkPermission [DefaultLinkPermission]',
        autocomplete: this.getSharingPermissionType()
      },
      {
        option: '--ConditionalAccessPolicy [ConditionalAccessPolicy]',
        autocomplete: this.getSPOConditionalAccessPolicyType()
      },
      {
        option: '--AllowDownloadingNonWebViewableFiles [AllowDownloadingNonWebViewableFiles]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowEditing [AllowEditing]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ApplyAppEnforcedRestrictionsToAdHocRecipients [ApplyAppEnforcedRestrictionsToAdHocRecipients]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--FilePickerExternalImageSearchEnabled [FilePickerExternalImageSearchEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EmailAttestationRequired [EmailAttestationRequired]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EmailAttestationReAuthDays [EmailAttestationReAuthDays]'
      },
      {
        option: '--HideDefaultThemes [HideDefaultThemes]',
        autocomplete: ['true', 'false']
      },
      // not included in the PnP PowerShell, most of them are new and maybe the cmdlet is not updated recently.
      {
        option: '--BlockAccessOnUnmanagedDevices [BlockAccessOnUnmanagedDevices]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowLimitedAccessOnUnmanagedDevices [AllowLimitedAccessOnUnmanagedDevices]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfAllFilesForGuests [BlockDownloadOfAllFilesForGuests]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfAllFilesOnUnmanagedDevices [BlockDownloadOfAllFilesOnUnmanagedDevices]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfViewableFilesForGuests [BlockDownloadOfViewableFilesForGuests]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockDownloadOfViewableFilesOnUnmanagedDevices [BlockDownloadOfViewableFilesOnUnmanagedDevices]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockMacSync [BlockMacSync]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisableReportProblemDialog [DisableReportProblemDialog]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisplayNamesOfFileViewers [DisplayNamesOfFileViewers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EnableMinimumVersionRequirement [EnableMinimumVersionRequirement]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--HideSyncButtonOnODB [HideSyncButtonOnODB]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsUnmanagedSyncClientForTenantRestricted [IsUnmanagedSyncClientForTenantRestricted]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--LimitedAccessFileType [LimitedAccessFileType]',
        autocomplete: this.getSPOLimitedAccessFileType()
      },
      {
        option: '--OptOutOfGrooveBlock [OptOutOfGrooveBlock]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OptOutOfGrooveSoftBlock [OptOutOfGrooveSoftBlock]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OrgNewsSiteUrl [OrgNewsSiteUrl]'
      },
      {
        option: '--PermissiveBrowserFileHandlingOverride [PermissiveBrowserFileHandlingOverride]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowNGSCDialogForSyncOnODB [ShowNGSCDialogForSyncOnODB]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SpecialCharactersStateInFileFolderNames [SpecialCharactersStateInFileFolderNames]',
        autocomplete: this.getSpecialCharactersState()
      },
      {
        option: '--SyncPrivacyProfileProperties [SyncPrivacyProfileProperties]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ExcludedFileExtensionsForSyncClient [ExcludedFileExtensionsForSyncClient]'
      },
      {
        option: '--AllowedDomainListForSyncClient [AllowedDomainListForSyncClient]'
      },
      {
        option: '--DisabledWebPartIds [DisabledWebPartIds]'
      },
      {
        option: '--DisableCustomAppAuthentication [DisableCustomAppAuthentication]',
        autocomplete: ['true', 'false']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const opts: any = args.options;
    let hasAtLeastOneOption: boolean = false;

    for (const propertyKey of Object.keys(opts)) {
      if (this.isExcludedOption(propertyKey)) {
        continue;
      }

      hasAtLeastOneOption = true;
      const propertyValue = opts[propertyKey];
      const commandOptions: CommandOption[] = this.options();

      for (const item of commandOptions) {
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
  }

  public isExcludedOption(optionKey: string): boolean {
    // it is not possible to dynamically get the GlobalOptions
    // prop keys since they are nullable
    // so we have to maintain that array bellow once new global option
    // is added to the GlobalOptions interface
    return ['output', 'o', 'debug', 'verbose', '_', 'query'].indexOf(optionKey) > -1;
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