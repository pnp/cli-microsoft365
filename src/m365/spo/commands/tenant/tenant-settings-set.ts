import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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
  EnableAzureADB2BIntegration: boolean;
  SyncAadB2BManagementPolicy: boolean;
}

class SpoTenantSettingsSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets tenant global settings';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        MinCompatibilityLevel: (!(!args.options.MinCompatibilityLevel)).toString(),
        MaxCompatibilityLevel: (!(!args.options.MaxCompatibilityLevel)).toString(),
        ExternalServicesEnabled: (!(!args.options.ExternalServicesEnabled)).toString(),
        NoAccessRedirectUrl: (!(!args.options.NoAccessRedirectUrl)).toString(),
        SharingCapability: (!(!args.options.SharingCapability)).toString(),
        DisplayStartASiteOption: (!(!args.options.DisplayStartASiteOption)).toString(),
        StartASiteFormUrl: (!(!args.options.StartASiteFormUrl)).toString(),
        ShowEveryoneClaim: (!(!args.options.ShowEveryoneClaim)).toString(),
        ShowAllUsersClaim: (!(!args.options.ShowAllUsersClaim)).toString(),
        ShowEveryoneExceptExternalUsersClaim: (!(!args.options.ShowEveryoneExceptExternalUsersClaim)).toString(),
        SearchResolveExactEmailOrUPN: (!(!args.options.SearchResolveExactEmailOrUPN)).toString(),
        OfficeClientADALDisabled: (!(!args.options.OfficeClientADALDisabled)).toString(),
        LegacyAuthProtocolsEnabled: (!(!args.options.LegacyAuthProtocolsEnabled)).toString(),
        RequireAcceptingAccountMatchInvitedAccount: (!(!args.options.RequireAcceptingAccountMatchInvitedAccount)).toString(),
        ProvisionSharedWithEveryoneFolder: (!(!args.options.ProvisionSharedWithEveryoneFolder)).toString(),
        SignInAccelerationDomain: (!(!args.options.SignInAccelerationDomain)).toString(),
        EnableGuestSignInAcceleration: (!(!args.options.EnableGuestSignInAcceleration)).toString(),
        UsePersistentCookiesForExplorerView: (!(!args.options.UsePersistentCookiesForExplorerView)).toString(),
        BccExternalSharingInvitations: (!(!args.options.BccExternalSharingInvitations)).toString(),
        BccExternalSharingInvitationsList: (!(!args.options.BccExternalSharingInvitationsList)).toString(),
        UserVoiceForFeedbackEnabled: (!(!args.options.UserVoiceForFeedbackEnabled)).toString(),
        PublicCdnEnabled: (!(!args.options.PublicCdnEnabled)).toString(),
        PublicCdnAllowedFileTypes: (!(!args.options.PublicCdnAllowedFileTypes)).toString(),
        RequireAnonymousLinksExpireInDays: (!(!args.options.RequireAnonymousLinksExpireInDays)).toString(),
        SharingAllowedDomainList: (!(!args.options.SharingAllowedDomainList)).toString(),
        SharingBlockedDomainList: (!(!args.options.SharingBlockedDomainList)).toString(),
        SharingDomainRestrictionMode: (!(!args.options.SharingDomainRestrictionMode)).toString(),
        OneDriveStorageQuota: (!(!args.options.OneDriveStorageQuota)).toString(),
        OneDriveForGuestsEnabled: (!(!args.options.OneDriveForGuestsEnabled)).toString(),
        IPAddressEnforcement: (!(!args.options.IPAddressEnforcement)).toString(),
        IPAddressAllowList: (!(!args.options.IPAddressAllowList)).toString(),
        IPAddressWACTokenLifetime: (!(!args.options.IPAddressWACTokenLifetime)).toString(),
        UseFindPeopleInPeoplePicker: (!(!args.options.UseFindPeopleInPeoplePicker)).toString(),
        DefaultSharingLinkType: (!(!args.options.DefaultSharingLinkType)).toString(),
        ODBMembersCanShare: (!(!args.options.ODBMembersCanShare)).toString(),
        ODBAccessRequests: (!(!args.options.ODBAccessRequests)).toString(),
        PreventExternalUsersFromResharing: (!(!args.options.PreventExternalUsersFromResharing)).toString(),
        ShowPeoplePickerSuggestionsForGuestUsers: (!(!args.options.ShowPeoplePickerSuggestionsForGuestUsers)).toString(),
        FileAnonymousLinkType: (!(!args.options.FileAnonymousLinkType)).toString(),
        FolderAnonymousLinkType: (!(!args.options.FolderAnonymousLinkType)).toString(),
        NotifyOwnersWhenItemsReshared: (!(!args.options.NotifyOwnersWhenItemsReshared)).toString(),
        NotifyOwnersWhenInvitationsAccepted: (!(!args.options.NotifyOwnersWhenInvitationsAccepted)).toString(),
        NotificationsInOneDriveForBusinessEnabled: (!(!args.options.NotificationsInOneDriveForBusinessEnabled)).toString(),
        NotificationsInSharePointEnabled: (!(!args.options.NotificationsInSharePointEnabled)).toString(),
        OwnerAnonymousNotification: (!(!args.options.OwnerAnonymousNotification)).toString(),
        CommentsOnSitePagesDisabled: (!(!args.options.CommentsOnSitePagesDisabled)).toString(),
        SocialBarOnSitePagesDisabled: (!(!args.options.SocialBarOnSitePagesDisabled)).toString(),
        OrphanedPersonalSitesRetentionPeriod: (!(!args.options.OrphanedPersonalSitesRetentionPeriod)).toString(),
        DisallowInfectedFileDownload: (!(!args.options.DisallowInfectedFileDownload)).toString(),
        DefaultLinkPermission: (!(!args.options.DefaultLinkPermission)).toString(),
        ConditionalAccessPolicy: (!(!args.options.ConditionalAccessPolicy)).toString(),
        AllowDownloadingNonWebViewableFiles: (!(!args.options.AllowDownloadingNonWebViewableFiles)).toString(),
        AllowEditing: (!(!args.options.AllowEditing)).toString(),
        ApplyAppEnforcedRestrictionsToAdHocRecipients: (!(!args.options.ApplyAppEnforcedRestrictionsToAdHocRecipients)).toString(),
        FilePickerExternalImageSearchEnabled: (!(!args.options.FilePickerExternalImageSearchEnabled)).toString(),
        EmailAttestationRequired: (!(!args.options.EmailAttestationRequired)).toString(),
        EmailAttestationReAuthDays: (!(!args.options.EmailAttestationReAuthDays)).toString(),
        HideDefaultThemes: (!(!args.options.HideDefaultThemes)).toString(),
        BlockAccessOnUnmanagedDevices: (!(!args.options.BlockAccessOnUnmanagedDevices)).toString(),
        AllowLimitedAccessOnUnmanagedDevices: (!(!args.options.AllowLimitedAccessOnUnmanagedDevices)).toString(),
        BlockDownloadOfAllFilesForGuests: (!(!args.options.BlockDownloadOfAllFilesForGuests)).toString(),
        BlockDownloadOfAllFilesOnUnmanagedDevices: (!(!args.options.BlockDownloadOfAllFilesOnUnmanagedDevices)).toString(),
        BlockDownloadOfViewableFilesForGuests: (!(!args.options.BlockDownloadOfViewableFilesForGuests)).toString(),
        BlockDownloadOfViewableFilesOnUnmanagedDevices: (!(!args.options.BlockDownloadOfViewableFilesOnUnmanagedDevices)).toString(),
        BlockMacSync: (!(!args.options.BlockMacSync)).toString(),
        DisableReportProblemDialog: (!(!args.options.DisableReportProblemDialog)).toString(),
        DisplayNamesOfFileViewers: (!(!args.options.DisplayNamesOfFileViewers)).toString(),
        EnableMinimumVersionRequirement: (!(!args.options.EnableMinimumVersionRequirement)).toString(),
        HideSyncButtonOnODB: (!(!args.options.HideSyncButtonOnODB)).toString(),
        IsUnmanagedSyncClientForTenantRestricted: (!(!args.options.IsUnmanagedSyncClientForTenantRestricted)).toString(),
        LimitedAccessFileType: (!(!args.options.LimitedAccessFileType)).toString(),
        OptOutOfGrooveBlock: (!(!args.options.OptOutOfGrooveBlock)).toString(),
        OptOutOfGrooveSoftBlock: (!(!args.options.OptOutOfGrooveSoftBlock)).toString(),
        OrgNewsSiteUrl: (!(!args.options.OrgNewsSiteUrl)).toString(),
        PermissiveBrowserFileHandlingOverride: (!(!args.options.PermissiveBrowserFileHandlingOverride)).toString(),
        ShowNGSCDialogForSyncOnODB: (!(!args.options.ShowNGSCDialogForSyncOnODB)).toString(),
        SpecialCharactersStateInFileFolderNames: (!(!args.options.SpecialCharactersStateInFileFolderNames)).toString(),
        SyncPrivacyProfileProperties: (!(!args.options.SyncPrivacyProfileProperties)).toString(),
        ExcludedFileExtensionsForSyncClient: (!(!args.options.ExcludedFileExtensionsForSyncClient)).toString(),
        DisabledWebPartIds: (!(!args.options.DisabledWebPartIds)).toString(),
        AllowedDomainListForSyncClient: (!(!args.options.AllowedDomainListForSyncClient)).toString(),
        DisableCustomAppAuthentication: (!(!args.options.DisableCustomAppAuthentication)).toString(),
        EnableAzureADB2BIntegration: typeof args.options.EnableAzureADB2BIntegration !== 'undefined',
        SyncAadB2BManagementPolicy: typeof args.options.SyncAadB2BManagementPolicy !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
      },
      {
        option: '--EnableAzureADB2BIntegration [EnableAzureADB2BIntegration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SyncAadB2BManagementPolicy [SyncAadB2BManagementPolicy]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const opts: any = args.options;
        let hasAtLeastOneOption: boolean = false;

        for (const propertyKey of Object.keys(opts)) {
          if (this.isExcludedOption(propertyKey)) {
            continue;
          }

          hasAtLeastOneOption = true;
          const propertyValue = opts[propertyKey];

          for (const item of this.options) {
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
    );
  }

  #initTypes(): void {
    this.types.boolean.push(
      'EnableAzureADB2BIntegration',
      'SyncAadB2BManagementPolicy'
    );
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let formDigestValue = '';
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    spo
      .getTenantId(logger, this.debug)
      .then((_tenantId: string): Promise<string> => {
        tenantId = _tenantId;
        return spo.getSpoAdminUrl(logger, this.debug);
      })
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return spo.getRequestDigest(spoAdminUrl);
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
              valuesXml += `<Object Type="Guid">{${formatting.escapeXml(value)}}</Object>`;
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

        if (args.options.EnableAzureADB2BIntegration === true) {
          this.warn(logger, 'WARNING: Make sure to also enable the Azure AD one-time passcode authentication preview. If it is not enabled then SharePoint will not use Azure AD B2B even if EnableAzureADB2BIntegration is set to true. Learn more at http://aka.ms/spo-b2b-integration.');
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
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