import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  MinCompatibilityLevel: number;
  MaxCompatibilityLevel: number;
  ExternalServicesEnabled?: boolean;
  NoAccessRedirectUrl: string;
  ArchiveRedirectUrl: string;
  ConditionalAccessPolicyErrorHelpLink: string;
  CustomizedExternalSharingServiceUrl: string;
  LabelMismatchEmailHelpLink: string;
  SharingCapability: string; // <SharingCapabilities>
  CoreSharingCapability: string; // <SharingCapabilities>
  ODBSharingCapability: string; // <SharingCapabilities>
  OneDriveLoopSharingCapability: string; // <SharingCapabilities>
  CoreLoopSharingCapability: string; // <SharingCapabilities>
  ContainerSharingCapability: string; // <SharingCapabilities>
  CoreDefaultShareLinkRole: string; // <Role>
  CoreLoopDefaultSharingLinkRole: string; // <Role>
  ContainerDefaultShareLinkRole: string; // <Role>
  ContainerLoopDefaultShareLinkRole: string; // <Role>
  OneDriveDefaultShareLinkRole: string; // <Role>
  OneDriveLoopDefaultSharingLinkRole: string; // <Role>
  CoreDefaultShareLinkScope: string; // <SharingScope>
  CoreLoopDefaultSharingLinkScope: string; // <SharingScope>
  ContainerDefaultShareLinkScope: string; // <SharingScope>
  ContainerLoopDefaultShareLinkScope: string; // <SharingScope>
  OneDriveDefaultShareLinkScope: string; // <SharingScope>
  OneDriveLoopDefaultSharingLinkScope: string; // <SharingScope>
  DisplayStartASiteOption?: boolean;
  StartASiteFormUrl: string;
  ShowEveryoneClaim?: boolean;
  ShowAllUsersClaim?: boolean;
  ShowEveryoneExceptExternalUsersClaim?: boolean;
  SearchResolveExactEmailOrUPN?: boolean;
  OfficeClientADALDisabled?: boolean;
  LegacyAuthProtocolsEnabled?: boolean;
  RequireAcceptingAccountMatchInvitedAccount?: boolean;
  ProvisionSharedWithEveryoneFolder?: boolean;
  SignInAccelerationDomain: string;
  EnableGuestSignInAcceleration?: boolean;
  UsePersistentCookiesForExplorerView?: boolean;
  BccExternalSharingInvitations?: boolean;
  BccExternalSharingInvitationsList: string;
  UserVoiceForFeedbackEnabled?: boolean;
  PublicCdnEnabled?: boolean;
  PublicCdnAllowedFileTypes: string;
  RequireAnonymousLinksExpireInDays: number;
  SharingAllowedDomainList: string;
  SharingBlockedDomainList: string;
  SharingDomainRestrictionMode: string; // <SharingDomainRestrictionModes>
  OneDriveStorageQuota: number;
  OneDriveForGuestsEnabled?: boolean;
  IPAddressEnforcement?: boolean;
  IPAddressAllowList: string;
  IPAddressWACTokenLifetime: number;
  UseFindPeopleInPeoplePicker?: boolean;
  DefaultSharingLinkType: string; // <SharingLinkType>
  ODBMembersCanShare: string; // <SharingState>
  ODBAccessRequests: string; // <SharingState>
  AllowAnonymousMeetingParticipantsToAccessWhiteboards: string; // <SharingState>
  PreventExternalUsersFromResharing?: boolean;
  ShowPeoplePickerSuggestionsForGuestUsers?: boolean;
  FileAnonymousLinkType: string; // <AnonymousLinkType>
  FolderAnonymousLinkType: string; // <AnonymousLinkType>
  NotifyOwnersWhenItemsReshared?: boolean;
  NotifyOwnersWhenInvitationsAccepted?: boolean;
  NotificationsInOneDriveForBusinessEnabled?: boolean;
  NotificationsInSharePointEnabled?: boolean;
  OwnerAnonymousNotification?: boolean;
  CommentsOnSitePagesDisabled?: boolean;
  SocialBarOnSitePagesDisabled?: boolean;
  OrphanedPersonalSitesRetentionPeriod?: number;
  CoreRequestFilesLinkExpirationInDays?: number;
  OneDriveRequestFilesLinkExpirationInDays?: number;
  ExternalUserExpireInDays: number;
  ReduceTempTokenLifetimeEnabled?: boolean;
  ReduceTempTokenLifetimeValue: number;
  ShowOpenInDesktopOptionForSyncedFiles?: boolean;
  ShowPeoplePickerGroupSuggestionsForIB?: boolean;
  SiteOwnerManageLegacyServicePrincipalEnabled?: boolean;
  StopNew2010Workflows?: boolean;
  StopNew2013Workflows?: boolean;
  ViewersCanCommentOnMediaDisabled?: boolean;
  AllowEveryoneExceptExternalUsersClaimInPrivateSite?: boolean;
  AnyoneLinkTrackUsers?: boolean;
  HasAdminCompletedCUConfiguration?: boolean;
  HasIntelligentContentServicesCapability?: boolean;
  HasTopicExperiencesCapability?: boolean;
  MachineLearningCaptureEnabled?: boolean;
  MassDeleteNotificationDisabled?: boolean;
  MobileFriendlyUrlEnabledInTenant?: boolean;
  DisallowInfectedFileDownload?: boolean;
  DefaultLinkPermission: string; // <SharingPermissionType>
  ConditionalAccessPolicy: string; // <SPOConditionalAccessPolicyType>
  AllowDownloadingNonWebViewableFiles?: boolean;
  AllowEditing?: boolean;
  ApplyAppEnforcedRestrictionsToAdHocRecipients?: boolean;
  FilePickerExternalImageSearchEnabled?: boolean;
  EmailAttestationRequired?: boolean;
  EmailAttestationReAuthDays: number;
  HideDefaultThemes?: boolean;
  // not included in the PnP PowerShell, most of them are new and maybe the cmdlet is not updated recently.
  BlockAccessOnUnmanagedDevices?: boolean;
  AllowLimitedAccessOnUnmanagedDevices?: boolean;
  BlockDownloadOfAllFilesForGuests?: boolean;
  BlockDownloadOfAllFilesOnUnmanagedDevices?: boolean;
  BlockDownloadOfViewableFilesForGuests?: boolean;
  BlockDownloadOfViewableFilesOnUnmanagedDevices?: boolean;
  BlockMacSync?: boolean;
  DisableReportProblemDialog?: boolean;
  DisplayNamesOfFileViewers?: boolean;
  EnableMinimumVersionRequirement?: boolean;
  HideSyncButtonOnODB?: boolean;
  IsUnmanagedSyncClientForTenantRestricted?: boolean;
  LimitedAccessFileType: string; // <LimitedAccessFileType>
  MediaTranscription: string // <MediaTranscriptionPolicyType>
  MediaTranscriptionAutomaticFeatures: string // <MediaTranscriptionAutomaticFeaturesPolicyType>
  ImageTaggingOption: string // <ImageTaggingChoice>
  MarkNewFilesSensitiveByDefault: string; // <SensitiveByDefaultState>
  OCRAdminSiteListFileName: string;
  OCRComplianceSiteListFileName: string;
  OCRModeForAdminSites: string; // <ObjectCharacterRecognitionMode>
  OCRModeForComplianceODBs: string; // <ObjectCharacterRecognitionMode>
  OCRModeForComplianceSites: string; // <ObjectCharacterRecognitionMode>
  OneDriveDefaultLinkToExistingAccess: boolean;
  ContainerDefaultLinkToExistingAccess: boolean;
  OptOutOfGrooveBlock?: boolean;
  OptOutOfGrooveSoftBlock?: boolean;
  OrgNewsSiteUrl: string;
  PermissiveBrowserFileHandlingOverride?: boolean;
  ShowNGSCDialogForSyncOnODB?: boolean;
  SpecialCharactersStateInFileFolderNames: string; // <SpecialCharactersState>
  SyncPrivacyProfileProperties?: boolean;
  ExcludedFileExtensionsForSyncClient: string[];
  AllowedDomainListForSyncClient: string[];
  DisabledWebPartIds: string[];
  DisabledModernListTemplateIds: string[];
  DisableCustomAppAuthentication?: boolean;
  CommentsOnListItemsDisabled?: boolean;
  EnableAzureADB2BIntegration?: boolean;
  EnableAutoNewsDigest?: boolean;
  AllowCommentsTextOnEmailEnabled?: boolean;
  CommentsOnFilesDisabled?: boolean;
  DisableAddToOneDrive?: boolean;
  DisableBackToClassic?: boolean;
  DisablePersonalListCreation?: boolean;
  ViewInFileExplorerEnabled?: boolean;
  AllowGuestUserShareToUsersNotInSiteCollection?: boolean;
  BlockSendLabelMismatchEmail?: boolean;
  CoreDefaultLinkToExistingAccess?: boolean;
  CoreRequestFilesLinkEnabled?: boolean;
  OneDriveRequestFilesLinkEnabled?: boolean;
  DisableDocumentLibraryDefaultLabeling?: boolean;
  DisableVivaConnectionsAnalytics?: boolean;
  DisplayNamesOfFileViewersInSpo?: boolean;
  EnableAIPIntegration?: boolean;
  EnableRestrictedAccessControl?: boolean;
  ExternalUserExpirationRequired?: boolean;
  HideSyncButtonOnDocLib?: boolean;
  IncludeAtAGlanceInShareEmails?: boolean;
  InformationBarriersSuspension?: boolean;
  IsFluidEnabled?: boolean;
  IsWBFluidEnabled?: boolean;
  IsCollabMeetingNotesFluidEnabled?: boolean;
  IsEnableAppAuthPopUpEnabled?: boolean;
  IsLoopEnabled?: boolean;
  SyncAadB2BManagementPolicy?: boolean;
}

class SpoTenantSettingsSetCommand extends SpoCommand {
  private static booleanOptions: string[] = [
    'ExternalServicesEnabled',
    'DisplayStartASiteOption',
    'ShowEveryoneClaim',
    'ShowAllUsersClaim',
    'ShowEveryoneExceptExternalUsersClaim',
    'SearchResolveExactEmailOrUPN',
    'OfficeClientADALDisabled',
    'LegacyAuthProtocolsEnabled',
    'RequireAcceptingAccountMatchInvitedAccount',
    'ProvisionSharedWithEveryoneFolder',
    'EnableGuestSignInAcceleration',
    'UsePersistentCookiesForExplorerView',
    'BccExternalSharingInvitations',
    'UserVoiceForFeedbackEnabled',
    'PublicCdnEnabled',
    'OneDriveForGuestsEnabled',
    'IPAddressEnforcement',
    'UseFindPeopleInPeoplePicker',
    'PreventExternalUsersFromResharing',
    'ShowPeoplePickerSuggestionsForGuestUsers',
    'NotifyOwnersWhenItemsReshared',
    'NotifyOwnersWhenInvitationsAccepted',
    'NotificationsInOneDriveForBusinessEnabled',
    'NotificationsInSharePointEnabled',
    'OwnerAnonymousNotification',
    'CommentsOnSitePagesDisabled',
    'SocialBarOnSitePagesDisabled',
    'ReduceTempTokenLifetimeEnabled',
    'ShowOpenInDesktopOptionForSyncedFiles',
    'ShowPeoplePickerGroupSuggestionsForIB',
    'SiteOwnerManageLegacyServicePrincipalEnabled',
    'StopNew2010Workflows',
    'StopNew2013Workflows',
    'ViewersCanCommentOnMediaDisabled',
    'AllowEveryoneExceptExternalUsersClaimInPrivateSite',
    'AnyoneLinkTrackUsers',
    'HasAdminCompletedCUConfiguration',
    'HasIntelligentContentServicesCapability',
    'HasTopicExperiencesCapability',
    'MachineLearningCaptureEnabled',
    'MassDeleteNotificationDisabled',
    'MobileFriendlyUrlEnabledInTenant',
    'DisallowInfectedFileDownload',
    'AllowDownloadingNonWebViewableFiles',
    'AllowEditing',
    'ApplyAppEnforcedRestrictionsToAdHocRecipients',
    'FilePickerExternalImageSearchEnabled',
    'EmailAttestationRequired',
    'HideDefaultThemes',
    'BlockAccessOnUnmanagedDevices',
    'AllowLimitedAccessOnUnmanagedDevices',
    'BlockDownloadOfAllFilesForGuests',
    'BlockDownloadOfAllFilesOnUnmanagedDevices',
    'BlockDownloadOfViewableFilesForGuests',
    'BlockDownloadOfViewableFilesOnUnmanagedDevices',
    'BlockMacSync',
    'DisableReportProblemDialog',
    'DisplayNamesOfFileViewers',
    'EnableMinimumVersionRequirement',
    'HideSyncButtonOnODB',
    'IsUnmanagedSyncClientForTenantRestricted',
    'OneDriveDefaultLinkToExistingAccess',
    'ContainerDefaultLinkToExistingAccess',
    'OptOutOfGrooveBlock',
    'OptOutOfGrooveSoftBlock',
    'PermissiveBrowserFileHandlingOverride',
    'ShowNGSCDialogForSyncOnODB',
    'SyncPrivacyProfileProperties',
    'DisableCustomAppAuthentication',
    'CommentsOnListItemsDisabled',
    'EnableAzureADB2BIntegration',
    'EnableAutoNewsDigest',
    'AllowCommentsTextOnEmailEnabled',
    'CommentsOnFilesDisabled',
    'DisableAddToOneDrive',
    'DisableBackToClassic',
    'DisablePersonalListCreation',
    'ViewInFileExplorerEnabled',
    'AllowGuestUserShareToUsersNotInSiteCollection',
    'BlockSendLabelMismatchEmail',
    'CoreDefaultLinkToExistingAccess',
    'CoreRequestFilesLinkEnabled',
    'OneDriveRequestFilesLinkEnabled',
    'DisableDocumentLibraryDefaultLabeling',
    'DisableVivaConnectionsAnalytics',
    'DisplayNamesOfFileViewersInSpo',
    'EnableAIPIntegration',
    'EnableRestrictedAccessControl',
    'ExternalUserExpirationRequired',
    'HideSyncButtonOnDocLib',
    'IncludeAtAGlanceInShareEmails',
    'InformationBarriersSuspension',
    'IsFluidEnabled',
    'IsWBFluidEnabled',
    'IsCollabMeetingNotesFluidEnabled',
    'IsEnableAppAuthPopUpEnabled',
    'IsLoopEnabled',
    'SyncAadB2BManagementPolicy'
  ];

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
      const telemetryProps: any = {
        MinCompatibilityLevel: (!(!args.options.MinCompatibilityLevel)).toString(),
        MaxCompatibilityLevel: (!(!args.options.MaxCompatibilityLevel)).toString(),
        NoAccessRedirectUrl: (!(!args.options.NoAccessRedirectUrl)).toString(),
        ArchiveRedirectUrl: (!(!args.options.ArchiveRedirectUrl)).toString(),
        OCRAdminSiteListFileName: (!(!args.options.OCRAdminSiteListFileName)).toString(),
        OCRComplianceSiteListFileName: (!(!args.options.OCRComplianceSiteListFileName)).toString(),
        ConditionalAccessPolicyErrorHelpLink: (!(!args.options.ConditionalAccessPolicyErrorHelpLink)).toString(),
        CustomizedExternalSharingServiceUrl: (!(!args.options.CustomizedExternalSharingServiceUrl)).toString(),
        LabelMismatchEmailHelpLink: (!(!args.options.LabelMismatchEmailHelpLink)).toString(),
        SharingCapability: (!(!args.options.SharingCapability)).toString(),
        CoreSharingCapability: (!(!args.options.CoreSharingCapability)).toString(),
        ODBSharingCapability: (!(!args.options.ODBSharingCapability)).toString(),
        OneDriveLoopSharingCapability: (!(!args.options.OneDriveLoopSharingCapability)).toString(),
        CoreLoopSharingCapability: (!(!args.options.CoreLoopSharingCapability)).toString(),
        ContainerSharingCapability: (!(!args.options.ContainerSharingCapability)).toString(),
        CoreDefaultShareLinkRole: (!(!args.options.CoreDefaultShareLinkRole)).toString(),
        CoreLoopDefaultSharingLinkRole: (!(!args.options.CoreLoopDefaultSharingLinkRole)).toString(),
        ContainerDefaultShareLinkRole: (!(!args.options.ContainerDefaultShareLinkRole)).toString(),
        ContainerLoopDefaultShareLinkRole: (!(!args.options.ContainerLoopDefaultShareLinkRole)).toString(),
        OneDriveDefaultShareLinkRole: (!(!args.options.OneDriveDefaultShareLinkRole)).toString(),
        OneDriveLoopDefaultSharingLinkRole: (!(!args.options.OneDriveLoopDefaultSharingLinkRole)).toString(),
        CoreDefaultShareLinkScope: (!(!args.options.CoreDefaultShareLinkScope)).toString(),
        CoreLoopDefaultSharingLinkScope: (!(!args.options.CoreLoopDefaultSharingLinkScope)).toString(),
        ContainerDefaultShareLinkScope: (!(!args.options.ContainerDefaultShareLinkScope)).toString(),
        ContainerLoopDefaultShareLinkScope: (!(!args.options.ContainerLoopDefaultShareLinkScope)).toString(),
        OneDriveDefaultShareLinkScope: (!(!args.options.OneDriveDefaultShareLinkScope)).toString(),
        OneDriveLoopDefaultSharingLinkScope: (!(!args.options.OneDriveLoopDefaultSharingLinkScope)).toString(),
        StartASiteFormUrl: (!(!args.options.StartASiteFormUrl)).toString(),
        SignInAccelerationDomain: (!(!args.options.SignInAccelerationDomain)).toString(),
        BccExternalSharingInvitationsList: (!(!args.options.BccExternalSharingInvitationsList)).toString(),
        PublicCdnAllowedFileTypes: (!(!args.options.PublicCdnAllowedFileTypes)).toString(),
        RequireAnonymousLinksExpireInDays: (!(!args.options.RequireAnonymousLinksExpireInDays)).toString(),
        SharingAllowedDomainList: (!(!args.options.SharingAllowedDomainList)).toString(),
        SharingBlockedDomainList: (!(!args.options.SharingBlockedDomainList)).toString(),
        SharingDomainRestrictionMode: (!(!args.options.SharingDomainRestrictionMode)).toString(),
        OneDriveStorageQuota: (!(!args.options.OneDriveStorageQuota)).toString(),
        IPAddressAllowList: (!(!args.options.IPAddressAllowList)).toString(),
        IPAddressWACTokenLifetime: (!(!args.options.IPAddressWACTokenLifetime)).toString(),
        DefaultSharingLinkType: (!(!args.options.DefaultSharingLinkType)).toString(),
        ODBMembersCanShare: (!(!args.options.ODBMembersCanShare)).toString(),
        ODBAccessRequests: (!(!args.options.ODBAccessRequests)).toString(),
        AllowAnonymousMeetingParticipantsToAccessWhiteboards: (!(!args.options.AllowAnonymousMeetingParticipantsToAccessWhiteboards)).toString(),
        FileAnonymousLinkType: (!(!args.options.FileAnonymousLinkType)).toString(),
        FolderAnonymousLinkType: (!(!args.options.FolderAnonymousLinkType)).toString(),
        OrphanedPersonalSitesRetentionPeriod: (!(!args.options.OrphanedPersonalSitesRetentionPeriod)).toString(),
        CoreRequestFilesLinkExpirationInDays: (!(!args.options.CoreRequestFilesLinkExpirationInDays)).toString(),
        OneDriveRequestFilesLinkExpirationInDays: (!(!args.options.OneDriveRequestFilesLinkExpirationInDays)).toString(),
        ExternalUserExpireInDays: (!(!args.options.ExternalUserExpireInDays)).toString(),
        ReduceTempTokenLifetimeValue: (!(!args.options.ReduceTempTokenLifetimeValue)).toString(),
        DefaultLinkPermission: (!(!args.options.DefaultLinkPermission)).toString(),
        ConditionalAccessPolicy: (!(!args.options.ConditionalAccessPolicy)).toString(),
        EmailAttestationReAuthDays: (!(!args.options.EmailAttestationReAuthDays)).toString(),
        LimitedAccessFileType: (!(!args.options.LimitedAccessFileType)).toString(),
        MediaTranscription: (!(!args.options.MediaTranscription)).toString(),
        MediaTranscriptionAutomaticFeatures: (!(!args.options.MediaTranscriptionAutomaticFeatures)).toString(),
        ImageTaggingOption: (!(!args.options.ImageTaggingOption)).toString(),
        MarkNewFilesSensitiveByDefault: (!(!args.options.MarkNewFilesSensitiveByDefault)).toString(),
        OCRModeForAdminSites: (!(!args.options.OCRModeForAdminSites)).toString(),
        OCRModeForComplianceODBs: (!(!args.options.OCRModeForComplianceODBs)).toString(),
        OCRModeForComplianceSites: (!(!args.options.OCRModeForComplianceSites)).toString(),
        OrgNewsSiteUrl: (!(!args.options.OrgNewsSiteUrl)).toString(),
        SpecialCharactersStateInFileFolderNames: (!(!args.options.SpecialCharactersStateInFileFolderNames)).toString(),
        ExcludedFileExtensionsForSyncClient: (!(!args.options.ExcludedFileExtensionsForSyncClient)).toString(),
        DisabledWebPartIds: (!(!args.options.DisabledWebPartIds)).toString(),
        DisabledModernListTemplateIds: (!(!args.options.DisabledModernListTemplateIds)).toString(),
        AllowedDomainListForSyncClient: (!(!args.options.AllowedDomainListForSyncClient)).toString()
      };

      // add boolean values
      SpoTenantSettingsSetCommand.booleanOptions.forEach(o => {
        const value: boolean = (args.options as any)[o];
        if (value !== undefined) {
          telemetryProps[o] = value;
        }
      });

      Object.assign(this.telemetryProperties, telemetryProps);
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
        option: '--ArchiveRedirectUrl [ArchiveRedirectUrl]'
      },
      {
        option: '--OCRAdminSiteListFileName [OCRAdminSiteListFileName]'
      },
      {
        option: '--OCRComplianceSiteListFileName [OCRComplianceSiteListFileName]'
      },
      {
        option: '--ConditionalAccessPolicyErrorHelpLink [ConditionalAccessPolicyErrorHelpLink]'
      },
      {
        option: '--CustomizedExternalSharingServiceUrl [CustomizedExternalSharingServiceUrl]'
      },
      {
        option: '--LabelMismatchEmailHelpLink [LabelMismatchEmailHelpLink]'
      },
      {
        option: '--SharingCapability [SharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--CoreSharingCapability [CoreSharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--ODBSharingCapability [ODBSharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--OneDriveLoopSharingCapability [OneDriveLoopSharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--CoreLoopSharingCapability [CoreLoopSharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--ContainerSharingCapability [ContainerSharingCapability]',
        autocomplete: this.getSharingCapabilities()
      },
      {
        option: '--CoreDefaultShareLinkRole [CoreDefaultShareLinkRole]',
        autocomplete: this.getRole()
      },
      {
        option: '--CoreLoopDefaultSharingLinkRole [CoreLoopDefaultSharingLinkRole]',
        autocomplete: this.getRole()
      },
      {
        option: '--ContainerDefaultShareLinkRole [ContainerDefaultShareLinkRole]',
        autocomplete: this.getRole()
      },
      {
        option: '--ContainerLoopDefaultShareLinkRole [ContainerLoopDefaultShareLinkRole]',
        autocomplete: this.getRole()
      },
      {
        option: '--OneDriveDefaultShareLinkRole [OneDriveDefaultShareLinkRole]',
        autocomplete: this.getRole()
      },
      {
        option: '--OneDriveLoopDefaultSharingLinkRole [OneDriveLoopDefaultSharingLinkRole]',
        autocomplete: this.getRole()
      },
      {
        option: '--CoreDefaultShareLinkScope [CoreDefaultShareLinkScope]',
        autocomplete: this.getSharingScope()
      },
      {
        option: '--CoreLoopDefaultSharingLinkScope [CoreLoopDefaultSharingLinkScope]',
        autocomplete: this.getSharingScope()
      },
      {
        option: '--ContainerDefaultShareLinkScope [ContainerDefaultShareLinkScope]',
        autocomplete: this.getSharingScope()
      },
      {
        option: '--ContainerLoopDefaultShareLinkScope [ContainerLoopDefaultShareLinkScope]',
        autocomplete: this.getSharingScope()
      },
      {
        option: '--OneDriveDefaultShareLinkScope [OneDriveDefaultShareLinkScope]',
        autocomplete: this.getSharingScope()
      },
      {
        option: '--OneDriveLoopDefaultSharingLinkScope [OneDriveLoopDefaultSharingLinkScope]',
        autocomplete: this.getSharingScope()
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
        option: '--AllowAnonymousMeetingParticipantsToAccessWhiteboards [AllowAnonymousMeetingParticipantsToAccessWhiteboards]',
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
        option: '--CoreRequestFilesLinkExpirationInDays [CoreRequestFilesLinkExpirationInDays]'
      },
      {
        option: '--OneDriveRequestFilesLinkExpirationInDays [OneDriveRequestFilesLinkExpirationInDays]'
      },
      {
        option: '--ExternalUserExpireInDays [ExternalUserExpireInDays]'
      },
      {
        option: '--ReduceTempTokenLifetimeValue [ReduceTempTokenLifetimeValue]'
      },
      {
        option: '--ReduceTempTokenLifetimeEnabled [ReduceTempTokenLifetimeEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowOpenInDesktopOptionForSyncedFiles [ShowOpenInDesktopOptionForSyncedFiles]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ShowPeoplePickerGroupSuggestionsForIB [ShowPeoplePickerGroupSuggestionsForIB]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--SiteOwnerManageLegacyServicePrincipalEnabled [SiteOwnerManageLegacyServicePrincipalEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--StopNew2010Workflows [StopNew2010Workflows]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--StopNew2013Workflows [StopNew2013Workflows]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ViewersCanCommentOnMediaDisabled [ViewersCanCommentOnMediaDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowEveryoneExceptExternalUsersClaimInPrivateSite [AllowEveryoneExceptExternalUsersClaimInPrivateSite]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AnyoneLinkTrackUsers [AnyoneLinkTrackUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--HasAdminCompletedCUConfiguration [HasAdminCompletedCUConfiguration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--HasIntelligentContentServicesCapability [HasIntelligentContentServicesCapability]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--HasTopicExperiencesCapability [HasTopicExperiencesCapability]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--MachineLearningCaptureEnabled [MachineLearningCaptureEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--MassDeleteNotificationDisabled [MassDeleteNotificationDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--MobileFriendlyUrlEnabledInTenant [MobileFriendlyUrlEnabledInTenant]',
        autocomplete: ['true', 'false']
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
        option: '--MediaTranscription [MediaTranscription]',
        autocomplete: this.getMediaTranscriptionPolicyType()
      },
      {
        option: '--MediaTranscriptionAutomaticFeatures [MediaTranscriptionAutomaticFeatures]',
        autocomplete: this.getMediaTranscriptionAutomaticFeaturesPolicyType()
      },
      {
        option: '--ImageTaggingOption [ImageTaggingOption]',
        autocomplete: this.getImageTaggingChoice()
      },
      {
        option: '--MarkNewFilesSensitiveByDefault [MarkNewFilesSensitiveByDefault]',
        autocomplete: this.getSensitiveByDefaultState()
      },
      {
        option: '--OCRModeForAdminSites [OCRModeForAdminSites]',
        autocomplete: this.getObjectCharacterRecognitionMode()
      },
      {
        option: '--OCRModeForComplianceODBs [OCRModeForComplianceODBs]',
        autocomplete: this.getObjectCharacterRecognitionMode()
      },
      {
        option: '--OCRModeForComplianceSites [OCRModeForComplianceSites]',
        autocomplete: this.getObjectCharacterRecognitionMode()
      },
      {
        option: '--OneDriveDefaultLinkToExistingAccess [OneDriveDefaultLinkToExistingAccess]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ContainerDefaultLinkToExistingAccess [ContainerDefaultLinkToExistingAccess]',
        autocomplete: ['true', 'false']
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
        option: '--DisabledModernListTemplateIds [DisabledModernListTemplateIds]'
      },
      {
        option: '--DisableCustomAppAuthentication [DisableCustomAppAuthentication]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--CommentsOnListItemsDisabled [CommentsOnListItemsDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EnableAzureADB2BIntegration [EnableAzureADB2BIntegration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EnableAutoNewsDigest [EnableAutoNewsDigest]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowCommentsTextOnEmailEnabled [AllowCommentsTextOnEmailEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--CommentsOnFilesDisabled [CommentsOnFilesDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisableAddToOneDrive [DisableAddToOneDrive]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisableBackToClassic [DisableBackToClassic]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisablePersonalListCreation [DisablePersonalListCreation]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ViewInFileExplorerEnabled [ViewInFileExplorerEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--AllowGuestUserShareToUsersNotInSiteCollection [AllowGuestUserShareToUsersNotInSiteCollection]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--BlockSendLabelMismatchEmail [BlockSendLabelMismatchEmail]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--CoreDefaultLinkToExistingAccess [CoreDefaultLinkToExistingAccess]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--CoreRequestFilesLinkEnabled [CoreRequestFilesLinkEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--OneDriveRequestFilesLinkEnabled [OneDriveRequestFilesLinkEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisableDocumentLibraryDefaultLabeling [DisableDocumentLibraryDefaultLabeling]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisableVivaConnectionsAnalytics [DisableVivaConnectionsAnalytics]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--DisplayNamesOfFileViewersInSpo [DisplayNamesOfFileViewersInSpo]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EnableAIPIntegration [EnableAIPIntegration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--EnableRestrictedAccessControl [EnableRestrictedAccessControl]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ExternalUserExpirationRequired [ExternalUserExpirationRequired]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--HideSyncButtonOnDocLib [HideSyncButtonOnDocLib]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IncludeAtAGlanceInShareEmails [IncludeAtAGlanceInShareEmails]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--InformationBarriersSuspension [InformationBarriersSuspension]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsFluidEnabled [IsFluidEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsWBFluidEnabled [IsWBFluidEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsCollabMeetingNotesFluidEnabled [IsCollabMeetingNotesFluidEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsEnableAppAuthPopUpEnabled [IsEnableAppAuthPopUpEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--IsLoopEnabled [IsLoopEnabled]',
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
    this.types.boolean.push(...SpoTenantSettingsSetCommand.booleanOptions);
  }

  public getAllEnumOptions(): string[] {
    return ['SharingCapability', 'CoreSharingCapability', 'ODBSharingCapability', 'OneDriveLoopSharingCapability', 'CoreLoopSharingCapability', 'ContainerSharingCapability', 'CoreDefaultShareLinkRole', 'CoreLoopDefaultSharingLinkRole', 'ContainerDefaultShareLinkRole', 'ContainerLoopDefaultShareLinkRole', 'OneDriveDefaultShareLinkRole', 'OneDriveLoopDefaultSharingLinkRole', 'CoreDefaultShareLinkScope', 'CoreLoopDefaultSharingLinkScope', 'ContainerDefaultShareLinkScope', 'ContainerLoopDefaultShareLinkScope', 'OneDriveDefaultShareLinkScope', 'OneDriveLoopDefaultSharingLinkScope', 'SharingDomainRestrictionMode', 'DefaultSharingLinkType', 'ODBMembersCanShare', 'ODBAccessRequests', 'AllowAnonymousMeetingParticipantsToAccessWhiteboards', 'FileAnonymousLinkType', 'FolderAnonymousLinkType', 'DefaultLinkPermission', 'ConditionalAccessPolicy', 'LimitedAccessFileType', 'MediaTranscription', 'MediaTranscriptionAutomaticFeatures', 'ImageTaggingOption', 'MarkNewFilesSensitiveByDefault', 'OCRModeForAdminSites', 'OCRModeForComplianceODBs', 'OCRModeForComplianceSites', 'SpecialCharactersStateInFileFolderNames'];
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
  private getMediaTranscriptionPolicyType(): string[] { return ['Enabled', 'Disabled']; }
  private getMediaTranscriptionAutomaticFeaturesPolicyType(): string[] { return ['Enabled', 'Disabled']; }
  private getImageTaggingChoice(): string[] { return ['Disabled', 'Basic', 'Enhanced']; }
  private getSensitiveByDefaultState(): string[] { return ['AllowExternalSharing', 'BlockExternalSharing']; }
  private getObjectCharacterRecognitionMode(): string[] { return ['Disabled', 'InclusionList', 'ExclusionList']; }
  private getSharingScope(): string[] { return ['Uninitialized', 'Anyone', 'Organization', 'SpecificPeople']; }
  private getRole(): string[] { return ['None', 'View', 'Edit', 'Owner', 'LimitedView', 'LimitedEdit', 'Review', 'RestrictedView', 'Submit', 'ManageList']; }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const tenantId: string = await spo.getTenantId(logger, this.debug);
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const formDigestValue = await spo.getRequestDigest(spoAdminUrl);

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

        if (['AllowedDomainListForSyncClient', 'DisabledWebPartIds', 'DisabledModernListTemplateIds'].indexOf(optionKey) > -1) {
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

      const res: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      if (args.options.EnableAzureADB2BIntegration === true) {
        await this.warn(logger, 'WARNING: Make sure to also enable the Azure AD one-time passcode authentication preview. If it is not enabled then SharePoint will not use Azure AD B2B even if EnableAzureADB2BIntegration is set to true. Learn more at http://aka.ms/spo-b2b-integration.');
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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
      case 'CoreSharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'ODBSharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'OneDriveLoopSharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'CoreLoopSharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'ContainerSharingCapability':
        return this.getSharingCapabilities().indexOf(value);
      case 'CoreDefaultShareLinkRole':
        return this.getRole().indexOf(value);
      case 'CoreLoopDefaultSharingLinkRole':
        return this.getRole().indexOf(value);
      case 'ContainerDefaultShareLinkRole':
        return this.getRole().indexOf(value);
      case 'ContainerLoopDefaultShareLinkRole':
        return this.getRole().indexOf(value);
      case 'OneDriveDefaultShareLinkRole':
        return this.getRole().indexOf(value);
      case 'OneDriveLoopDefaultSharingLinkRole':
        return this.getRole().indexOf(value);
      case 'CoreDefaultShareLinkScope':
        return this.getSharingScope().indexOf(value) - 1;
      case 'CoreLoopDefaultSharingLinkScope':
        return this.getSharingScope().indexOf(value) - 1;
      case 'ContainerDefaultShareLinkScope':
        return this.getSharingScope().indexOf(value) - 1;
      case 'ContainerLoopDefaultShareLinkScope':
        return this.getSharingScope().indexOf(value) - 1;
      case 'OneDriveDefaultShareLinkScope':
        return this.getSharingScope().indexOf(value) - 1;
      case 'OneDriveLoopDefaultSharingLinkScope':
        return this.getSharingScope().indexOf(value) - 1;
      case 'SharingDomainRestrictionMode':
        return this.getSharingDomainRestrictionModes().indexOf(value);
      case 'DefaultSharingLinkType':
        return this.getSharingLinkType().indexOf(value);
      case 'ODBMembersCanShare':
        return this.getSharingState().indexOf(value);
      case 'ODBAccessRequests':
        return this.getSharingState().indexOf(value);
      case 'AllowAnonymousMeetingParticipantsToAccessWhiteboards':
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
      case 'MediaTranscription':
        return this.getMediaTranscriptionPolicyType().indexOf(value);
      case 'MediaTranscriptionAutomaticFeatures':
        return this.getMediaTranscriptionAutomaticFeaturesPolicyType().indexOf(value);
      case 'ImageTaggingOption':
        return this.getImageTaggingChoice().indexOf(value);
      case 'MarkNewFilesSensitiveByDefault':
        return this.getSensitiveByDefaultState().indexOf(value);
      case 'OCRModeForAdminSites':
        return this.getObjectCharacterRecognitionMode().indexOf(value);
      case 'OCRModeForComplianceODBs':
        return this.getObjectCharacterRecognitionMode().indexOf(value);
      case 'OCRModeForComplianceSites':
        return this.getObjectCharacterRecognitionMode().indexOf(value);
      case 'SpecialCharactersStateInFileFolderNames':
        return this.getSpecialCharactersState().indexOf(value);
      default:
        return -1;
    }
  }
}

export default new SpoTenantSettingsSetCommand();