import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  MinCompatibilityLevel: z.string().refine(val => !isNaN(Number(val)), { message: 'MinCompatibilityLevel is not a number' }).optional(),
  MaxCompatibilityLevel: z.string().refine(val => !isNaN(Number(val)), { message: 'MaxCompatibilityLevel is not a number' }).optional(),
  ExternalServicesEnabled: z.boolean().optional(),
  NoAccessRedirectUrl: z.string().optional(),
  SharingCapability: z.enum(['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']).optional(),
  DisplayStartASiteOption: z.boolean().optional(),
  StartASiteFormUrl: z.string().optional(),
  ShowEveryoneClaim: z.boolean().optional(),
  ShowAllUsersClaim: z.boolean().optional(),
  ShowEveryoneExceptExternalUsersClaim: z.boolean().optional(),
  SearchResolveExactEmailOrUPN: z.boolean().optional(),
  OfficeClientADALDisabled: z.boolean().optional(),
  LegacyAuthProtocolsEnabled: z.boolean().optional(),
  RequireAcceptingAccountMatchInvitedAccount: z.boolean().optional(),
  ProvisionSharedWithEveryoneFolder: z.boolean().optional(),
  SignInAccelerationDomain: z.string().optional(),
  EnableGuestSignInAcceleration: z.boolean().optional(),
  UsePersistentCookiesForExplorerView: z.boolean().optional(),
  BccExternalSharingInvitations: z.boolean().optional(),
  BccExternalSharingInvitationsList: z.string().optional(),
  UserVoiceForFeedbackEnabled: z.boolean().optional(),
  PublicCdnEnabled: z.boolean().optional(),
  PublicCdnAllowedFileTypes: z.string().optional(),
  RequireAnonymousLinksExpireInDays: z.string().refine(val => !isNaN(Number(val)), { message: 'RequireAnonymousLinksExpireInDays is not a number' }).optional(),
  SharingAllowedDomainList: z.string().optional(),
  SharingBlockedDomainList: z.string().optional(),
  SharingDomainRestrictionMode: z.enum(['None', 'AllowList', 'BlockList']).optional(),
  OneDriveStorageQuota: z.string().refine(val => !isNaN(Number(val)), { message: 'OneDriveStorageQuota is not a number' }).optional(),
  OneDriveForGuestsEnabled: z.boolean().optional(),
  IPAddressEnforcement: z.boolean().optional(),
  IPAddressAllowList: z.string().optional(),
  IPAddressWACTokenLifetime: z.string().refine(val => !isNaN(Number(val)), { message: 'IPAddressWACTokenLifetime is not a number' }).optional(),
  UseFindPeopleInPeoplePicker: z.boolean().optional(),
  DefaultSharingLinkType: z.enum(['None', 'Direct', 'Internal', 'AnonymousAccess']).optional(),
  ODBMembersCanShare: z.enum(['Unspecified', 'On', 'Off']).optional(),
  ODBAccessRequests: z.enum(['Unspecified', 'On', 'Off']).optional(),
  PreventExternalUsersFromResharing: z.boolean().optional(),
  ShowPeoplePickerSuggestionsForGuestUsers: z.boolean().optional(),
  FileAnonymousLinkType: z.enum(['None', 'View', 'Edit']).optional(),
  FolderAnonymousLinkType: z.enum(['None', 'View', 'Edit']).optional(),
  NotifyOwnersWhenItemsReshared: z.boolean().optional(),
  NotifyOwnersWhenInvitationsAccepted: z.boolean().optional(),
  NotificationsInOneDriveForBusinessEnabled: z.boolean().optional(),
  NotificationsInSharePointEnabled: z.boolean().optional(),
  OwnerAnonymousNotification: z.boolean().optional(),
  CommentsOnSitePagesDisabled: z.boolean().optional(),
  SocialBarOnSitePagesDisabled: z.boolean().optional(),
  OrphanedPersonalSitesRetentionPeriod: z.string().refine(val => !isNaN(Number(val)), { message: 'OrphanedPersonalSitesRetentionPeriod is not a number' }).optional(),
  DisallowInfectedFileDownload: z.boolean().optional(),
  DefaultLinkPermission: z.enum(['None', 'View', 'Edit']).optional(),
  ConditionalAccessPolicy: z.enum(['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess']).optional(),
  AllowDownloadingNonWebViewableFiles: z.boolean().optional(),
  AllowEditing: z.boolean().optional(),
  ApplyAppEnforcedRestrictionsToAdHocRecipients: z.boolean().optional(),
  FilePickerExternalImageSearchEnabled: z.boolean().optional(),
  EmailAttestationRequired: z.boolean().optional(),
  EmailAttestationReAuthDays: z.string().refine(val => !isNaN(Number(val)), { message: 'EmailAttestationReAuthDays is not a number' }).optional(),
  HideDefaultThemes: z.boolean().optional(),
  BlockAccessOnUnmanagedDevices: z.boolean().optional(),
  AllowLimitedAccessOnUnmanagedDevices: z.boolean().optional(),
  BlockDownloadOfAllFilesForGuests: z.boolean().optional(),
  BlockDownloadOfAllFilesOnUnmanagedDevices: z.boolean().optional(),
  BlockDownloadOfViewableFilesForGuests: z.boolean().optional(),
  BlockDownloadOfViewableFilesOnUnmanagedDevices: z.boolean().optional(),
  BlockMacSync: z.boolean().optional(),
  DisableReportProblemDialog: z.boolean().optional(),
  DisplayNamesOfFileViewers: z.boolean().optional(),
  EnableMinimumVersionRequirement: z.boolean().optional(),
  HideSyncButtonOnODB: z.boolean().optional(),
  IsUnmanagedSyncClientForTenantRestricted: z.boolean().optional(),
  LimitedAccessFileType: z.enum(['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles']).optional(),
  OptOutOfGrooveBlock: z.boolean().optional(),
  OptOutOfGrooveSoftBlock: z.boolean().optional(),
  OrgNewsSiteUrl: z.string().optional(),
  PermissiveBrowserFileHandlingOverride: z.boolean().optional(),
  ShowNGSCDialogForSyncOnODB: z.boolean().optional(),
  SpecialCharactersStateInFileFolderNames: z.enum(['NoPreference', 'Allowed', 'Disallowed']).optional(),
  SyncPrivacyProfileProperties: z.boolean().optional(),
  ExcludedFileExtensionsForSyncClient: z.string().optional(),
  AllowedDomainListForSyncClient: z.string().optional(),
  DisabledWebPartIds: z.string().optional(),
  DisableCustomAppAuthentication: z.boolean().optional(),
  CommentsOnListItemsDisabled: z.boolean().optional(),
  EnableAzureADB2BIntegration: z.boolean().optional(),
  SyncAadB2BManagementPolicy: z.boolean().optional(),
  AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantSettingsSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets tenant global settings';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    const excluded = ['output', 'o', 'debug', 'verbose', '_', 'query'];
    return schema.refine(opts => Object.keys(opts).some(key => !excluded.includes(key)), {
      error: 'You must specify at least one option'
    });
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const tenantId: string = await spo.getTenantId(logger, this.debug);
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const formDigestValue = await spo.getRequestDigest(spoAdminUrl);

      // map the args.options to XML Properties
      let propsXml: string = '';
      let id: number = 42; // geek's humor
      const optionsRecord = args.options as Record<string, any>;
      for (const optionKey of Object.keys(optionsRecord)) {
        if (this.isExcludedOption(optionKey)) {
          continue;
        }

        let optionValue = optionsRecord[optionKey];
        if (this.getAllEnumOptions().indexOf(optionKey) > -1) {
          // map enum values to int
          optionValue = this.mapEnumToInt(optionKey, optionsRecord[optionKey]);
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

      const res: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      if (args.options.EnableAzureADB2BIntegration === true) {
        await this.warn(logger, 'WARNING: Make sure to also enable the Microsoft Entra one-time passcode authentication preview. If it is not enabled then SharePoint will not use Microsoft Entra B2B even if EnableAzureADB2BIntegration is set to true. Learn more at http://aka.ms/spo-b2b-integration.');
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

export default new SpoTenantSettingsSetCommand();