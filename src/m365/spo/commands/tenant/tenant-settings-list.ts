import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import request from '../../../../request';
import config from '../../../../config';
import commands from '../../commands';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { CommandInstance } from '../../../../cli';

class SpoTenantSettingsListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the global tenant settings';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties><Property Name="HideDefaultThemes" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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

        let result = json[4];
        delete result['_ObjectIdentity_'];
        delete result['_ObjectType_'];

        // map integers to their enums
        const sharingLinkType = ['None', 'Direct', 'Internal', 'AnonymousAccess'];
        const sharingCapabilities = ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly'];
        const sharingDomainRestrictionModes = ['None', 'AllowList', 'BlockList'];
        const sharingState = ['Unspecified', 'On', 'Off'];
        const anonymousLinkType = ['None', 'View', 'Edit'];
        const sharingPermissionType = ['None', 'View', 'Edit'];
        const sPOConditionalAccessPolicyType = ['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess'];
        const specialCharactersState = ['NoPreference', 'Allowed', 'Disallowed'];
        const sPOLimitedAccessFileType = ['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles'];

        result['SharingCapability'] = sharingCapabilities[result['SharingCapability']];
        result['SharingDomainRestrictionMode'] = sharingDomainRestrictionModes[result['SharingDomainRestrictionMode']];
        result['ODBMembersCanShare'] = sharingState[result['ODBMembersCanShare']];
        result['ODBAccessRequests'] = sharingState[result['ODBAccessRequests']];
        result['DefaultSharingLinkType'] = sharingLinkType[result['DefaultSharingLinkType']];
        result['FileAnonymousLinkType'] = anonymousLinkType[result['FileAnonymousLinkType']];
        result['FolderAnonymousLinkType'] = anonymousLinkType[result['FolderAnonymousLinkType']];
        result['DefaultLinkPermission'] = sharingPermissionType[result['DefaultLinkPermission']];
        result['ConditionalAccessPolicy'] = sPOConditionalAccessPolicyType[result['ConditionalAccessPolicy']];
        result['SpecialCharactersStateInFileFolderNames'] = specialCharactersState[result['SpecialCharactersStateInFileFolderNames']];
        result['LimitedAccessFileType'] = sPOLimitedAccessFileType[result['LimitedAccessFileType']];

        cmd.log(result);

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }
}

module.exports = new SpoTenantSettingsListCommand();