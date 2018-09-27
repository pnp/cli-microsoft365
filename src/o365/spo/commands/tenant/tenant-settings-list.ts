import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import * as request from 'request-promise-native';
import config from '../../../../config';
import commands from '../../commands';
import Utils from '../../../../Utils';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoTenantSettingsListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the global tenant settings';
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties><Property Name="HideDefaultThemes" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online
    tenant admin site, using the ${chalk.blue(commands.LOGIN)} command.

  Examples:
  
    Lists the settings of the tenant
      ${chalk.grey(config.delimiter)} ${commands.TENANT_SETTINGS_LIST}
  ` );
  }
}

module.exports = new SpoTenantSettingsListCommand();