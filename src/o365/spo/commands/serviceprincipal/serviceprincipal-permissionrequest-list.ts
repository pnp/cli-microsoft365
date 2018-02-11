import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { SPOWebAppServicePrincipalPermissionRequest } from './SPOWebAppServicePrincipalPermissionRequest';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoServicePrincipalPermissionRequestListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST}`;
  }

  public get description(): string {
    return 'Lists pending permission requests';
  }

  public alias(): string[] | undefined {
    return [commands.SP_PERMISSIONREQUEST_LIST];
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`
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
          cmd.log(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: SPOWebAppServicePrincipalPermissionRequest[] = json[json.length - 1]._Child_Items_;
          cmd.log(result.map(r => {
            return {
              Id: r.Id.replace('/Guid(', '').replace(')/', ''),
              Resource: r.Resource,
              ResourceId: r.ResourceId,
              Scope: r.Scope
            };
          }));

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site using the
      ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To list pending permission requests, you have to first connect to a SharePoint tenant admin
    site using the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.

  Examples:
  
    List all pending permission requests
      ${chalk.grey(config.delimiter)} ${commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST}
`);
  }
}

module.exports = new SpoServicePrincipalPermissionRequestListCommand();