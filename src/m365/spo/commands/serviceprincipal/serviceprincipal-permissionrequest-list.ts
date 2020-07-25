import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { SPOWebAppServicePrincipalPermissionRequest } from './SPOWebAppServicePrincipalPermissionRequest';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: {}, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`
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
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }
}

module.exports = new SpoServicePrincipalPermissionRequestListCommand();