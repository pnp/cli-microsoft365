import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandError, CommandOption
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { SPOWebAppServicePrincipalPermissionGrant } from './SPOWebAppServicePrincipalPermissionGrant';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  resource: string;
  scope: string;
}

class SpoServicePrincipalGrantAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SERVICEPRINCIPAL_GRANT_ADD}`;
  }

  public get description(): string {
    return 'Grants the service principal permission to the specified API';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GRANT_ADD];
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
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
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.resource)}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.scope)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
          const result: SPOWebAppServicePrincipalPermissionGrant = json[json.length - 1];
          delete result._ObjectType_;
          cmd.log(result);

          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-r, --resource <resource>',
        description: 'The name of the resource for which permissions should be granted'
      },
      {
        option: '-s, --scope <scope>',
        description: 'The name of the permission that should be granted'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoServicePrincipalGrantAddCommand();