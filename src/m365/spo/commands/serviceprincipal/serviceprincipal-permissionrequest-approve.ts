import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandError, CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  requestId: string;
}

class SpoServicePrincipalPermissionRequestApproveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE}`;
  }

  public get description(): string {
    return 'Approves the specified permission request';
  }

  public alias(): string[] | undefined {
    return [commands.SP_PERMISSIONREQUEST_APPROVE];
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
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${Utils.escapeXml(args.options.requestId)}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`
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
          const output: any = json[json.length - 1];
          delete output._ObjectType_;

          cmd.log(output);

          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-i, --requestId <requestId>',
      description: 'ID of the permission request to approve'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.requestId)) {
        return `${args.options.requestId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoServicePrincipalPermissionRequestApproveCommand();