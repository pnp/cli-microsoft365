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

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  requestId: string;
}

class SpoServicePrincipalPermissionRequestDenyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY}`;
  }

  public alias(): string[] | undefined {
    return [commands.SP_PERMISSIONREQUEST_DENY];
  }

  public get description(): string {
    return 'Denies the specified permission request';
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
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="160" ObjectPathId="159" /><ObjectPath Id="162" ObjectPathId="161" /><ObjectPath Id="164" ObjectPathId="163" /><Method Name="Deny" Id="165" ObjectPathId="163" /></Actions><ObjectPaths><Constructor Id="159" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="161" ParentId="159" Name="PermissionRequests" /><Method Id="163" ParentId="161" Name="GetById"><Parameters><Parameter Type="Guid">{${Utils.escapeXml(args.options.requestId)}}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-i, --requestId <requestId>',
      description: 'ID of the permission request to deny'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.requestId) {
        return 'Required parameter requestId missing';
      }

      if (!Utils.isValidGuid(args.options.requestId)) {
        return `${args.options.requestId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
        
  Remarks:

    The permission request you want to deny is denoted using its ${chalk.grey('ID')}. You can
    retrieve it using the ${chalk.grey(`${commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST}`)} command.

  Examples:
  
    Deny permission request with id ${chalk.grey('4dc4c043-25ee-40f2-81d3-b3bf63da7538')}
      ${commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY} --requestId 4dc4c043-25ee-40f2-81d3-b3bf63da7538
`);
  }
}

module.exports = new SpoServicePrincipalPermissionRequestDenyCommand();