import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandError, CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
}

class SpoServicePrincipalGrantRevokeCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SERVICEPRINCIPAL_GRANT_REVOKE}`;
  }

  public get description(): string {
    return 'Revokes the specified set of permissions granted to the service principal';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GRANT_REVOKE];
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<string> => {
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
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.grantId)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-i, --grantId <grantId>',
      description: 'ObjectId of the permission grant to revoke'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.grantId) {
        return 'Required parameter grantId missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SERVICEPRINCIPAL_GRANT_REVOKE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site using the
      ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To revoke permissions granted to the service principal, you have to first connect to a SharePoint tenant admin
    site using the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.

    The permission grant you want to revoke is denoted using its ${chalk.grey('ObjectId')}. You can retrieve
    it using the ${chalk.grey(`${commands.SERVICEPRINCIPAL_GRANT_LIST}`)} command.

  Examples:
  
    Revoke permission grant with ObjectId ${chalk.grey('50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA')}
      ${chalk.grey(config.delimiter)} ${commands.SERVICEPRINCIPAL_GRANT_REVOKE} --grantId 50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA
`);
  }
}

module.exports = new SpoServicePrincipalGrantRevokeCommand();