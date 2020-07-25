import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  principals: string;
  rights: string;
}

class SpoHubSiteRightsGrantCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_RIGHTS_GRANT}`;
  }

  public get description(): string {
    return 'Grants permissions to join the hub site for one or more principals';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Granting permissions to join the hub site ${args.options.url} to principals ${args.options.principals}...`);
        }

        const principals: string = args.options.principals
          .split(',')
          .map(p => `<Object Type="String">${Utils.escapeXml(p.trim())}</Object>`)
          .join('');
        const grantedRights: string = '1';

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter><Parameter Type="Array">${principals}</Parameter><Parameter Type="Enum">${grantedRights}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the hub site to grant rights on'
      },
      {
        option: '-p, --principals <principals>',
        description: 'Comma-separated list of principals to grant join rights. Principals can be users or mail-enabled security groups in the form of "alias" or "alias@<domain name>.com"'
      },
      {
        option: '-r, --rights <rights>',
        description: 'Rights to grant to principals. Available values Join',
        autocomplete: ['Join']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.rights !== 'Join') {
        return `${args.options.rights} is not a valid rights value. Allowed values Join`;
      }

      return true;
    };
  }
}

module.exports = new SpoHubSiteRightsGrantCommand();
