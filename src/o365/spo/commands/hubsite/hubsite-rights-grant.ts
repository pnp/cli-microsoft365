import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      cmd.log(`Retrieving access token for ${auth.service.resource}...`);
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(accessToken);
          cmd.log('');
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Granting permissions to join the hub site ${args.options.url} to principals ${args.options.principals}...`);
        }

        const principals: string = args.options.principals
          .split(',')
          .map(p => `<Object Type="String">${Utils.escapeXml(p.trim())}</Object>`)
          .join('');
        const grantedRights: string = '1';

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter><Parameter Type="Array">${principals}</Parameter><Parameter Type="Enum">${grantedRights}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
        else {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.principals) {
        return 'Required parameter principals missing';
      }

      if (!args.options.rights) {
        return 'Required parameter rights missing';
      }

      if (args.options.rights !== 'Join') {
        return `${args.options.rights} is not a valid rights value. Allowed values Join`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online tenant admin
    site using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    ${chalk.yellow('Attention:')} This command is based on a SharePoint API that is currently
    in preview and is subject to change once the API reached general
    availability.

    To grant permissions to join the hub site, you have to first log in
    to a tenant admin site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso-admin.sharepoint.com`)}.
    If you are logged in to a different site and will try to grant permissions
    to join the hub site, you will get an error.

  Examples:
  
    Grant user with alias ${chalk.grey('PattiF')} permission to join sites to the hub site with
    URL ${chalk.grey('https://contoso.sharepoint.com/sites/sales')}
      ${chalk.grey(config.delimiter)} ${this.name} --url https://contoso.sharepoint.com/sites/sales --principals PattiF --rights Join

    Grant users with aliases ${chalk.grey('PattiF')} and ${chalk.grey('AdeleV')} permission to join sites
    to the hub site with URL ${chalk.grey('https://contoso.sharepoint.com/sites/sales')}
      ${chalk.grey(config.delimiter)} ${this.name} --url https://contoso.sharepoint.com/sites/sales --principals PattiF,AdeleV --rights Join

    Grant user with email ${chalk.grey('PattiF@contoso.com')} permission to join sites
    to the hub site with URL ${chalk.grey('https://contoso.sharepoint.com/sites/sales')}
      ${chalk.grey(config.delimiter)} ${this.name} --url https://contoso.sharepoint.com/sites/sales --principals PattiF@contoso.com --rights Join

  More information:

    SharePoint hub sites new in Office 365
      https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547
`);
  }
}

module.exports = new SpoHubSiteRightsGrantCommand();