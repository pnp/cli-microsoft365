import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class SpoThemeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_GET;
  }

  public get description(): string {
    return 'Gets custom theme information';
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Getting ${args.options.name} theme from tenant...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
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
          const { _ObjectType_, ...theme } = json[6];
          cmd.log(theme);
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-n, --name <name>',
      description: 'The name of the theme to retrieve'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To get information about a theme, you have to first connect to a tenant
    admin site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
        
  Examples:
  
    Get information about a theme
      ${chalk.grey(config.delimiter)} ${commands.THEME_GET} --name Contoso-Blue

  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview
    `);
  }
}

module.exports = new SpoThemeGetCommand();