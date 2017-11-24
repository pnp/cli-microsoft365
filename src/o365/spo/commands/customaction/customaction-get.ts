import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import VerboseOption from '../../../../VerboseOption';
import {
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  id: string;
  url: string;
  scope?: string;
}

class SpoCustomActionGetCommand extends SpoCommand {

  public get name(): string {
    return `${commands.CUSTOMACTION_GET}`;
  }

  public get description(): string {
    return 'Gets details for the specified custom action';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving access token for ${auth.service.resource}...`);
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.verbose)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (this.verbose) {
          cmd.log(`Retrieved access token ${accessToken}. Loading details for the ${args.options.id} custom action...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/usercustomactions('${encodeURIComponent(args.options.id)}')`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        if (this.verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: any): void => {  // ContextInfo
        if (this.verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Retrieving custom action...`);

        const json: any = JSON.parse(res); // RestResponse<AppMetadata>
        cmd.log(json);

        const apps = json.d.results;   

        cmd.log(apps.toString());
        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'Id (Guid) of the custom action to retrieve'
      },
      {
        option: '-u, --url <url>',
        description: 'Url of the site (collection) to retrieve the custom action from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      }
  ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      if (Utils.isValidGuid(args.options.id) === false) {
          return `${args.options.id} is not valid. Custom action id (Guid) expected.`;
      }

      if (SpoCommand.isValidSharePointUrl(args.options.url) === false) {
        return 'Missing required option url';
      }

      if (args.options.scope) {
        if (args.options.scope !== 'Site' &&
          args.options.scope !== 'Web' &&
          args.options.scope !== 'All') {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }
      }
      
      return true;
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.CUSTOMACTION_GET).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
        using the ${chalk.blue(commands.CONNECT)} command.
                      
        Remarks:
      
          To retrieve custom action, you have to first connect to a SharePoint Online site using the
          ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
      
        Examples:
        
          ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_GET} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test
      
          ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_GET} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test -s "Site"
      
        More information:
      
          UserCustomAction REST API resources:
            https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction
      `);
    };
  }
}

module.exports = new SpoCustomActionGetCommand();