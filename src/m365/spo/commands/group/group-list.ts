import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { GroupPropertiesCollection } from "./GroupPropertiesCollection";
import { GroupProperties } from "./GroupProperties";
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoGroupListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all the groups within specific web';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving list of groups for specified web at ${args.options.webUrl}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/sitegroups`;

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    }

    request
      .get<GroupPropertiesCollection>(requestOptions)
      .then((groupProperties: GroupPropertiesCollection): void => {
        if (args.options.output === 'json') {
          cmd.log(groupProperties);
        }
        else {
          cmd.log(groupProperties.value.map((g: GroupProperties) => {
            return {
              Id: g.Id,
              Title: g.Title,
              LoginName: g.LoginName,
              IsHiddenInUI: g.IsHiddenInUI,
              PrincipalType: g.PrincipalType
            };
          }))
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Url of the web to list the group within'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Lists all the groups within specific web
      m365 ${this.name} --webUrl "https://contoso.sharepoint.com/sites/contoso"  
    `);
  }
}

module.exports = new SpoGroupListCommand();