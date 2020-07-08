import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ListInstanceCollection } from "./ListInstanceCollection";

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class ListListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Lists all available list in the specified site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving all lists in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string;

    if (args.options.output === 'json') {
      requestUrl = `${args.options.webUrl}/_api/web/lists?$expand=RootFolder`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists?$expand=RootFolder&$select=Title,Id,RootFolder/ServerRelativeURL`;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<ListInstanceCollection>(requestOptions)
      .then((listInstances: ListInstanceCollection): void => {
        if (args.options.output === 'json') {
          if (listInstances.value) {
            cmd.log(listInstances.value);
          }
        }
        else {
          cmd.log(listInstances.value.map(l => {
            return {
              Title: l.Title,
              Url: l.RootFolder.ServerRelativeUrl,
              Id: l.Id
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the lists to retrieve are located'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Return all lists located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x
      `);
  }
}

module.exports = new ListListCommand();