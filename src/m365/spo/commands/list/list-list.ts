import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstanceCollection } from "./ListInstanceCollection";

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Retrieving all lists in site at ${args.options.webUrl}...`);
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
            logger.log(listInstances.value);
          }
        }
        else {
          logger.log(listInstances.value.map(l => {
            return {
              Title: l.Title,
              Url: l.RootFolder.ServerRelativeUrl,
              Id: l.Id
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new ListListCommand();