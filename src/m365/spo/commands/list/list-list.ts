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

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url', 'Id'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all lists in site at ${args.options.webUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/lists?$expand=RootFolder`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ListInstanceCollection>(requestOptions)
      .then((listInstances: ListInstanceCollection): void => {
        listInstances.value.forEach(l => {
          l.Url = l.RootFolder.ServerRelativeUrl;
        });

        logger.log(listInstances.value);
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