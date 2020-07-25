import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoPageListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_LIST}`;
  }

  public get description(): string {
    return 'Lists all modern pages in the given site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving client-side pages...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          const clientSidePages: any[] = res.value.filter(p => p.ListItemAllFields.ClientSideApplicationId === 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec');
          if (args.options.output === 'json') {
            cmd.log(clientSidePages);
          }
          else {
            cmd.log(clientSidePages.map(p => {
              return {
                Name: p.Name,
                Title: p.Title
              }
            }));
          }
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site from which to retrieve available pages'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoPageListCommand();