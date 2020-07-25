import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoHubSiteThemeSyncCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_THEME_SYNC}`;
  }

  public get description(): string {
    return 'Applies any theme updates from the parent hub site.';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log('Syncing hub site theme...');
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/SyncHubSiteTheme`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((): void => {
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
        description: 'URL of the site to apply theme updates from the hub site to'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    }
  }
}

module.exports = new SpoHubSiteThemeSyncCommand();