import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class TeamsChannelListCommand extends GraphItemsListCommand<Channel>{
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_LIST}`;
  }

  public get description(): string {
    return 'Lists channels in the specified Microsoft Teams team';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels`;

    this
      .getAllItems(endpoint, logger, true)
      .then((): void => {
        if (args.options.output === 'json') {
          logger.log(this.items);
        }
        else {
          logger.log(this.items.map(m => {
            return {
              id: m.id,
              displayName: m.displayName
            }
          }));
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team to list the channels of'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.teamId as string)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsChannelListCommand();