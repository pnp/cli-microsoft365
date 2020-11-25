import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Tab } from '../../Tab';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
}

class TeamsTabListCommand extends GraphItemsListCommand<Tab> {
  public get name(): string {
    return `${commands.TEAMS_TAB_LIST}`;
  }

  public get description(): string {
    return 'Lists tabs in the specified Microsoft Teams channel';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'teamsAppTabId'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${encodeURIComponent(args.options.channelId)}/tabs?$expand=teamsApp`;

    this
      .getAllItems(endpoint, logger, true)
      .then((): void => {
        this.items.forEach(i => {
          (i as any).teamsAppTabId = i.teamsApp.id;
        });

        logger.log(this.items);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team where the channel is located'
      },
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the channel for which to list tabs'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.teamId as string)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
      return `${args.options.channelId} is not a valid Teams ChannelId`;
    }

    return true;
  }
}

module.exports = new TeamsTabListCommand();