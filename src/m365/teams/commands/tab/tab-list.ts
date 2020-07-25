import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { Tab } from '../../Tab';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${encodeURIComponent(args.options.channelId)}/tabs?$expand=teamsApp`;

    this
      .getAllItems(endpoint, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map((t: Tab) => {
            return {
              id: t.id,
              displayName: t.displayName,
              teamsAppTabId: t.teamsApp.id,
            }
          }));
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

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
      } 

      return true;
    };
  }
}

module.exports = new TeamsTabListCommand();