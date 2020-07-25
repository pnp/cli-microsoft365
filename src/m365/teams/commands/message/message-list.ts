import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';
import { Message } from '../../Message';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  since?: string;
}

class TeamsMessageListCommand extends GraphItemsListCommand<Message> {
  public get name(): string {
    return `${commands.TEAMS_MESSAGE_LIST}`;
  }

  public get description(): string {
    return 'Lists all messages from a channel in a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const deltaExtension: string = args.options.since !== undefined ? `/delta?$filter=lastModifiedDateTime gt ${args.options.since}` : '';
    const endpoint: string = `${this.resource}/beta/teams/${args.options.teamId}/channels/${args.options.channelId}/messages${deltaExtension}`;

    this
    .getAllItems(endpoint, cmd, true)
    .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(m => {
            return {
              id: m.id,
              summary: m.summary,
              body: m.body.content
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
        description: 'The ID of the team where the channel is located'
      },
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the channel for which to list messages'
      },
      {
        option: '-s, --since [since]',
        description: 'Date (ISO standard, dash separator) to get delta of messages from (in last 8 months)'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
      }

      if (args.options.since && !Utils.isValidISODateDashOnly(args.options.since as string)) {
        return `${args.options.since} is not a valid ISO Date (with dash separator)`;
      }
      
      if (args.options.since && !Utils.isDateInRange(args.options.since as string, 8)) {
        return `${args.options.since} is not in the last 8 months (for delta messages)`;
      }

      return true;
    };
  }
}

module.exports = new TeamsMessageListCommand();