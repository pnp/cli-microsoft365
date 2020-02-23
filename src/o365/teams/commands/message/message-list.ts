import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';
import { Message } from '../../Message';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!args.options.channelId) {
        return 'Required parameter channelId missing';
      }

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
  
  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

    You can list all the messages from a Microsoft Teams team if you are
    a member of that team.

  Examples:
  
    Lists all the messages from a channel of the Microsoft Teams team
      ${this.name} --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype

    List the messages from a channel of the Microsoft Teams team that have been
    created or modified since the date specified by the ${chalk.blue('--since')} parameter
    (WARNING: only captures the last 8 months of data) 
      ${this.name} --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype --since 2019-12-31T14:00:00Z
`   );
  }
}

module.exports = new TeamsMessageListCommand();