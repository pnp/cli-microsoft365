import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';
import { Reply } from '../../Reply';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  messageId: string;
}

class TeamsMessageReplyListCommand extends GraphItemsListCommand<Reply>  {
  public get name(): string {
    return `${commands.TEAMS_MESSAGE_REPLY_LIST}`;
  }

  public get description(): string {
    return 'Retrieves replies to a message from a channel in a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/beta/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}/replies`;

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
        description: 'The ID of the channel that contains the message'
      },
      {
        option: '-m, --messageId <messageId>',
        description: 'The ID of the message to retrieve replies for'
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

      if (!args.options.messageId) {
        return 'Required parameter messageId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
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

    You can only retrieve replies to a message from a Microsoft Teams team
    if you are a member of that team.

  Examples:
  
    Retrieve the replies from a specified message from a channel of the
    Microsoft Teams team
      ${this.name} --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --messageId 1540747442203
`   );
  }
}

module.exports = new TeamsMessageReplyListCommand();