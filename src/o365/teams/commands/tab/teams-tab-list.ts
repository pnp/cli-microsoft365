import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { Tab } from '../../Tab';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!args.options.channelId) {
        return 'Required parameter channelId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
      } 

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can only retrieve tabs for teams of which you are a member.

    Tabs 'Conversations' and 'Files' are present in every team and therefore not
    included in the list of available tabs.

  Examples:
  
    List all tabs in a Microsoft Teams channel
      ${this.name} --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype

    Include all the values from the tab configuration and associated teams app
      ${this.name} --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --output json
`);
  }
}

module.exports = new TeamsTabListCommand();