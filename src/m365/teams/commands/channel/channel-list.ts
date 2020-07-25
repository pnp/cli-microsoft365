import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';
import { Channel } from '../../Channel';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels`;
    
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
              displayName: m.displayName
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
        description: 'The ID of the team to list the channels of'
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
      
      return true;
    };
  }
}

module.exports = new TeamsChannelListCommand();