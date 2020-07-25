import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import request from '../../../../request';
import { Channel } from '../../Channel';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  channelId: string;
  teamId: string;
}

class TeamsChannelGetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specific Microsoft Teams team channel';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    }

    request
      .get<Channel>(requestOptions)
      .then((res: Channel): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team to which the channel belongs'
      },
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the channel for which to retrieve more information'
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

      return true;
    };
  }
}

module.exports = new TeamsChannelGetCommand();