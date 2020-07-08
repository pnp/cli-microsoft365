import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import request from '../../../../request';
import { Channel } from '../../Channel';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!args.options.channelId) {
        return 'Required parameter channelId missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get information about Microsoft Teams team channel with id
    ${chalk.grey('19:493665404ebd4a18adb8a980a31b4986@thread.skype')}
      ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
    `);
  }
}

module.exports = new TeamsChannelGetCommand();