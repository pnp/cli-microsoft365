import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import GraphCommand from "../../../base/GraphCommand";
import { Team } from '../../Team';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class TeamsMessagingSettingsListCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_MESSAGINGSETTINGS_LIST}`;
  }

  public get description(): string {
    return 'Lists messaging settings for a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}?$select=messagingSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get<Team>(requestOptions)
      .then((res: Team): void => {
        cmd.log(res.messagingSettings);

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
        description: 'The ID of the team for which to get the messaging settings'
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

module.exports = new TeamsMessagingSettingsListCommand();