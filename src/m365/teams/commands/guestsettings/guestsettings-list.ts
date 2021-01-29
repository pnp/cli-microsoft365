import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { Team } from '../../Team';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class TeamsGuestSettingsListCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_GUESTSETTINGS_LIST}`;
  }

  public get description(): string {
    return 'Lists guest settings for a Microsoft Teams team';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}?$select=guestSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<Team>(requestOptions)
      .then((res: Team): void => {
        logger.log(res.guestSettings);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsGuestSettingsListCommand();