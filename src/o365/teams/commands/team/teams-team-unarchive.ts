import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class TeamsUnarchiveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TEAM_UNARCHIVE}`;
  }

  public get description(): string {
    return 'Restores an archived Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0`;

    const requestOptions: any = {
      url: `${endpoint}/teams/${encodeURIComponent(args.options.teamId)}/unarchive`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team to restore'
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

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
          
    This command supports admin permissions. Global admins and Microsoft Teams
    service admins can restore teams that they are not a member of.

    This command restores users' ability to send messages and edit the team,
    abiding by tenant and team settings.

  Examples:
    
    Restore an archived Microsoft Teams team
      ${commands.TEAMS_TEAM_UNARCHIVE} --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
    `);
  }
}

module.exports = new TeamsUnarchiveCommand();