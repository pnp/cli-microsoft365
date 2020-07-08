import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  confirm?: boolean;
}

class TeamsRemoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TEAM_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeTeam: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/groups/${encodeURIComponent(args.options.teamId)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeTeam();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the team ${args.options.teamId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeTeam();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Teams team to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the specified team'
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

    When deleted, Microsoft 365 groups are moved to a temporary container and
    can be restored within 30 days. After that time, they are permanently
    deleted. This applies only to Microsoft 365 groups.

  Examples:
  
    Removes the specified team 
      ${this.name} --teamId '00000000-0000-0000-0000-000000000000'

    Removes the specified team without confirmation
      ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --confirm

  More information:

    directory resource type (deleted items)
      https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0
  `);
  }
}

module.exports = new TeamsRemoveCommand();