import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  confirm?: boolean;
}

class TeamsTeamRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeTeam: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/groups/${encodeURIComponent(args.options.teamId)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .delete(requestOptions)
        .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeTeam();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the team ${args.options.teamId}?`
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
        option: '-i, --teamId <teamId>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsTeamRemoveCommand();