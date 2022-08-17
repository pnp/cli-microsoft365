import { Group } from '@microsoft/microsoft-graph-types';
import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  teamId?: string;
  confirm?: boolean;
}

class TeamsTeamRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.id && !args.options.name && !args.options.teamId) {
	      return 'Specify either id or name';
	    }

	    if (args.options.name && (args.options.id || args.options.teamId)) {
	      return 'Specify either id or name but not both';
	    }

	    if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
	      return `${args.options.teamId} is not a valid GUID`;
	    }

	    if (args.options.id && !validation.isValidGuid(args.options.id)) {
	      return `${args.options.id} is not a valid GUID`;
	    }

	    return true;
      }
    );
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.name!)
      .then(group => {
        if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        return group.id!;
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.teamId) {
      args.options.id = args.options.teamId;

      this.warn(logger, `Option 'teamId' is deprecated. Please use 'id' instead.`);
    }

    const removeTeam: () => void = (): void => {
      this
        .getTeamId(args)
        .then((teamId: string): Promise<void> => {
          const requestOptions: any = {
            url: `${this.resource}/v1.0/groups/${encodeURIComponent(teamId)}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          return request.delete(requestOptions);
        })
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
}

module.exports = new TeamsTeamRemoveCommand();