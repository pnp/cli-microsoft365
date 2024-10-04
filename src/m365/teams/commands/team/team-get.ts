import { Team } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class TeamsTeamGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_GET;
  }

  public get description(): string {
    return 'Retrieve information about the specified Microsoft Team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.id) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.id)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        const res = await request.get<Team>(requestOptions);
        await logger.log(res);
      }
      else {
        const team = await teams.getTeamByDisplayName(args.options.name!);
        await logger.log(team);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsTeamGetCommand();