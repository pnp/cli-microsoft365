import { Channel } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  id?: string;
  name?: string;
  primary?: boolean;
}

class TeamsChannelGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_GET;
  }

  public get description(): string {
    return 'Gets information about the specific Microsoft Teams team channel';
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
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        primary: !!args.options.primary
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '--primary'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.id && !validation.isValidTeamsChannelId(args.options.id)) {
          return `${args.options.id} is not a valid Teams channel id`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['teamId', 'teamName'] },
      { options: ['id', 'name', 'primary'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId = args.options.teamId || await teams.getTeamIdByDisplayName(args.options.teamName!);

      if (args.options.primary || args.options.id) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${teamId}/${args.options.primary ? 'primaryChannel' : `channels/${args.options.id}`}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res = await request.get<Channel>(requestOptions);
        await logger.log(res);
      }
      else {
        const channel = await teams.getChannelByDisplayName(teamId, args.options.name!);
        await logger.log(channel);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChannelGetCommand();