import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  name: string;
  description?: string;
  type: string;
  owner: string;
}

class TeamsChannelAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_ADD;
  }

  public get description(): string {
    return 'Adds a channel to the specified Microsoft Teams team';
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
        description: typeof args.options.description !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        type: args.options.type || 'standard',
        owner: typeof args.options.owner !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--type [type]',
        autocomplete: ['standard', 'private', 'shared']
      },
      {
        option: '--owner [owner]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.type && ['standard', 'private', 'shared'].indexOf(args.options.type) === -1) {
          return `${args.options.type} is not a valid type value. Allowed values standard|private|shared.`;
        }

        if ((args.options.type === 'private' || args.options.type === 'shared') && !args.options.owner) {
          return `Specify owner when creating a ${args.options.type} channel.`;
        }

        if ((args.options.type !== 'private' && args.options.type !== 'shared') && args.options.owner) {
          return `Specify owner only when creating a private or shared channel.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      const res: any = await this.createChannel(args, teamId);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    return await teams.getTeamIdByDisplayName(args.options.teamName!);
  }

  private async createChannel(args: CommandArgs, teamId: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${teamId}/channels`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      data: {
        membershipType: args.options.type || 'standard',
        displayName: args.options.name
      },
      responseType: 'json'
    };

    if (args.options.type === 'private' || args.options.type === 'shared') {
      // Private and Shared channels must have at least 1 owner
      requestOptions.data.members = [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `${this.resource}/v1.0/users('${args.options.owner}')`,
          roles: ['owner']
        }
      ];
    }

    return request.post(requestOptions);
  }
}

export default new TeamsChannelAddCommand();