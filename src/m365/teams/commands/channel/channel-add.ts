import { Team } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';

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

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const teamRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/me/joinedTeams`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Team[] }>(teamRequestOptions);

    const matchingTeams: string[] = response.value
      .filter(team => team.displayName! === args.options.teamName)
      .map(team => team.id!);

    if (matchingTeams.length < 1) {
      throw `The specified team does not exist in the Microsoft Teams`;
    }

    if (matchingTeams.length > 1) {
      throw `Multiple Microsoft Teams teams with name ${args.options.teamName} found: ${matchingTeams.join(', ')}`;
    }

    return matchingTeams[0];
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
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${args.options.owner}')`,
          roles: ['owner']
        }
      ];
    }

    return request.post(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      const res: any = await this.createChannel(args, teamId);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsChannelAddCommand();