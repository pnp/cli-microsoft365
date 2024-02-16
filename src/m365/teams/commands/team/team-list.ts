import { Team } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  joined?: boolean;
  associated?: boolean;
  userId?: string;
  userName?: string;
}

class TeamsTeamListCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Teams in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isArchived', 'description'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        joined: !!args.options.joined,
        associated: !!args.options.associated,
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-j, --joined'
      },
      {
        option: '-a, --associated'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid UPN for userName.`;
        }

        if ((args.options.userId || args.options.userName) && !args.options.joined && !args.options.associated) {
          return 'You must specify either joined or associated when specifying userId or userName.';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['joined', 'associated'],
        runsWhen: (args: CommandArgs) => !!args.options.joined || !!args.options.associated
      },
      {
        options: ['userId', 'userName'],
        runsWhen: (args: CommandArgs) => typeof args.options.userId !== 'undefined' || typeof args.options.userName !== 'undefined'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('userId', 'userName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      if (!args.options.joined && !args.options.associated) {
        await logger.logToStderr(`Retrieving Microsoft Teams in the tenant...`);
      }
      else {
        const user = args.options.userId || args.options.userName || 'me';
        await logger.logToStderr(`Retrieving Microsoft Teams ${args.options.joined ? 'joined by' : 'associated with'} ${user}...`);
      }
    }

    try {
      let endpoint = `${this.resource}/v1.0`;
      if (args.options.joined || args.options.associated) {
        if (!args.options.userId && !args.options.userName && accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
          throw `You must specify either userId or userName when using application only permissions and specifying the ${args.options.joined ? 'joined' : 'associated'} option`;
        }

        endpoint += args.options.userId || args.options.userName ? `/users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}` : '/me';
        endpoint += args.options.joined ? '/joinedTeams' : '/teamwork/associatedTeams';
        endpoint += '?$select=id';
      }
      else {
        // Get all team groups within the tenant
        endpoint += `/groups?$select=id&$filter=resourceProvisioningOptions/Any(x:x eq 'Team')`;
      }

      const groupResponse = await odata.getAllItems<{ id: string }>(endpoint);
      const groupIds = groupResponse.map(g => g.id);

      if (this.verbose) {
        await logger.logToStderr(`Retrieved ${groupIds.length} Microsoft Teams, getting additional information...`);
      }

      let teams = await this.getAllTeams(groupIds);
      // Sort teams by display name
      teams = teams.sort((x: Team, y: Team) => x.displayName!.localeCompare(y.displayName!));
      await logger.log(teams);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAllTeams(groupIds: string[]): Promise<Team[]> {
    const groupBatches: string[][] = [];
    for (let i = 0; groupIds.length > i; i += 20) {
      groupBatches.push(groupIds.slice(i, i + 20));
    }

    const promises = groupBatches.map(g => this.getTeamsBatch(g));
    const teams = await Promise.all(promises);
    const result = teams.reduce((prev: Team[], val: Team[]) => prev.concat(val), []);

    return result;
  }

  private async getTeamsBatch(groupIds: string[]): Promise<Team[]> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/$batch`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        requests: groupIds.map((id, index) => ({
          id: index.toString(),
          method: 'GET',
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          url: `/teams/${id}`
        }))
      }
    };

    const response = await request.post<{ responses: { status: number; body: Team }[] }>(requestOptions);

    // Throw error if any of the requests failed
    for (const item of response.responses) {
      if (item.status !== 200) {
        throw item.body;
      }
    }

    return response.responses.map(r => r.body);
  }
}

export default new TeamsTeamListCommand();