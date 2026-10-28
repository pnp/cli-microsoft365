import { Team } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  joined: z.boolean().optional().alias('j'),
  associated: z.boolean().optional().alias('a'),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value must be a valid user principal name.'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.joined, opts.associated].filter(o => o !== undefined).length <= 1, {
        message: `Specify either 'joined' or 'associated', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['joined', 'associated']
        }
      })
      .refine(opts => [opts.userId, opts.userName].filter(o => o !== undefined).length <= 1, {
        message: `Specify either 'userId' or 'userName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      })
      .refine(opts => !(opts.userId || opts.userName) || opts.joined || opts.associated, {
        message: `Specify 'joined' or 'associated' when specifying 'userId' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['joined', 'associated']
        }
      });
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