import { Channel } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional()
    .alias('i'),
  teamName: z.string()
    .optional(),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value is not a valid Teams channel id.'
    })
    .optional()
    .alias('c'),
  channelName: z.string()
    .optional(),
  userIds: z.string()
    .optional(),
  userDisplayNames: z.string()
    .optional(),
  owner: z.boolean()
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a specified member in the specified Microsoft Teams private or shared team channel';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.teamId, opts.teamName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either teamId or teamName, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['teamId', 'teamName']
        }
      })
      .refine(opts => [opts.channelId, opts.channelName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either channelId or channelName, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['channelId', 'channelName']
        }
      })
      .refine(opts => [opts.userIds, opts.userDisplayNames].filter(x => x !== undefined).length === 1, {
        message: 'Specify either userIds or userDisplayNames, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['userIds', 'userDisplayNames']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      const channelId: string = await this.getChannelId(teamId, args);
      const userIds: string[] = await this.getUserId(args);
      const endpoint: string = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(channelId)}/members`;
      const roles: string[] = args.options.owner ? ["owner"] : [];
      const tasks: Promise<void>[] = [];

      for (const userId of userIds) {
        tasks.push(this.addUser(userId, endpoint, roles));
      }

      await Promise.all(tasks);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addUser(userId: string, endpoint: string, roles: string[]): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: endpoint,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'roles': roles,
        'user@odata.bind': `${this.resource}/v1.0/users('${userId}')`
      }
    };

    return request.post(requestOptions);
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    return teams.getTeamIdByDisplayName(args.options.teamName!);
  }

  private async getChannelId(teamId: string, args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return args.options.channelId;
    }

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw `The specified channel '${args.options.channelName}' does not exist in the Microsoft Teams team with ID '${teamId}'`;
    }

    if (channelItem.membershipType !== "private") {
      throw `The specified channel is not a private channel`;
    }

    return channelItem.id!;
  }

  private async getUserId(args: CommandArgs): Promise<string[]> {
    if (args.options.userIds) {
      return args.options.userIds.split(',').map(u => u.trim());
    }

    const tasks: Promise<string>[] = [];
    const userDisplayNames: any | undefined = args.options.userDisplayNames && args.options.userDisplayNames.split(',').map(u => u.trim());

    for (const userName of userDisplayNames) {
      tasks.push(this.getSingleUser(userName));
    }

    return Promise.all(tasks);
  }

  private async getSingleUser(userDisplayName: string): Promise<string> {
    const userRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter(userDisplayName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: any[] }>(userRequestOptions);
    const userItem: any | undefined = response.value[0];

    if (!userItem) {
      throw `The specified user '${userDisplayName}' does not exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<any>(`Multiple users with display name '${userDisplayName}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return userItem.id;
  }
}

export default new TeamsChannelMemberAddCommand();