import { Channel, ConversationMember } from '@microsoft/microsoft-graph-types';
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
    .optional(),
  teamName: z.string()
    .optional(),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value is not a valid Teams channel id.'
    })
    .optional(),
  channelName: z.string()
    .optional(),
  userName: z.string()
    .optional(),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  id: z.string()
    .optional(),
  role: z.enum(['owner', 'member'])
    .alias('r')
});

interface ExtendedConversationMember extends ConversationMember {
  userId?: string;
  email?: string;
}

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelMemberSetCommand extends GraphCommand {
  private teamId: string = '';
  private channelId: string = '';

  public get name(): string {
    return commands.CHANNEL_MEMBER_SET;
  }

  public get description(): string {
    return 'Updates the role of the specified member in the specified Microsoft Teams private or shared team channel';
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
      .refine(opts => [opts.userName, opts.userId, opts.id].filter(x => x !== undefined).length === 1, {
        message: 'Specify either userName, userId, or id, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['userName', 'userId', 'id']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.teamId = await this.getTeamId(args);
      this.channelId = await this.getChannelId(args);
      const memberId: string = await this.getMemberId(args);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members/${memberId}`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: [args.options.role]
        }
      };

      const member = await request.patch(requestOptions);
      await logger.log(member);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    return teams.getTeamIdByDisplayName(args.options.teamName!);
  }

  private async getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return args.options.channelId;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(requestOptions);

    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw 'The specified channel does not exist in the Microsoft Teams team';
    }

    if (channelItem.membershipType !== "private") {
      throw 'The specified channel is not a private channel';
    }

    return channelItem.id!;
  }

  private async getMemberId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: ExtendedConversationMember[] }>(requestOptions);

    const conversationMembers = response.value.filter(x =>
      args.options.userId && x.userId?.toLocaleLowerCase() === args.options.userId.toLocaleLowerCase() ||
      args.options.userName && x.email?.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase()
    );

    const conversationMember: ConversationMember | undefined = conversationMembers[0];

    if (!conversationMember) {
      throw 'The specified member does not exist in the Microsoft Teams channel';
    }

    if (conversationMembers.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', conversationMembers);
      const result = await cli.handleMultipleResultsFound<any>(`Multiple Microsoft Teams channel members with name ${args.options.userName} found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return conversationMember.id!;
  }
}

export default new TeamsChannelMemberSetCommand();