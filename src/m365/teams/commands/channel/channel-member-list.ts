import { Channel, ConversationMember } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
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
  role: z.enum(['owner', 'member', 'guest'])
    .optional()
    .alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelMemberListCommand extends GraphCommand {
  private teamId: string = '';

  public get name(): string {
    return commands.CHANNEL_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists members of the specified Microsoft Teams team channel';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'roles', 'displayName', 'userId', 'email'];
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.teamId = await this.getTeamId(args);
      const channelId: string = await this.getChannelId(args);
      const endpoint = `${this.resource}/v1.0/teams/${this.teamId}/channels/${channelId}/members`;
      let memberships = await odata.getAllItems<ConversationMember>(endpoint);
      if (args.options.role) {
        if (args.options.role === 'member') {
          // Members have no role value
          memberships = memberships.filter(i => i.roles!.length === 0);
        }
        else {
          memberships = memberships.filter(i => i.roles!.indexOf(args.options.role!) !== -1);
        }
      }

      await logger.log(memberships);
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

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw 'The specified channel does not exist in the Microsoft Teams team';
    }

    return channelItem.id!;
  }
}

export default new TeamsChannelMemberListCommand();