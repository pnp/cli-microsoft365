import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';
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
  name: z.string()
    .alias('n'),
  description: z.string()
    .optional()
    .alias('d'),
  type: z.enum(['standard', 'private', 'shared'])
    .optional(),
  owner: z.string()
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_ADD;
  }

  public get description(): string {
    return 'Adds a channel to the specified Microsoft Teams team';
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
      .refine(opts => !((opts.type === 'private' || opts.type === 'shared') && !opts.owner), {
        message: 'Specify owner when creating a private or shared channel.',
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => !((opts.type !== 'private' && opts.type !== 'shared') && opts.owner), {
        message: 'Specify owner only when creating a private or shared channel.',
        params: {
          customCode: 'required'
        }
      });
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