import { Channel } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
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
  type: z.enum(['standard', 'private', 'shared'])
    .optional()
    .transform(v => v?.toLowerCase())
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_LIST;
  }

  public get description(): string {
    return 'Lists channels in the specified Microsoft Teams team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
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
      });
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    return teams.getTeamIdByDisplayName(args.options.teamName!);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId: string = await this.getTeamId(args);
      let endpoint: string = `${this.resource}/v1.0/teams/${teamId}/channels`;

      if (args.options.type) {
        endpoint += `?$filter=membershipType eq '${args.options.type}'`;
      }

      const items = await odata.getAllItems<Channel>(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChannelListCommand();