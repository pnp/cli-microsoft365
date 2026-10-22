import { Channel } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
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
    .optional(),
  teamName: z.string()
    .optional(),
  id: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value is not a valid Teams channel id.'
    })
    .optional()
    .alias('i'),
  name: z.string()
    .optional(),
  primary: z.boolean()
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_GET;
  }

  public get description(): string {
    return 'Gets information about the specific Microsoft Teams team channel';
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
      .refine(opts => [opts.id, opts.name, opts.primary].filter(x => x !== undefined).length === 1, {
        message: 'Specify either id, name, or primary, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'name', 'primary']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId = args.options.teamId || await teams.getTeamIdByDisplayName(args.options.teamName!);

      let channel: Channel;
      if (args.options.primary || args.options.id) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${teamId}/${args.options.primary ? 'primaryChannel' : `channels/${args.options.id}`}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        channel = await request.get<Channel>(requestOptions);
      }
      else {
        channel = await teams.getChannelByDisplayName(teamId, args.options.name!);
      }

      await logger.log(channel);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChannelGetCommand();