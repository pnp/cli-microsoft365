import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { teams } from '../../../../utils/teams.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => !val || validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  teamName: z.string()
    .optional(),
  channelId: z.string()
    .refine(val => !val || validation.isValidTeamsChannelId(val), {
      message: 'The value is not a valid Teams ChannelId.'
    })
    .optional(),
  channelName: z.string()
    .optional(),
  id: z.string()
    .alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMessageRestoreCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.MESSAGE_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted message from a channel in a Microsoft Teams team';
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
      if (this.verbose) {
        await logger.logToStderr(`Restoring deleted message '${args.options.id}' from channel '${args.options.channelId || args.options.channelName}' in the Microsoft Teams team '${args.options.teamId || args.options.teamName}'.`);
      }

      const teamId: string = await this.getTeamId(args.options, logger);
      const channelId: string = await this.getChannelId(args.options, teamId, logger);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${teamId}/channels/${channelId}/messages/${args.options.id}/undoSoftDelete`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(options: Options, logger: Logger): Promise<string> {
    if (options.teamId) {
      return options.teamId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Getting the Team ID.`);
    }

    const groupId = await teams.getTeamIdByDisplayName(options.teamName!);

    return groupId;
  }

  private async getChannelId(options: Options, teamId: string, logger: Logger): Promise<string> {
    if (options.channelId) {
      return options.channelId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Getting the channel ID.`);
    }

    const channelId = await teams.getChannelIdByDisplayName(teamId, options.channelName!);
    return channelId;
  }
}

export default new TeamsMessageRestoreCommand();