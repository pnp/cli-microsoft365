import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';

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
    .alias('i'),
  force: z.boolean()
    .optional()
    .alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMessageRemoveCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a message from a channel in a Microsoft Teams team';
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
    const removeTeamMessage = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing message '${args.options.id}' from channel '${args.options.channelId || args.options.channelName}' in team '${args.options.teamId || args.options.teamName}'.`);
        }

        const teamId: string = args.options.teamId || await teams.getTeamIdByDisplayName(args.options.teamName!);
        const channelId: string = args.options.channelId || await teams.getChannelIdByDisplayName(teamId, args.options.channelName!);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${teamId}/channels/${formatting.encodeQueryParameter(channelId)}/messages/${args.options.id}/softDelete`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        if (err.error?.error?.code === 'NotFound') {
          this.handleError('The message was not found in the Teams channel.');
        }
        else {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await removeTeamMessage();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove this message?` });

      if (result) {
        await removeTeamMessage();
      }
    }
  }
}

export default new TeamsMessageRemoveCommand();