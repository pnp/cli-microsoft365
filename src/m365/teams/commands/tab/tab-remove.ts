import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .alias('t'),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value must be a valid Teams channel ID.'
    })
    .alias('c'),
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTabRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TAB_REMOVE;
  }

  public get description(): string {
    return "Removes a tab from the specified channel";
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeTab = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}/channels/${args.options.channelId}/tabs/${formatting.encodeQueryParameter(args.options.id)}`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };
    if (args.options.force) {
      await removeTab();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the tab with id ${args.options.id} from channel ${args.options.channelId} in team ${args.options.teamId}?` });

      if (result) {
        await removeTab();
      }
    }
  }
}

export default new TeamsTabRemoveCommand();