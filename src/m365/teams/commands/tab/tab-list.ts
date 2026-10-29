import { TeamsTab } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .alias('i'),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value must be a valid Teams channel ID.'
    })
    .alias('c')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTabListCommand extends GraphCommand {
  public get name(): string {
    return commands.TAB_LIST;
  }

  public get description(): string {
    return 'Lists tabs in the specified Microsoft Teams channel';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'teamsAppTabId'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${formatting.encodeQueryParameter(args.options.channelId)}/tabs?$expand=teamsApp`;

    try {
      const items = await odata.getAllItems<TeamsTab>(endpoint);

      if (args.options.output !== 'json') {
        items.forEach(i => {
          (i as any).teamsAppTabId = i.teamsApp!.id;
        });
      }

      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsTabListCommand();