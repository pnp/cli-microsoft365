import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  chatId: z.string().refine(val => validation.isValidTeamsChatId(val), {
    message: 'The value of the option chatId must be a valid Teams ChatId.'
  }).alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChatMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists all members from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['userId', 'displayName', 'email'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;

    try {
      const items = await odata.getAllItems(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatMemberListCommand();