import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.coerce.number(),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageMessageRemoveCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a Viva Engage message';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeMessage(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Viva Engage message ${args.options.id}?` });

      if (result) {
        await this.removeMessage(args.options);
      }
    }
  }

  private async removeMessage(options: Options): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1/messages/${options.id}.json`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageMessageRemoveCommand();
