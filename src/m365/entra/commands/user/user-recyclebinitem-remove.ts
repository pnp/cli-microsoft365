import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserRecycleBinItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes a user from the recycle bin in the current tenant';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearRecycleBinItem: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Permanently deleting user with id ${args.options.id} from Microsoft Entra ID`);
      }

      try {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}`,
          headers: {}
        };
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearRecycleBinItem();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to permanently delete the user with id ${args.options.id}?` });

      if (result) {
        await clearRecycleBinItem();
      }
    }
  }
}

export default new EntraUserRecycleBinItemRemoveCommand();
