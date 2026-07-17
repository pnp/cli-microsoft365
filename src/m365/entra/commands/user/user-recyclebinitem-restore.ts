import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserRecycleBinItemRestoreCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores a user from the tenant recycle bin';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Restoring user with id ${args.options.id} from the recycle bin.`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}/restore`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      const user = await request.post<User>(requestOptions);
      await logger.log(user);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserRecycleBinItemRestoreCommand();