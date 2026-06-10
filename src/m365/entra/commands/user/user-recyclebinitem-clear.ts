import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserRecycleBinItemClearCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Removes all users from the tenant recycle bin';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearRecycleBinUsers = async (): Promise<void> => {
      try {
        const users = await odata.getAllItems<User>(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.user?$select=id`);
        if (this.verbose) {
          await logger.logToStderr(`Amount of users to permanently delete: ${users.length}`);
        }
        const batchRequests = users.map((user, index) => {
          return {
            id: index,
            method: 'DELETE',
            url: `/directory/deletedItems/${user.id}`
          };
        });
        for (let i = 0; i < batchRequests.length; i += 20) {
          const batchRequestChunk = batchRequests.slice(i, i + 20);
          if (this.verbose) {
            await logger.logToStderr(`Deleting users: ${i + batchRequestChunk.length}/${users.length}`);
          }

          const requestOptions: CliRequestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
              accept: 'application/json',
              'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
              requests: batchRequestChunk
            }
          };
          await request.post(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearRecycleBinUsers();
    }
    else {
      const result = await cli.promptForConfirmation({ message: 'Are you sure you want to permanently delete all deleted users?' });

      if (result) {
        await clearRecycleBinUsers();
      }
    }
  }
}

export default new EntraUserRecycleBinItemClearCommand();
