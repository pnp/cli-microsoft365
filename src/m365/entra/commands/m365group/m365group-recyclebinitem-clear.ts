import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupRecycleBinItemClearCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Clears all M365 Groups from recycle bin.';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearM365GroupRecycleBinItems = async (): Promise<void> => {
      try {
        await this.processRecycleBinItemsClear();
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearM365GroupRecycleBinItems();
    }
    else {
      const response = await cli.promptForConfirmation({ message: `Are you sure you want to clear all M365 Groups from recycle bin?` });

      if (response) {
        await clearM365GroupRecycleBinItems();
      }
    }
  }

  public async processRecycleBinItemsClear(): Promise<void> {
    const filter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const topCount: string = '&$top=100';
    const endpoint: string = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${topCount}`;

    const recycleBinItems = await odata.getAllItems<DirectoryObject>(endpoint);
    if (recycleBinItems.length === 0) {
      return;
    }

    const deletePromises: Promise<any>[] = [];
    // Logic to delete a group from recycle bin items.
    recycleBinItems.forEach(grp => {
      deletePromises.push(
        request.delete({
          url: `${this.resource}/v1.0/directory/deletedItems/${grp.id}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        })
      );
    });
    await Promise.all(deletePromises);
  }
}

export default new EntraM365GroupRecycleBinItemClearCommand();