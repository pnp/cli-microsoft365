import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
}

class EntraM365GroupRecycleBinItemClearCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Clears all M365 Groups from recycle bin.';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_RECYCLEBINITEM_CLEAR];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: typeof args.options.force !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --force'
      }
    );
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