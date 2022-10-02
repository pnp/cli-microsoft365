import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class AadO365GroupRecycleBinItemClearCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Clears all O365 Groups from recycle bin.';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: typeof args.options.confirm !== 'undefined'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '--confirm'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearO365GroupRecycleBinItems: () => Promise<void> = async (): Promise<void> => {
      try {
        await this.processRecycleBinItemsClear();
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await clearO365GroupRecycleBinItems();
    }
    else {
      const response = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear all O365 Groups from recycle bin ?`
      });

      if (response.continue){
        await clearO365GroupRecycleBinItems();
      }
    }
  }

  public processRecycleBinItemsClear(): Promise<any> {
    const filter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const topCount: string = '&$top=100';
    const endpoint: string = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${topCount}`;

    return odata
      .getAllItems<DirectoryObject>(endpoint)
      .then((recycleBinItems): Promise<any> => {
        if (recycleBinItems.length === 0) {
          return Promise.resolve();
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
        return Promise.all(deletePromises);
      });
  }
}

module.exports = new AadO365GroupRecycleBinItemClearCommand();