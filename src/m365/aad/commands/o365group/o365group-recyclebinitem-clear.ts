import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class AadO365GroupRecycleBinItemClearCommand extends GraphItemsListCommand<DirectoryObject> {
  public get name(): string {
    return commands.O365GROUP_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Clears all O365 Groups from recycle bin.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const clearO365GroupRecycleBinItems: () => void = (): void => {
      this.processRecycleBinItemsClear(logger).then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      clearO365GroupRecycleBinItems();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear all O365 Groups from recycle bin ?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          clearO365GroupRecycleBinItems();
        }
      });
    }
  }

  public processRecycleBinItemsClear(logger: Logger): Promise<any> {
    const filter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const topCount: string = '&$top=100';
    const endpoint: string = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${topCount}`;

    return this
      .getAllItems(endpoint, logger, true)
      .then((): Promise<any> => {
        if (this.items.length === 0) {
          return Promise.resolve();
        }

        const deletePromises: Promise<any>[] = [];
        // Logic to delete a group from recycle bin items.
        this.items.forEach(grp => {
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadO365GroupRecycleBinItemClearCommand();