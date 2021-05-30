import { Cli,Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Group } from './Group';
import request from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class AadO365GroupRecycleBinItemClearCommand extends GraphItemsListCommand<Group> {
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

    if (args.options.confirm) {
      this.ClearO365GroupRecycleBinItems(logger,args,cb).then(()=>{
        cb();
      });
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
          this.ClearO365GroupRecycleBinItems(logger,args,cb).then(()=>{
            cb();
          });
        }
      });
    }
 
  }

  public ClearO365GroupRecycleBinItems(logger: Logger, args: CommandArgs,cb: () => void):Promise<void>{
    const filter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const topCount: string = '&$top=100';

    const endpoint: string = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${topCount}`;

    return this
      .getAllItems(endpoint, logger, true)
      .then(():Promise<void> => {

        if(this.items.length === 0){
          return Promise.resolve();
        }

        const deletePromises:any[] = [];
        // Logic to delete a group from recycle bin items.
        this.items.forEach(grp =>{
          deletePromises.push(
            request.delete({
              url: `${this.resource}/v1.0/directory/deletedItems/${grp.id}`,
              headers: {
                'accept': 'application/json;odata.metadata=none'
              }
            })
          );
        });
        
        return Promise.all(deletePromises).then(_=>Promise.resolve(),(err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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