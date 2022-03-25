import { Logger } from '../../../../cli';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
}

class SearchExternalConnectionListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.EXTERNALCONNECTION_LIST;
  }

  public get description(): string {
    return 'Adds a new External Connection for Microsoft Search';
  }
  
  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let urlStub = 'v1.0/external/connections';
    if (args.options.id !== null){
      urlStub += `?$filter=id eq ${args.options.id}`;
    }
    
    this
      .getAllItems(`${this.resource}/${urlStub}`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      });
  }

  public defaultProperties(): string[] | undefined { 
    return ['name', 'description', 'id']; 
  } 
}

module.exports = new SearchExternalConnectionListCommand();