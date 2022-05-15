import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ToDoList } from '../../ToDoList';

interface CommandArgs {
  options: GlobalOptions;
}

class TodoListListCommand extends GraphCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Returns a list of Microsoft To Do task lists';
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'id'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    odata
      .getAllItems<ToDoList>(`${this.resource}/v1.0/me/todo/lists`)
      .then((items): void => {
        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TodoListListCommand();