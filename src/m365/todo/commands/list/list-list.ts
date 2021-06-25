import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { ToDoList } from '../../ToDoList';

interface CommandArgs {
  options: GlobalOptions;
}

class TodoListListCommand extends GraphItemsListCommand<ToDoList> {
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
    this
      .getAllItems(`${this.resource}/v1.0/me/todo/lists`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TodoListListCommand();