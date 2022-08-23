import { Logger } from '../../../../cli';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ToDoList } from '../../ToDoList';

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

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const items: any = await odata.getAllItems<ToDoList>(`${this.resource}/v1.0/me/todo/lists`);
      logger.log(items);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TodoListListCommand();