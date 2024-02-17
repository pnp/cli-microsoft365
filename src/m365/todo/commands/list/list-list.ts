import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import ToDoCommand from '../../../base/ToDoCommand.js';
import commands from '../../commands.js';
import { ToDoList } from '../../ToDoList.js';

class TodoListListCommand extends ToDoCommand {
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
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TodoListListCommand();