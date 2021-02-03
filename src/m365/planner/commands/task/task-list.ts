import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Task } from '../../Task';

interface CommandArgs {
  options: GlobalOptions;
}

class GraphPlannerTaskListCommand extends GraphItemsListCommand<Task> {
  public get name(): string {
    return commands.PLANNER_TASK_LIST;
  }

  public get description(): string {
    return 'Lists Planner tasks for the currently logged in user';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/me/planner/tasks`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new GraphPlannerTaskListCommand();