import * as chalk from 'chalk';
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
    return `${commands.PLANNER_TASK_LIST}`;
  }

  public get description(): string {
    return 'Lists Planner tasks for the currently logged in user';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/me/planner/tasks`, logger, true)
      .then((): void => {
        if (args.options.output === 'json') {
          logger.log(this.items);
        }
        else {
          logger.log(this.items.map(t => {
            const task: any = {
              id: t.id,
              title: t.title,
              startDateTime: t.startDateTime,
              dueDateTime: t.dueDateTime,
              completedDateTime: t.completedDateTime
            };
            return task;
          }));
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new GraphPlannerTaskListCommand();