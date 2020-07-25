import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { Task } from '../../Task';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/me/planner/tasks`, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(t => {
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
          cmd.log(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new GraphPlannerTaskListCommand();