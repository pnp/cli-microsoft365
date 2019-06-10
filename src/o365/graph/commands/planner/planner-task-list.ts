import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import { Task } from './Task';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
      .getAllItems(`${auth.service.resource}/v1.0/me/planner/tasks`, cmd, true)
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
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${ chalk.blue(commands.LOGIN)} command.

  Remarks:

    To list Planner tasks for the currently logged in user, you have to first
    log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    List tasks for the currently logged in user
      ${chalk.grey(config.delimiter)} ${this.name}`
);
  }
}

module.exports = new GraphPlannerTaskListCommand();